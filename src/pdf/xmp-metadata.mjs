import crypto from "node:crypto";

const XMP_MAX_BYTES = 1024 * 1024;

const NAMESPACES = Object.freeze({
  dc: "http://purl.org/dc/elements/1.1/",
  pdf: "http://ns.adobe.com/pdf/1.3/",
  rdf: "http://www.w3.org/1999/02/22-rdf-syntax-ns#",
  xmp: "http://ns.adobe.com/xap/1.0/",
  xml: "http://www.w3.org/XML/1998/namespace",
});

const PROPERTY_SHAPES = Object.freeze({
  author: { namespace: NAMESPACES.dc, localName: "creator", container: "Seq" },
  title: { namespace: NAMESPACES.dc, localName: "title", container: "Alt" },
  subject: { namespace: NAMESPACES.dc, localName: "description", container: "Alt" },
  keywords: { namespace: NAMESPACES.pdf, localName: "Keywords", attribute: true },
  creator: { namespace: NAMESPACES.xmp, localName: "CreatorTool", attribute: true },
  producer: { namespace: NAMESPACES.pdf, localName: "Producer", attribute: true },
  creationDate: { namespace: NAMESPACES.xmp, localName: "CreateDate", attribute: true },
  modificationDate: { namespace: NAMESPACES.xmp, localName: "ModifyDate", attribute: true },
});

function sha256(value) {
  return crypto.createHash("sha256").update(value).digest("hex");
}

function validateXmlCharacters(value) {
  for (const character of value) {
    const codePoint = character.codePointAt(0);
    if (codePoint === 0x9 || codePoint === 0xa || codePoint === 0xd) continue;
    if ((codePoint >= 0x20 && codePoint <= 0xd7ff)
      || (codePoint >= 0xe000 && codePoint <= 0xfffd)
      || (codePoint >= 0x10000 && codePoint <= 0x10ffff)) continue;
    throw new Error(`invalid XML character U+${codePoint.toString(16).toUpperCase().padStart(4, "0")}`);
  }
  return value;
}

function decodeXmlText(value) {
  const unsupportedEntity = /&(?!amp;|lt;|gt;|quot;|apos;|#\d+;|#x[0-9a-fA-F]+;)/u.exec(value);
  if (unsupportedEntity) throw new Error(`unsupported XML entity reference near offset ${unsupportedEntity.index}`);
  const decoded = value.replace(/&(?:amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);/gu, (entity) => {
    if (entity === "&amp;") return "&";
    if (entity === "&lt;") return "<";
    if (entity === "&gt;") return ">";
    if (entity === "&quot;") return "\"";
    if (entity === "&apos;") return "'";
    const hexadecimal = entity.startsWith("&#x");
    const codePoint = Number.parseInt(entity.slice(hexadecimal ? 3 : 2, -1), hexadecimal ? 16 : 10);
    if (!Number.isSafeInteger(codePoint) || codePoint < 1 || codePoint > 0x10ffff || (codePoint >= 0xd800 && codePoint <= 0xdfff)) {
      throw new Error(`invalid XML character reference: ${entity}`);
    }
    return String.fromCodePoint(codePoint);
  });
  return validateXmlCharacters(decoded);
}

function encodeXmlText(value) {
  validateXmlCharacters(value);
  return value.replace(/&/gu, "&amp;").replace(/</gu, "&lt;").replace(/>/gu, "&gt;");
}

function encodeXmlAttribute(value, quote) {
  return encodeXmlText(value)
    .replace(quote === "\"" ? /"/gu : /'/gu, quote === "\"" ? "&quot;" : "&apos;")
    .replace(/\t/gu, "&#x9;")
    .replace(/\n/gu, "&#xA;")
    .replace(/\r/gu, "&#xD;");
}

function expandedName(qname, namespaces, attribute = false) {
  const separator = qname.indexOf(":");
  if (separator === -1) return { namespace: attribute ? "" : namespaces.get("") || "", localName: qname };
  if (qname.indexOf(":", separator + 1) !== -1) throw new Error(`invalid qualified XML name: ${qname}`);
  const prefix = qname.slice(0, separator);
  const localName = qname.slice(separator + 1);
  if (!prefix || !localName || !namespaces.has(prefix)) throw new Error(`undeclared XML namespace prefix: ${prefix || "(empty)"}`);
  return { namespace: namespaces.get(prefix), localName };
}

function findMarkupEnd(xml, start) {
  let quote = "";
  for (let index = start; index < xml.length; index += 1) {
    const character = xml[index];
    if (quote) {
      if (character === quote) quote = "";
    } else if (character === "\"" || character === "'") quote = character;
    else if (character === ">") return index;
  }
  throw new Error("unterminated XML start tag");
}

function parseAttributes(source, sourceOffset) {
  const attributes = [];
  const names = new Set();
  let index = 0;
  while (index < source.length) {
    while (/\s/u.test(source[index] || "")) index += 1;
    if (index >= source.length) break;
    const nameMatch = /^[A-Za-z_][A-Za-z0-9_.:-]*/u.exec(source.slice(index));
    if (!nameMatch) throw new Error("invalid XML attribute name");
    const name = nameMatch[0];
    if (names.has(name)) throw new Error(`duplicate XML attribute: ${name}`);
    names.add(name);
    index += name.length;
    while (/\s/u.test(source[index] || "")) index += 1;
    if (source[index] !== "=") throw new Error(`XML attribute ${name} lacks '='`);
    index += 1;
    while (/\s/u.test(source[index] || "")) index += 1;
    const quote = source[index];
    if (quote !== "\"" && quote !== "'") throw new Error(`XML attribute ${name} is not quoted`);
    index += 1;
    const valueStart = index;
    const valueEnd = source.indexOf(quote, valueStart);
    if (valueEnd === -1) throw new Error(`XML attribute ${name} is unterminated`);
    const rawValue = source.slice(valueStart, valueEnd);
    if (/[\t\r\n]/u.test(rawValue)) throw new Error(`XML attribute ${name} contains non-canonical literal whitespace`);
    attributes.push({
      name,
      value: decodeXmlText(rawValue),
      quote,
      valueStart: sourceOffset + valueStart,
      valueEnd: sourceOffset + valueEnd,
    });
    index = valueEnd + 1;
  }
  return attributes;
}

function parseOpenTag(xml, start, inheritedNamespaces) {
  const end = findMarkupEnd(xml, start + 1);
  let body = xml.slice(start + 1, end);
  const selfClosing = /\/\s*$/u.test(body);
  if (selfClosing) body = body.replace(/\/\s*$/u, "");
  const nameMatch = /^\s*([A-Za-z_][A-Za-z0-9_.:-]*)/u.exec(body);
  if (!nameMatch) throw new Error("invalid XML element name");
  const qname = nameMatch[1];
  const attributeOffset = start + 1 + nameMatch[0].length;
  const attributes = parseAttributes(body.slice(nameMatch[0].length), attributeOffset);
  const namespaces = new Map(inheritedNamespaces);
  for (const attribute of attributes) {
    if (attribute.name === "xmlns") namespaces.set("", attribute.value);
    else if (attribute.name.startsWith("xmlns:")) namespaces.set(attribute.name.slice(6), attribute.value);
  }
  if (namespaces.get("xml") !== NAMESPACES.xml) throw new Error("the xml namespace prefix must retain its standard binding");
  const name = expandedName(qname, namespaces);
  const expandedAttributes = attributes
    .filter((attribute) => attribute.name !== "xmlns" && !attribute.name.startsWith("xmlns:"))
    .map((attribute) => ({ ...attribute, ...expandedName(attribute.name, namespaces, true) }));
  return { end, selfClosing, qname, name, namespaces, attributes: expandedAttributes };
}

function parseXml(xml) {
  if (xml.includes("\u0000")) throw new Error("XMP contains a NUL byte");
  const documentNode = { children: [], text: [], misc: [], namespaces: new Map([["xml", NAMESPACES.xml]]) };
  const stack = [documentNode];
  let index = 0;
  while (index < xml.length) {
    const markup = xml.indexOf("<", index);
    const textEnd = markup === -1 ? xml.length : markup;
    if (textEnd > index) {
      const value = xml.slice(index, textEnd);
      decodeXmlText(value);
      stack.at(-1).text.push({ start: index, end: textEnd, value });
    }
    if (markup === -1) break;
    if (xml.startsWith("<?", markup)) {
      const end = xml.indexOf("?>", markup + 2);
      if (end === -1) throw new Error("unterminated XML processing instruction");
      stack.at(-1).misc.push({ start: markup, end: end + 2, kind: "processing-instruction" });
      index = end + 2;
      continue;
    }
    if (xml.startsWith("<!--", markup)) {
      const end = xml.indexOf("-->", markup + 4);
      if (end === -1) throw new Error("unterminated XML comment");
      stack.at(-1).misc.push({ start: markup, end: end + 3, kind: "comment" });
      index = end + 3;
      continue;
    }
    if (xml.startsWith("<!", markup)) throw new Error("XMP declarations, DTDs, and CDATA are unsupported");
    if (xml.startsWith("</", markup)) {
      const end = xml.indexOf(">", markup + 2);
      if (end === -1) throw new Error("unterminated XML closing tag");
      const qname = xml.slice(markup + 2, end).trim();
      const node = stack.pop();
      if (!node || node === documentNode || node.qname !== qname) throw new Error(`mismatched XML closing tag: ${qname}`);
      node.closeStart = markup;
      node.end = end + 1;
      index = end + 1;
      continue;
    }
    const parent = stack.at(-1);
    const parsed = parseOpenTag(xml, markup, parent.namespaces);
    const node = {
      qname: parsed.qname,
      namespace: parsed.name.namespace,
      localName: parsed.name.localName,
      attributes: parsed.attributes,
      namespaces: parsed.namespaces,
      children: [],
      text: [],
      misc: [],
      start: markup,
      openEnd: parsed.end + 1,
      closeStart: parsed.selfClosing ? parsed.end : null,
      end: parsed.selfClosing ? parsed.end + 1 : null,
      selfClosing: parsed.selfClosing,
    };
    parent.children.push(node);
    if (!parsed.selfClosing) stack.push(node);
    index = parsed.end + 1;
  }
  if (stack.length !== 1) throw new Error(`unclosed XML element: ${stack.at(-1).qname}`);
  if (documentNode.text.some((segment) => segment.value.trim())) throw new Error("XMP has text outside its document element");
  if (documentNode.children.length !== 1) throw new Error("XMP must contain exactly one document element");
  return documentNode.children[0];
}

function elementChildren(node, namespace, localName) {
  return node.children.filter((child) => child.namespace === namespace && child.localName === localName);
}

function hasMixedContent(node) {
  return node.text.some((segment) => segment.value.trim()) || node.misc.length > 0;
}

function directTextSlot(node, label) {
  if (node.selfClosing || node.children.length || node.misc.length) throw new Error(`${label} must contain direct text only`);
  const raw = node.text.map((segment) => segment.value).join("");
  return { value: decodeXmlText(raw), start: node.openEnd, end: node.closeStart, kind: "text" };
}

function structuredTextSlot(property, shape, label) {
  if (property.attributes.length) throw new Error(`${label} property must not carry attributes`);
  if (property.selfClosing || hasMixedContent(property)) throw new Error(`${label} has unsupported mixed content`);
  const containers = elementChildren(property, NAMESPACES.rdf, shape.container);
  if (containers.length !== 1 || property.children.length !== 1) throw new Error(`${label} must contain exactly one rdf:${shape.container}`);
  const container = containers[0];
  if (container.attributes.length || container.selfClosing || hasMixedContent(container)) {
    throw new Error(`${label} rdf:${shape.container} has unsupported attributes or mixed content`);
  }
  const items = elementChildren(container, NAMESPACES.rdf, "li");
  if (!items.length || items.length !== container.children.length) throw new Error(`${label} rdf:${shape.container} must contain only rdf:li values`);
  if (shape.container === "Seq") {
    if (items.length !== 1) throw new Error(`${label} rdf:Seq contains multiple values`);
    if (items[0].attributes.length) throw new Error(`${label} rdf:Seq item must not carry attributes`);
    return directTextSlot(items[0], `${label} rdf:li`);
  }
  const languages = new Set();
  let defaultItem = null;
  for (const item of items) {
    const languageAttributes = item.attributes.filter((attribute) => attribute.namespace === NAMESPACES.xml && attribute.localName === "lang");
    if (item.attributes.length !== 1 || languageAttributes.length !== 1 || !languageAttributes[0].value) {
      throw new Error(`${label} rdf:Alt items must carry exactly one non-empty xml:lang`);
    }
    const language = languageAttributes[0].value.toLowerCase();
    if (languages.has(language)) throw new Error(`${label} rdf:Alt contains duplicate language ${language}`);
    languages.add(language);
    directTextSlot(item, `${label} rdf:li[xml:lang=${languageAttributes[0].value}]`);
    if (language === "x-default") defaultItem = item;
  }
  if (!defaultItem) throw new Error(`${label} rdf:Alt lacks x-default`);
  return directTextSlot(defaultItem, `${label} rdf:li[xml:lang=x-default]`);
}

function fieldSlot(descriptions, shape, field) {
  const elementMatches = descriptions.flatMap((description) => elementChildren(description, shape.namespace, shape.localName));
  const attributeMatches = descriptions.flatMap((description) => description.attributes
    .filter((attribute) => attribute.namespace === shape.namespace && attribute.localName === shape.localName));
  if (!elementMatches.length && !attributeMatches.length) return null;
  if (elementMatches.length + attributeMatches.length !== 1) throw new Error(`${field} appears more than once`);
  if (attributeMatches.length) {
    if (!shape.attribute) throw new Error(`${field} does not support an attribute-valued representation`);
    const attribute = attributeMatches[0];
    return {
      value: attribute.value,
      start: attribute.valueStart,
      end: attribute.valueEnd,
      kind: "attribute",
      quote: attribute.quote,
    };
  }
  const property = elementMatches[0];
  if (shape.container) return structuredTextSlot(property, shape, field);
  if (property.attributes.length) throw new Error(`${field} property must not carry attributes`);
  return directTextSlot(property, field);
}

export function inspectXmpMetadata(value) {
  const bytes = Buffer.from(value);
  const base = {
    byteLength: bytes.byteLength,
    sha256: sha256(bytes),
    profile: "field-safe-v1",
    values: {},
    mutableFields: [],
    blockedFields: [],
    issues: [],
    slots: {},
  };
  if (bytes.byteLength > XMP_MAX_BYTES) return { ...base, profile: "unsupported", issues: [`xmp-stream-exceeds-${XMP_MAX_BYTES}-bytes`] };
  try {
    const xml = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
    const root = parseXml(xml);
    if (root.namespace !== "adobe:ns:meta/" || root.localName !== "xmpmeta") throw new Error("document element must be x:xmpmeta");
    const rdfRoots = elementChildren(root, NAMESPACES.rdf, "RDF");
    if (rdfRoots.length !== 1 || root.children.length !== 1) throw new Error("x:xmpmeta must contain exactly one rdf:RDF element");
    const descriptions = elementChildren(rdfRoots[0], NAMESPACES.rdf, "Description");
    if (!descriptions.length || descriptions.length !== rdfRoots[0].children.length) throw new Error("rdf:RDF must contain only one or more rdf:Description elements");
    for (const [field, shape] of Object.entries(PROPERTY_SHAPES)) {
      try {
        const slot = fieldSlot(descriptions, shape, field);
        if (!slot) continue;
        base.slots[field] = { start: slot.start, end: slot.end, kind: slot.kind, quote: slot.quote };
        if (slot.value !== "") base.values[field] = slot.value;
        base.mutableFields.push(field);
      } catch (error) {
        base.blockedFields.push({ field, reason: String(error?.message || error) });
      }
    }
    return base;
  } catch (error) {
    return {
      ...base,
      profile: "unsupported",
      values: {},
      mutableFields: [],
      blockedFields: [],
      slots: {},
      issues: [String(error?.message || error)],
    };
  }
}

export function patchXmpMetadata(value, profile, patch) {
  const bytes = Buffer.from(value);
  if (profile.profile !== "field-safe-v1" || profile.issues.length) throw new Error(`XMP metadata profile is unsupported: ${profile.issues.join(", ")}`);
  if (profile.sha256 !== sha256(bytes) || profile.byteLength !== bytes.byteLength) throw new Error("XMP metadata bytes no longer match the inspected profile");
  const replacements = [];
  for (const [field, requested] of Object.entries(patch)) {
    const slot = profile.slots[field];
    if (!slot) {
      const blocked = profile.blockedFields.find((entry) => entry.field === field);
      throw new Error(blocked
        ? `set_metadata cannot synchronize XMP field ${field}: ${blocked.reason}`
        : `set_metadata cannot synchronize XMP field ${field} because the packet does not contain an editable representation`);
    }
    const stringValue = requested ?? "";
    replacements.push({
      start: slot.start,
      end: slot.end,
      value: slot.kind === "attribute" ? encodeXmlAttribute(stringValue, slot.quote) : encodeXmlText(stringValue),
    });
  }
  replacements.sort((left, right) => right.start - left.start);
  let xml = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  for (const replacement of replacements) xml = `${xml.slice(0, replacement.start)}${replacement.value}${xml.slice(replacement.end)}`;
  return Buffer.from(xml, "utf8");
}
