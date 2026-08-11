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
  keywords: { namespace: NAMESPACES.pdf, localName: "Keywords" },
  creator: { namespace: NAMESPACES.xmp, localName: "CreatorTool" },
  producer: { namespace: NAMESPACES.pdf, localName: "Producer" },
  creationDate: { namespace: NAMESPACES.xmp, localName: "CreateDate" },
  modificationDate: { namespace: NAMESPACES.xmp, localName: "ModifyDate" },
});

function sha256(value) {
  return crypto.createHash("sha256").update(value).digest("hex");
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

function parseAttributes(source) {
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
    attributes.push({ name, value: decodeXmlText(source.slice(valueStart, valueEnd)) });
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
  const attributes = parseAttributes(body.slice(nameMatch[0].length));
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
  const documentNode = { children: [], text: [], namespaces: new Map([["xml", NAMESPACES.xml]]) };
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
      if (stack.length > 1) throw new Error("processing instructions inside XMP elements are outside the canonical simple profile");
      const end = xml.indexOf("?>", markup + 2);
      if (end === -1) throw new Error("unterminated XML processing instruction");
      index = end + 2;
      continue;
    }
    if (xml.startsWith("<!--", markup)) {
      if (stack.length > 1) throw new Error("comments inside XMP elements are outside the canonical simple profile");
      const end = xml.indexOf("-->", markup + 4);
      if (end === -1) throw new Error("unterminated XML comment");
      index = end + 3;
      continue;
    }
    if (xml.startsWith("<!", markup)) throw new Error("XMP declarations, DTDs, and CDATA are outside the canonical simple profile");
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

function decodeXmlText(value) {
  return value.replace(/&(?:amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);/gu, (entity) => {
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
  }).replace(/&[^;\s<]+;/gu, (entity) => {
    throw new Error(`unsupported XML entity reference: ${entity}`);
  });
}

function encodeXmlText(value) {
  return value.replace(/&/gu, "&amp;").replace(/</gu, "&lt;").replace(/>/gu, "&gt;");
}

function elementChildren(node, namespace, localName) {
  return node.children.filter((child) => child.namespace === namespace && child.localName === localName);
}

function directTextSlot(node, label) {
  if (node.selfClosing || node.children.length) throw new Error(`${label} must contain direct text only`);
  const raw = node.text.map((segment) => segment.value).join("");
  return { value: decodeXmlText(raw), start: node.openEnd, end: node.closeStart };
}

function propertyTextSlot(property, shape, label) {
  if (!shape.container) return directTextSlot(property, label);
  if (property.selfClosing || property.text.some((segment) => segment.value.trim())) throw new Error(`${label} has non-canonical mixed content`);
  const containers = elementChildren(property, NAMESPACES.rdf, shape.container);
  if (containers.length !== 1 || property.children.length !== 1) throw new Error(`${label} must contain exactly one rdf:${shape.container}`);
  const container = containers[0];
  if (container.selfClosing || container.text.some((segment) => segment.value.trim())) throw new Error(`${label} rdf:${shape.container} has non-canonical mixed content`);
  const items = elementChildren(container, NAMESPACES.rdf, "li");
  if (items.length !== 1 || container.children.length !== 1) throw new Error(`${label} must contain exactly one rdf:li`);
  const item = items[0];
  if (shape.container === "Alt") {
    const language = item.attributes.filter((attribute) => attribute.namespace === NAMESPACES.xml && attribute.localName === "lang");
    if (language.length !== 1 || language[0].value !== "x-default") throw new Error(`${label} rdf:Alt must contain one x-default rdf:li`);
  } else if (item.attributes.length) throw new Error(`${label} rdf:Seq item must not carry attributes`);
  return directTextSlot(item, `${label} rdf:li`);
}

function descendants(node, namespace, localName, output = []) {
  for (const child of node.children) {
    if (child.namespace === namespace && child.localName === localName) output.push(child);
    descendants(child, namespace, localName, output);
  }
  return output;
}

export function inspectCanonicalXmpMetadata(value) {
  const bytes = Buffer.from(value);
  const base = {
    byteLength: bytes.byteLength,
    sha256: sha256(bytes),
    profile: "canonical-simple-v1",
    values: {},
    mutableFields: [],
    issues: [],
    slots: {},
  };
  if (bytes.byteLength > XMP_MAX_BYTES) return { ...base, profile: "unsupported", issues: [`xmp-stream-exceeds-${XMP_MAX_BYTES}-bytes`] };
  let xml;
  try {
    xml = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
    const root = parseXml(xml);
    if (root.namespace !== "adobe:ns:meta/" || root.localName !== "xmpmeta") throw new Error("document element must be x:xmpmeta");
    const rdfRoots = elementChildren(root, NAMESPACES.rdf, "RDF");
    if (rdfRoots.length !== 1 || root.children.length !== 1) throw new Error("x:xmpmeta must contain exactly one rdf:RDF element");
    const descriptions = elementChildren(rdfRoots[0], NAMESPACES.rdf, "Description");
    if (!descriptions.length || descriptions.length !== rdfRoots[0].children.length) throw new Error("rdf:RDF must contain only one or more rdf:Description elements");
    for (const description of descriptions) {
      for (const attribute of description.attributes) {
        if (Object.values(PROPERTY_SHAPES).some((shape) => shape.namespace === attribute.namespace && shape.localName === attribute.localName)) {
          throw new Error(`${attribute.localName} uses an attribute-valued representation outside the canonical simple profile`);
        }
      }
    }
    for (const [field, shape] of Object.entries(PROPERTY_SHAPES)) {
      const matches = descriptions.flatMap((description) => elementChildren(description, shape.namespace, shape.localName));
      if (!matches.length) continue;
      if (matches.length !== 1) throw new Error(`${field} appears more than once`);
      const slot = propertyTextSlot(matches[0], shape, field);
      base.slots[field] = { start: slot.start, end: slot.end };
      if (slot.value !== "") base.values[field] = slot.value;
      base.mutableFields.push(field);
    }
    if (descendants(root, NAMESPACES.rdf, "Description").length !== descriptions.length) {
      throw new Error("nested rdf:Description elements are outside the canonical simple profile");
    }
    return base;
  } catch (error) {
    return { ...base, profile: "unsupported", values: {}, mutableFields: [], slots: {}, issues: [String(error?.message || error)] };
  }
}

export function patchCanonicalXmpMetadata(value, profile, patch) {
  const bytes = Buffer.from(value);
  if (profile.profile !== "canonical-simple-v1" || profile.issues.length) throw new Error(`XMP metadata profile is unsupported: ${profile.issues.join(", ")}`);
  if (profile.sha256 !== sha256(bytes) || profile.byteLength !== bytes.byteLength) throw new Error("XMP metadata bytes no longer match the inspected profile");
  const replacements = [];
  for (const [field, requested] of Object.entries(patch)) {
    const slot = profile.slots[field];
    if (!slot) throw new Error(`set_metadata cannot synchronize XMP field ${field} because the canonical packet does not contain that property`);
    replacements.push({ start: slot.start, end: slot.end, value: encodeXmlText(requested ?? "") });
  }
  replacements.sort((left, right) => right.start - left.start);
  let xml = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  for (const replacement of replacements) xml = `${xml.slice(0, replacement.start)}${replacement.value}${xml.slice(replacement.end)}`;
  return Buffer.from(xml, "utf8");
}
