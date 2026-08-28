import { create } from "@bufbuild/protobuf";
import { createHash } from "node:crypto";
import { DocumentModel } from "../document/index.mjs";
import { FileBlob } from "../shared/file-blob.mjs";
import { isXmlSafeText } from "../shared/xml.mjs";
import {
  ArtifactFamily,
  CodecOperation,
  DocumentChangeType,
  DocumentContentControlType,
  DocumentHeaderFooterReference,
  DocumentImageHorizontalRelativeFrom,
  DocumentImageVerticalRelativeFrom,
  DocumentImageWrapMode,
  DocumentImageWrapSide,
  DocumentNoteKind,
  DocumentProtectionMode,
  DocumentPictureBulletSchema,
  DocumentRevisionFinalizationMode,
  DocumentSectionBreak,
  DocumentSectionLineNumberRestart,
  DocumentSectionPageNumberFormat,
  DocumentStyleType,
  DocumentTableHorizontalAlignment,
  DocumentTableVerticalAlignment,
  DocumentTableVerticalMerge,
} from "../generated/office_kit/artifact/v1/office_artifact_pb.js";
import { OfficeKitCodecError } from "./office-kit-error.mjs";
import {
  assertCodecOptions,
  boundedInputBytes,
  codecLimits,
  invokeOfficeKit,
  invokeOfficeKitLazy,
  OFFICE_KIT_PROTOCOL_VERSION,
  uint32,
} from "./office-kit-runtime.mjs";
import { assertTrustedImportedState } from "./office-kit-source-state.mjs";

const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const DOCUMENT_STATE = Symbol.for("office-kit.document-state");
const DOCUMENT_RUN_STYLE_KEYS = new Set(["runStyleId", "bold", "italic", "underline", "fontFamily", "fontSize", "color", "characterSpacing", "characterSpacingTwips"]);
const DOCUMENT_RUN_DERIVED_STYLE_KEYS = new Set(["resolvedColor", "resolvedFontFamily", "resolvedFontFamilyEastAsia", "resolvedFontFamilyComplexScript"]);
const DOCUMENT_BIBLIOGRAPHY_FIELD_KEYS = [
  "title", "year", "city", "stateProvince", "countryRegion", "publisher", "bookTitle", "journalName", "periodicalTitle", "publicationTitle", "internetSiteTitle",
  "conferenceName", "institution", "department", "volume", "issue", "pages", "edition", "numberVolumes", "chapterNumber", "standardNumber", "shortTitle", "comments", "medium",
  "month", "day", "yearAccessed", "monthAccessed", "dayAccessed", "url", "guid", "lcid", "reporter", "caseNumber", "abbreviatedCaseNumber", "court", "patentNumber", "patentType",
  "broadcaster", "broadcastTitle", "station", "theater", "productionCompany", "distributor", "recordingNumber", "albumTitle", "thesisType", "version", "referenceOrder",
];
const DOCUMENT_CITATION_TAG = /^[A-Za-z0-9_.:-]{1,255}$/;
const DOCUMENT_HEADER_FOOTER_FIELD_COMMANDS = new Set(["PAGE", "NUMPAGES", "SECTION", "SECTIONPAGES", "DATE", "TIME", "CREATEDATE", "SAVEDATE", "PRINTDATE", "AUTHOR", "TITLE", "SUBJECT", "COMMENTS", "FILENAME", "FILESIZE", "NUMWORDS", "NUMCHARS"]);
const DOCUMENT_FIELD_COMMANDS = new Set([...DOCUMENT_HEADER_FOOTER_FIELD_COMMANDS, "BIBLIOGRAPHY"]);
const DOCUMENT_INLINE_FIELD_INSTRUCTION = /^(?:SEQ [A-Za-z][A-Za-z0-9_]{0,39} \\[*] ARABIC|(?:REF|PAGEREF) [A-Za-z][A-Za-z0-9_]{0,39} \\h)$/;

function isCanonicalBibliographyFieldInstruction(value) {
  return /^\s*BIBLIOGRAPHY\s*$/i.test(String(value ?? ""));
}

function documentRgb(value, label) {
  if (value == null || value === "") return undefined;
  const rgb = String(value).replace(/^#/, "").toUpperCase();
  if (!/^[0-9A-F]{6}$/.test(rgb)) throw new OfficeKitCodecError(`${label} color must be a six-digit RGB value.`, [], { code: "invalid_document_formatting" });
  return rgb;
}

function documentRunFormatting(style = {}, label = "Document run") {
  const unsupported = Object.keys(style).filter((key) => !DOCUMENT_RUN_STYLE_KEYS.has(key) && !DOCUMENT_RUN_DERIVED_STYLE_KEYS.has(key));
  if (unsupported.length) throw new OfficeKitCodecError(`${label} uses unsupported run style fields: ${unsupported.join(", ")}.`, [], { code: "unsupported_document_features" });
  const formatting = {};
  if (Object.hasOwn(style, "fontFamily")) formatting.fontFamily = String(style.fontFamily || "");
  if (Object.hasOwn(style, "fontSize")) {
    const points = Number(style.fontSize);
    if (!Number.isFinite(points) || points <= 0 || points > 1_638) throw new OfficeKitCodecError(`${label} fontSize must be greater than 0 and no more than 1638 points.`, [], { code: "invalid_document_formatting" });
    formatting.fontSizeHalfPoints = uint32(Math.round(points * 2), `${label} fontSize`);
  }
  if (Object.hasOwn(style, "color")) formatting.colorRgb = documentRgb(style.color, label);
  if (Object.hasOwn(style, "characterSpacing") || Object.hasOwn(style, "characterSpacingTwips")) {
    const value = Number(style.characterSpacingTwips ?? style.characterSpacing);
    if (!Number.isInteger(value) || value < -31_680 || value > 31_680) throw new OfficeKitCodecError(`${label} character spacing must be an integer from -31680 through 31680 twips.`, [], { code: "invalid_document_formatting" });
    formatting.characterSpacingTwips = value;
  }
  for (const key of ["bold", "italic"]) if (Object.hasOwn(style, key)) formatting[key] = Boolean(style[key]);
  if (Object.hasOwn(style, "underline")) formatting.underline = style.underline === true || style.underline === "single";
  return Object.keys(formatting).length ? formatting : undefined;
}

function publicDocumentRunFormatting(formatting) {
  if (!formatting) return {};
  return {
    ...(formatting.fontFamily !== undefined ? { fontFamily: formatting.fontFamily } : {}),
    ...(formatting.fontSizeHalfPoints !== undefined ? { fontSize: formatting.fontSizeHalfPoints / 2 } : {}),
    ...(formatting.colorRgb !== undefined ? { color: `#${formatting.colorRgb}` } : {}),
    ...(formatting.characterSpacingTwips !== undefined ? { characterSpacingTwips: formatting.characterSpacingTwips } : {}),
    ...(formatting.bold !== undefined ? { bold: formatting.bold } : {}),
    ...(formatting.italic !== undefined ? { italic: formatting.italic } : {}),
    ...(formatting.underline !== undefined ? { underline: formatting.underline } : {}),
  };
}

const DOCUMENT_PARAGRAPH_BORDER_SIDES = ["top", "left", "bottom", "right", "between", "bar"];
const DOCUMENT_PARAGRAPH_BORDER_EDGE_KEYS = new Set(["color", "size", "space"]);

function documentParagraphBorders(value, label) {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new OfficeKitCodecError(`${label} must be a non-empty border object.`, [], { code: "invalid_document_formatting" });
  }
  const unsupported = Object.keys(value).filter((key) => !DOCUMENT_PARAGRAPH_BORDER_SIDES.includes(key));
  if (unsupported.length) {
    throw new OfficeKitCodecError(`${label} uses unsupported border sides: ${unsupported.join(", ")}.`, [], { code: "unsupported_document_features" });
  }
  const result = {};
  for (const side of DOCUMENT_PARAGRAPH_BORDER_SIDES) {
    if (!Object.hasOwn(value, side)) continue;
    const edge = value[side];
    if (!edge || typeof edge !== "object" || Array.isArray(edge)) {
      throw new OfficeKitCodecError(`${label}.${side} must be a border edge object.`, [], { code: "invalid_document_formatting" });
    }
    const unknownEdgeKeys = Object.keys(edge).filter((key) => !DOCUMENT_PARAGRAPH_BORDER_EDGE_KEYS.has(key));
    if (unknownEdgeKeys.length) {
      throw new OfficeKitCodecError(`${label}.${side} uses unsupported fields: ${unknownEdgeKeys.join(", ")}.`, [], { code: "unsupported_document_features" });
    }
    if (!Object.hasOwn(edge, "color") || !Object.hasOwn(edge, "size")) {
      throw new OfficeKitCodecError(`${label}.${side} requires color and size.`, [], { code: "invalid_document_formatting" });
    }
    const size = Number(edge.size);
    if (!Number.isInteger(size) || size < 2 || size > 96) {
      throw new OfficeKitCodecError(`${label}.${side} size must be an integer from 2 through 96 eighths of a point.`, [], { code: "invalid_document_formatting" });
    }
    const space = edge.space === undefined ? 0 : Number(edge.space);
    if (!Number.isInteger(space) || space < 0 || space > 31) {
      throw new OfficeKitCodecError(`${label}.${side} space must be an integer from 0 through 31 points.`, [], { code: "invalid_document_formatting" });
    }
    if (typeof edge.color !== "string" || !/^#[0-9A-Fa-f]{6}$/.test(edge.color)) {
      throw new OfficeKitCodecError(`${label}.${side} color must be a #RRGGBB value.`, [], { code: "invalid_document_formatting" });
    }
    const colorRgb = documentRgb(edge.color, `${label}.${side}`);
    result[side] = {
      colorRgb,
      sizeEighthPoints: size,
      spacePoints: space,
    };
  }
  if (!Object.keys(result).length) {
    throw new OfficeKitCodecError(`${label} requires at least one border edge.`, [], { code: "invalid_document_formatting" });
  }
  return result;
}

function publicDocumentParagraphBorders(value) {
  if (!value) return undefined;
  const result = {};
  for (const side of DOCUMENT_PARAGRAPH_BORDER_SIDES) {
    const edge = value[side];
    if (!edge) continue;
    result[side] = {
      color: `#${edge.colorRgb}`,
      size: edge.sizeEighthPoints,
      space: edge.spacePoints,
    };
  }
  return Object.keys(result).length ? result : undefined;
}

function documentParagraphFormatting(block) {
  const value = block?.paragraphFormat || block?.formatting || {};
  const result = {};
  const text = (model, wire) => { if (value[model] != null) result[wire] = String(value[model]); };
  const integer = (model, wire) => {
    if (value[model] == null) return;
    const number = Number(value[model]);
    if (!Number.isInteger(number) || number < -1_000_000 || number > 1_000_000) throw new OfficeKitCodecError(`Document paragraph ${block.id} ${model} must be a bounded integer.`, [], { code: "invalid_document_formatting" });
    result[wire] = number;
  };
  text("alignment", "alignment");
  for (const [model, wire] of [["leftIndentTwips", "leftIndentTwips"], ["rightIndentTwips", "rightIndentTwips"], ["firstLineIndentTwips", "firstLineIndentTwips"], ["hangingIndentTwips", "hangingIndentTwips"], ["spaceBeforeTwips", "spaceBeforeTwips"], ["spaceAfterTwips", "spaceAfterTwips"], ["lineSpacingTwips", "lineSpacingTwips"]]) integer(model, wire);
  text("lineSpacingRule", "lineSpacingRule");
  if (value.shadingFill != null) result.shadingFill = documentRgb(value.shadingFill, `Document paragraph ${block.id} shadingFill`);
  if (value.borders != null) result.borders = documentParagraphBorders(value.borders, `Document paragraph ${block.id} borders`);
  if (value.keepNext != null) result.keepNext = Boolean(value.keepNext);
  if (value.keepLinesTogether != null) {
    if (typeof value.keepLinesTogether !== "boolean") throw new OfficeKitCodecError(`Document paragraph ${block.id} keepLinesTogether must be boolean.`, [], { code: "invalid_document_formatting" });
    result.keepLinesTogether = value.keepLinesTogether;
  }
  if (value.pageBreakBefore != null) result.pageBreakBefore = Boolean(value.pageBreakBefore);
  if (value.widowControl != null) {
    if (typeof value.widowControl !== "boolean") throw new OfficeKitCodecError(`Document paragraph ${block.id} widowControl must be boolean.`, [], { code: "invalid_document_formatting" });
    result.widowControl = value.widowControl;
  }
  if (value.outlineLevel != null) {
    if (!Number.isInteger(value.outlineLevel) || value.outlineLevel < 0 || value.outlineLevel > 9) throw new OfficeKitCodecError(`Document paragraph ${block.id} outlineLevel must be an integer from 0 through 9.`, [], { code: "invalid_document_formatting" });
    result.outlineLevel = value.outlineLevel;
  }
  if (value.contextualSpacing != null) {
    if (typeof value.contextualSpacing !== "boolean") throw new OfficeKitCodecError(`Document paragraph ${block.id} contextualSpacing must be boolean.`, [], { code: "invalid_document_formatting" });
    result.contextualSpacing = value.contextualSpacing;
  }
  if (value.suppressLineNumbers != null) {
    if (typeof value.suppressLineNumbers !== "boolean") throw new OfficeKitCodecError(`Document paragraph ${block.id} suppressLineNumbers must be boolean.`, [], { code: "invalid_document_formatting" });
    result.suppressLineNumbers = value.suppressLineNumbers;
  }
  return Object.keys(result).length ? result : undefined;
}

function publicDocumentParagraphFormatting(value) {
  if (!value) return undefined;
  const result = {};
  for (const key of ["alignment", "leftIndentTwips", "rightIndentTwips", "firstLineIndentTwips", "hangingIndentTwips", "spaceBeforeTwips", "spaceAfterTwips", "lineSpacingTwips", "lineSpacingRule", "shadingFill", "keepNext", "keepLinesTogether", "pageBreakBefore", "widowControl", "outlineLevel", "contextualSpacing", "suppressLineNumbers"]) {
    if (value[key] !== undefined) result[key] = value[key];
  }
  if (result.shadingFill !== undefined) result.shadingFill = `#${result.shadingFill}`;
  const borders = publicDocumentParagraphBorders(value.borders);
  if (borders) result.borders = borders;
  return Object.keys(result).length ? result : undefined;
}

function planDocumentContentControls(document) {
  const controls = document.blocks.flatMap((block) => {
    if (block.kind === "paragraph") return [
      ...(block.blockContentControl ? [{ block, target: block, control: block.blockContentControl }] : []),
      ...block.runs.filter((run) => run.contentControl).map((run) => ({ block, target: run, control: run.contentControl })),
    ];
    if (block.kind === "table") return (block.cells || []).flatMap((cell) => cell.contentControl
      ? [{ block, target: cell, control: cell.contentControl }]
      : []);
    return [];
  });
  const used = new Set();
  for (const { block, control } of controls) {
    if (!control.id || !String(control.id).trim()) throw new OfficeKitCodecError(`Document block ${block.id} content control requires a non-empty model ID.`, [], { code: "invalid_document_content_control" });
    const nativeId = control.nativeId == null ? undefined : Number(control.nativeId);
    if (nativeId === undefined) continue;
    if (!Number.isInteger(nativeId) || nativeId < 1 || nativeId > 0x7fffffff || used.has(nativeId)) throw new OfficeKitCodecError(`Document block ${block.id} content control ${control.id} has an invalid or duplicate nativeId.`, [], { code: "invalid_document_content_control" });
    used.add(nativeId);
  }
  const result = new Map();
  let next = 1;
  for (const { target, control } of controls) {
    if (control.nativeId != null) {
      result.set(target, Number(control.nativeId));
      continue;
    }
    while (used.has(next)) next += 1;
    if (next > 0x7fffffff) throw new OfficeKitCodecError("Document content controls exhausted the positive native ID range.", [], { code: "invalid_document_content_control" });
    result.set(target, next);
    used.add(next);
    next += 1;
  }
  return result;
}

function documentContentControlTypeName(control) {
  if (!control) return undefined;
  const value = control?.controlType;
  if (value === DocumentContentControlType.CHECKBOX || value === "checkbox") return "checkbox";
  if (value === DocumentContentControlType.DROP_DOWN || value === "dropdown" || value === "drop-down" || value === "drop_down") return "dropdown";
  if (value === DocumentContentControlType.COMBO_BOX || value === "comboBox" || value === "combobox" || value === "combo-box" || value === "combo_box") return "comboBox";
  if (value === DocumentContentControlType.DATE || value === "date" || value === "datepicker" || value === "date-picker" || value === "date_picker") return "date";
  if (value === DocumentContentControlType.PLAIN_TEXT || value === DocumentContentControlType.UNSPECIFIED || value === undefined || value === "text") return "text";
  return undefined;
}

function wireDocumentContentControlChoices(control, blockId, label) {
  if (!Array.isArray(control?.choices) || control.choices.length < 1 || control.choices.length > 256) {
    throw new OfficeKitCodecError(`Document block ${blockId} ${label} content control requires 1 through 256 choices.`, [], { code: "invalid_document_content_control" });
  }
  const values = new Set();
  const displayTexts = new Set();
  const choices = control.choices.map((choice, index) => {
    const displayText = choice?.displayText;
    const value = choice?.value;
    if (typeof displayText !== "string" || typeof value !== "string" || !displayText || !value || displayText.length > 255 || value.length > 255 || !isXmlSafeText(displayText) || !isXmlSafeText(value) || /[\u0000-\u001f\u007f]/.test(displayText + value)) {
      throw new OfficeKitCodecError(`Document block ${blockId} ${label} choice ${index + 1} requires XML-safe displayText and value strings of 1 through 255 characters.`, [], { code: "invalid_document_content_control" });
    }
    if (values.has(value) || displayTexts.has(displayText)) {
      throw new OfficeKitCodecError(`Document block ${blockId} ${label} choice values and displayText strings must be unique.`, [], { code: "invalid_document_content_control" });
    }
    values.add(value);
    displayTexts.add(displayText);
    return { displayText, value };
  });
  return { choices, values };
}

function wireDocumentDropdownState(control, blockId) {
  const { choices, values } = wireDocumentContentControlChoices(control, blockId, "drop-down");
  const selectedValue = control?.selectedValue;
  if (typeof selectedValue !== "string" || !values.has(selectedValue)) {
    throw new OfficeKitCodecError(`Document block ${blockId} drop-down selectedValue must match one declared choice value.`, [], { code: "invalid_document_content_control" });
  }
  return { choices, selectedValue };
}

function wireDocumentComboBoxState(control, blockId) {
  const { choices } = wireDocumentContentControlChoices(control, blockId, "combo-box");
  const value = control?.value;
  if (typeof value !== "string" || !value || value.length > 255 || !isXmlSafeText(value) || /[\u0000-\u001f\u007f]/.test(value)) {
    throw new OfficeKitCodecError(`Document block ${blockId} combo-box value must be an XML-safe string of 1 through 255 characters.`, [], { code: "invalid_document_content_control" });
  }
  return { choices, value };
}

function wireDocumentDateValue(control, blockId) {
  const value = control?.dateValue;
  const match = typeof value === "string" ? /^(\d{4})-(\d{2})-(\d{2})$/.exec(value) : null;
  if (!match) throw new OfficeKitCodecError(`Document block ${blockId} date content-control dateValue must use canonical YYYY-MM-DD form.`, [], { code: "invalid_document_content_control" });
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const leap = year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0);
  const daysInMonth = [31, leap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  if (year < 1 || month < 1 || month > 12 || day < 1 || day > daysInMonth[month - 1]) {
    throw new OfficeKitCodecError(`Document block ${blockId} date content-control dateValue must be a real Gregorian date from 0001-01-01 through 9999-12-31.`, [], { code: "invalid_document_content_control" });
  }
  return value;
}

function wireDocumentContentControl(control, nativeId, blockId) {
  const id = String(control?.id || "").trim();
  const tag = String(control?.tag || "").trim();
  const alias = String(control?.alias ?? tag);
  const typeName = documentContentControlTypeName(control);
  const controlType = typeName === "checkbox"
    ? DocumentContentControlType.CHECKBOX
    : typeName === "dropdown" ? DocumentContentControlType.DROP_DOWN
      : typeName === "comboBox" ? DocumentContentControlType.COMBO_BOX
        : typeName === "date" ? DocumentContentControlType.DATE
      : typeName === "text" ? DocumentContentControlType.PLAIN_TEXT : undefined;
  if (!id || !tag || tag.length > 64 || alias.length > 255 || /[\u0000-\u001f\u007f]/.test(tag + alias) || controlType === undefined) throw new OfficeKitCodecError(`Document block ${blockId} has an invalid content control.`, [], { code: "invalid_document_content_control" });
  if (controlType === DocumentContentControlType.CHECKBOX && typeof control.checked !== "boolean") {
    throw new OfficeKitCodecError(`Document block ${blockId} has an invalid checkbox content-control state.`, [], { code: "invalid_document_content_control" });
  }
  const dropdown = controlType === DocumentContentControlType.DROP_DOWN ? wireDocumentDropdownState(control, blockId) : undefined;
  const comboBox = controlType === DocumentContentControlType.COMBO_BOX ? wireDocumentComboBoxState(control, blockId) : undefined;
  const dateValue = controlType === DocumentContentControlType.DATE ? wireDocumentDateValue(control, blockId) : undefined;
  return {
    id,
    tag,
    alias,
    nativeId,
    controlType,
    checked: controlType === DocumentContentControlType.CHECKBOX && control.checked === true,
    ...(dropdown || {}),
    ...(comboBox || {}),
    ...(dateValue ? { dateValue } : {}),
  };
}

function assertDocumentContentControlVisibleText(control, visibleText, blockId) {
  if (!control) return;
  const controlType = documentContentControlTypeName(control);
  let expected;
  if (controlType === "checkbox") {
    if (typeof control.checked !== "boolean") throw new OfficeKitCodecError(`Document block ${blockId} has an invalid checkbox content-control state.`, [], { code: "invalid_document_content_control" });
    expected = control.checked ? "☒" : "☐";
  } else if (controlType === "dropdown") {
    const state = wireDocumentDropdownState(control, blockId);
    expected = state.choices.find((choice) => choice.value === state.selectedValue).displayText;
  } else if (controlType === "comboBox") {
    const state = wireDocumentComboBoxState(control, blockId);
    expected = state.choices.find((choice) => choice.value === state.value)?.displayText ?? state.value;
  } else if (controlType === "date") {
    expected = wireDocumentDateValue(control, blockId);
  } else if (controlType === "text") {
    return;
  } else {
    throw new OfficeKitCodecError(`Document block ${blockId} has an unsupported content-control type.`, [], { code: "invalid_document_content_control" });
  }
  if (String(visibleText ?? "") !== expected) {
    throw new OfficeKitCodecError(`Document block ${blockId} ${controlType} content-control visible text does not match its typed state.`, [], { code: "invalid_document_content_control" });
  }
}

function wireDocumentTableCellContentControl(control, nativeId, blockId, visibleText) {
  if (!String(control?.alias ?? "").length) {
    throw new OfficeKitCodecError(`Document table cell ${blockId} content control requires a non-empty alias.`, [], { code: "invalid_document_content_control" });
  }
  assertDocumentContentControlVisibleText(control, visibleText, blockId);
  return wireDocumentContentControl(control, nativeId, blockId);
}

function documentRun(run, blockId, contentControlNativeId) {
  const style = run.style || {};
  const formatting = documentRunFormatting(style, `Document block ${blockId}`);
  const inlineInstruction = run.inlineField ? String(run.inlineField.instruction || "").trim() : undefined;
  if (inlineInstruction !== undefined && !DOCUMENT_INLINE_FIELD_INSTRUCTION.test(inlineInstruction)) {
    throw new OfficeKitCodecError(`Document block ${blockId} inline field must be canonical SEQ <label> \\* ARABIC, REF <bookmark> \\h, or PAGEREF <bookmark> \\h.`, [], { code: "invalid_document_inline_field" });
  }
  if (run.contentControl && inlineInstruction !== undefined) throw new OfficeKitCodecError(`Document block ${blockId} run cannot combine a content control and an inline field.`, [], { code: "invalid_document_inline_field" });
  assertDocumentContentControlVisibleText(run.contentControl, run.text, blockId);
  const bookmarkName = inlineInstruction === undefined ? "" : String(run.inlineField?.bookmarkName || "").trim();
  let bookmarkNativeId = "";
  if (run.inlineField?.bookmarkNativeId !== undefined) {
    const value = Number(run.inlineField.bookmarkNativeId);
    if (!Number.isInteger(value) || value < 0 || value > 0xffffffff) throw new OfficeKitCodecError(`Document block ${blockId} inline field bookmarkNativeId must be an unsigned 32-bit integer.`, [], { code: "invalid_document_inline_field" });
    bookmarkNativeId = String(value);
  }
  if (bookmarkName && (!/^[A-Za-z][A-Za-z0-9_]{0,39}$/.test(bookmarkName) || !inlineInstruction.startsWith("SEQ "))) {
    throw new OfficeKitCodecError(`Document block ${blockId} may bookmark only a canonical SEQ cached result with a valid Word bookmark name.`, [], { code: "invalid_document_inline_field" });
  }
  if (!bookmarkName && bookmarkNativeId) throw new OfficeKitCodecError(`Document block ${blockId} inline field bookmarkNativeId requires bookmarkName.`, [], { code: "invalid_document_inline_field" });
  return {
    text: String(run.text ?? ""),
    styleId: style.runStyleId || "",
    bold: style.bold === true,
    italic: style.italic === true,
    underline: style.underline === true || style.underline === "single",
    formatting,
    textContentControl: run.contentControl ? wireDocumentContentControl(run.contentControl, contentControlNativeId, blockId) : undefined,
    inlineField: inlineInstruction === undefined ? undefined : { instruction: inlineInstruction, bookmarkName, bookmarkNativeId },
  };
}

function documentContentControlTopology(paragraph = {}) {
  const blockControl = paragraph.blockContentControl || paragraph.block_content_control;
  const blockControlType = documentContentControlTypeName(blockControl);
  return {
    block: blockControl ? { nativeId: Number(blockControl.nativeId), controlType: blockControlType } : undefined,
    inline: (paragraph.runs || []).flatMap((run, index) => {
      const control = run.textContentControl || run.contentControl;
      const controlType = documentContentControlTypeName(control);
      return control
        ? [{
            index,
            nativeId: Number(control.nativeId),
            controlType,
            ...(controlType === "dropdown" || controlType === "comboBox" ? { choices: (control.choices || []).map((choice) => [String(choice.displayText), String(choice.value)]) } : {}),
          }]
        : [];
    }),
  };
}

function assertDocumentContentControlTopology(block, original) {
  if (!original || original.content.case !== "paragraph") return;
  const requested = documentContentControlTopology(block);
  const source = documentContentControlTopology(original.content.value);
  if (JSON.stringify(requested) !== JSON.stringify(source)) {
    throw new OfficeKitCodecError(`Imported document paragraph ${block.id} content-control topology is source-bound.`, [], { code: "document_content_control_topology_changed" });
  }
}

function documentInlineFieldTopology(runs = []) {
  return runs.flatMap((run, index) => {
    const field = run.inlineField || run.field;
    return field ? [{
      index,
      instruction: String(field.instruction || "").trim(),
      bookmarkName: String(field.bookmarkName || ""),
      bookmarkNativeId: field.bookmarkNativeId === undefined || field.bookmarkNativeId === "" ? undefined : Number(field.bookmarkNativeId),
    }] : [];
  });
}

function assertDocumentInlineFieldTopology(block, original) {
  if (!original || original.content.case !== "paragraph") return;
  const requested = documentInlineFieldTopology(block.runs || []);
  const source = documentInlineFieldTopology(original.content.value.runs || []);
  if (JSON.stringify(requested) !== JSON.stringify(source)) {
    throw new OfficeKitCodecError(`Imported document paragraph ${block.id} inline-field positions and instructions are source-bound.`, [], { code: "document_inline_field_topology_changed" });
  }
}

function sameTableValues(block, original) {
  return JSON.stringify(block.values || []) === JSON.stringify((original.content.value.rows || []).map((row) => [...row.cells]));
}

function documentTableCells(table) {
  const mergeName = (value) => value === DocumentTableVerticalMerge.RESTART
    ? "restart"
    : value === DocumentTableVerticalMerge.CONTINUE ? "continue" : "none";
  return table.rows.flatMap((row, rowIndex) => row.richCells.map((cell, column) => ({
    row: rowIndex,
    column,
    gridColumn: cell.gridColumn,
    columnSpan: cell.columnSpan || 1,
    rowSpan: cell.rowSpan,
    verticalMerge: mergeName(cell.verticalMerge),
    editable: cell.editable,
    textPatchable: cell.textPatchable,
    contentControl: publicDocumentContentControl(cell.textContentControl),
  })));
}

function sameDocumentTableGeometry(block, table) {
  if (block.gridColumns !== table.gridColumns) return false;
  const sourceCells = documentTableCells(table);
  if (!Array.isArray(block.cells) || block.cells.length !== sourceCells.length) return false;
  return block.cells.every((cell, index) => {
    const source = sourceCells[index];
    return cell.row === source.row && cell.column === source.column &&
      cell.gridColumn === source.gridColumn && cell.columnSpan === source.columnSpan &&
      cell.rowSpan === source.rowSpan && cell.verticalMerge === source.verticalMerge &&
      cell.editable === source.editable && cell.textPatchable === source.textPatchable;
  });
}

function documentTableHeaderRowCount(block, rowCount) {
  const value = Number(block.headerRowCount ?? 0);
  if (!Number.isInteger(value) || value < 0 || value > rowCount) {
    throw new OfficeKitCodecError(`Document table ${block.id} headerRowCount must be an integer from 0 through ${rowCount}.`, [], { code: "invalid_document_table" });
  }
  return value;
}

function sameDocumentTableHeaderRows(block, table) {
  return documentTableHeaderRowCount(block, table.rows.length) === Number(table.headerRowCount || 0);
}

function documentTableKeepTogetherRows(block, rowCount) {
  const values = block.keepTogetherRows == null ? [] : block.keepTogetherRows;
  if (!Array.isArray(values)) {
    throw new OfficeKitCodecError(`Document table ${block.id} keepTogetherRows must be an array of physical row indexes.`, [], { code: "invalid_document_table" });
  }
  const normalized = values.map((value) => Number(value));
  if (normalized.some((value) => !Number.isInteger(value) || value < 0 || value >= rowCount)) {
    throw new OfficeKitCodecError(`Document table ${block.id} keepTogetherRows must contain integer row indexes from 0 through ${Math.max(0, rowCount - 1)}.`, [], { code: "invalid_document_table" });
  }
  return [...new Set(normalized)].sort((left, right) => left - right);
}

function documentTableMinimumRowHeights(block, rowCount) {
  const values = block.minimumRowHeightsDxa == null
    ? Array.from({ length: rowCount }, () => 0)
    : block.minimumRowHeightsDxa;
  if (!Array.isArray(values) || values.length !== rowCount) {
    throw new OfficeKitCodecError(`Document table ${block.id} minimumRowHeightsDxa must contain one value for each of its ${rowCount} physical rows.`, [], { code: "invalid_document_table" });
  }
  const normalized = values.map((value) => Number(value));
  if (normalized.some((value) => !Number.isInteger(value) || value < 0 || value > 1_000_000)) {
    throw new OfficeKitCodecError(`Document table ${block.id} minimumRowHeightsDxa values must be integer DXA values from 0 through 1000000.`, [], { code: "invalid_document_table" });
  }
  return normalized;
}

function sameDocumentTableMinimumRowHeights(block, table) {
  const requested = documentTableMinimumRowHeights(block, table.rows.length);
  const source = (table.minimumRowHeightsDxa || []).length === table.rows.length
    ? table.minimumRowHeightsDxa.map((value) => Number(value))
    : Array.from({ length: table.rows.length }, () => 0);
  return requested.length === source.length && requested.every((value, index) => value === source[index]);
}

function documentTableAccessibility(block) {
  if (block.accessibility == null) return {};
  if (typeof block.accessibility !== "object" || Array.isArray(block.accessibility)) {
    throw new OfficeKitCodecError(`Document table ${block.id} accessibility must be an object with title and/or description.`, [], { code: "invalid_document_table" });
  }
  const unsupported = Object.keys(block.accessibility).filter((key) => key !== "title" && key !== "description");
  if (unsupported.length) {
    throw new OfficeKitCodecError(`Document table ${block.id} accessibility does not support ${unsupported.join(", ")}.`, [], { code: "invalid_document_table" });
  }
  const output = {};
  for (const [property, field] of [["title", "accessibilityTitle"], ["description", "accessibilityDescription"]]) {
    if (!Object.hasOwn(block.accessibility, property) || block.accessibility[property] == null) continue;
    const value = block.accessibility[property];
    if (typeof value !== "string" || !value.length || value.length > 32_767 || !isXmlSafeText(value)) {
      throw new OfficeKitCodecError(`Document table ${block.id} accessibility.${property} must contain 1 through 32767 XML-safe characters.`, [], { code: "invalid_document_table" });
    }
    output[field] = value;
  }
  return output;
}

function sameDocumentTableAccessibility(block, table) {
  const requested = documentTableAccessibility(block);
  return requested.accessibilityTitle === (table.accessibilityTitle || undefined) &&
    requested.accessibilityDescription === (table.accessibilityDescription || undefined);
}

function sameDocumentTableKeepTogetherRows(block, table) {
  const requested = documentTableKeepTogetherRows(block, table.rows.length);
  const source = (table.keepTogetherRows || []).map((value) => Number(value));
  return requested.length === source.length && requested.every((value, index) => value === source[index]);
}

function sameDocumentTableContentControlTopology(block, table) {
  const sourceCells = documentTableCells(table);
  if (!Array.isArray(block.cells) || block.cells.length !== sourceCells.length) return false;
  const topology = (control) => {
    if (!control) return undefined;
    const controlType = documentContentControlTypeName(control);
    return {
      nativeId: control.nativeId ?? undefined,
      controlType,
      ...(controlType === "dropdown" || controlType === "comboBox"
        ? { choices: (control.choices || []).map((choice) => [String(choice.displayText), String(choice.value)]) }
        : {}),
    };
  };
  return block.cells.every((cell, index) => JSON.stringify(topology(cell.contentControl)) === JSON.stringify(topology(sourceCells[index].contentControl)));
}

function sameDocumentTableContentControls(block, table) {
  const sourceCells = documentTableCells(table);
  if (!Array.isArray(block.cells) || block.cells.length !== sourceCells.length) return false;
  return block.cells.every((cell, index) => JSON.stringify(cell.contentControl) === JSON.stringify(sourceCells[index].contentControl));
}

function wireDocumentTableTextPatches(block, source) {
  const patches = Array.isArray(block.textPatches) ? block.textPatches : [];
  if (!patches.length) return [];
  if (!source) throw new OfficeKitCodecError(`Document table ${block.id} text patches require a validated imported source.`, [], { code: "unsupported_document_edit" });
  if (patches.length > 10_000) throw new OfficeKitCodecError(`Document table ${block.id} exceeds 10,000 source text patches.`, [], { code: "invalid_document_table" });
  return patches.map((patch) => {
    const row = Number(patch.row);
    const column = Number(patch.column);
    const sourceRow = source.rows?.[row];
    const sourceCell = sourceRow?.richCells?.[column];
    if (!Number.isInteger(row) || !Number.isInteger(column) || row < 0 || column < 0 || !sourceCell) {
      throw new OfficeKitCodecError(`Document table ${block.id} text patch ${row},${column} is outside the source cell matrix.`, [], { code: "invalid_document_table" });
    }
    if (!sourceCell.textPatchable) {
      throw new OfficeKitCodecError(`Document table ${block.id} cell ${row},${column} does not advertise source-bound text replacement capability.`, [], { code: "unsupported_document_edit" });
    }
    const search = String(patch.search ?? "");
    const replacement = String(patch.replacement ?? "");
    if (!search || search.length > 1_000_000 || replacement.length > 1_000_000 || !isXmlSafeText(search) || !isXmlSafeText(replacement)) {
      throw new OfficeKitCodecError(`Document table ${block.id} cell ${row},${column} text patch requires bounded XML-safe strings.`, [], { code: "invalid_document_table" });
    }
    const sourceText = String(sourceRow.cells[column] ?? "");
    return {
      row,
      column,
      search,
      replacement,
      sourceTextSha256: createHash("sha256").update(sourceText, "utf8").digest("hex"),
    };
  });
}

function authoredDocumentTableGeometry(block, contentControlNativeIds) {
  const invalid = (message) => {
    throw new OfficeKitCodecError(`Document table ${block.id} ${message}`, [], { code: "invalid_document_table" });
  };
  if (!Array.isArray(block.cells) || block.cells.length === 0) invalid("requires one explicit geometry record for every physical cell.");
  if (!Number.isInteger(block.gridColumns) || block.gridColumns < 1 || block.gridColumns > 4_096) {
    invalid("gridColumns must be an integer from 1 through 4096.");
  }

  const records = new Map();
  for (const cell of block.cells) {
    if (!Number.isInteger(cell.row) || !Number.isInteger(cell.column) || cell.row < 0 || cell.column < 0 ||
        cell.row >= block.values.length || cell.column >= (block.values[cell.row]?.length || 0)) {
      invalid(`cell ${cell.row},${cell.column} does not identify a physical value cell.`);
    }
    const key = `${cell.row}:${cell.column}`;
    if (records.has(key)) invalid(`contains duplicate geometry for cell ${cell.row},${cell.column}.`);
    records.set(key, cell);
  }

  const rows = block.values.map((values, rowIndex) => {
    if (values.length === 0) invalid(`row ${rowIndex} has no physical cells.`);
    let cursor;
    const richCells = values.map((_value, column) => {
      const source = records.get(`${rowIndex}:${column}`);
      if (!source) invalid(`is missing geometry for cell ${rowIndex},${column}.`);
      if (!Number.isInteger(source.gridColumn) || source.gridColumn < 0 || source.gridColumn > 4_096 ||
          !Number.isInteger(source.columnSpan) || source.columnSpan < 1 || source.columnSpan > 4_096) {
        invalid(`cell ${rowIndex},${column} has invalid bounded grid geometry.`);
      }
      if (cursor !== undefined && source.gridColumn !== cursor) {
        invalid(`cell ${rowIndex},${column} must begin at grid column ${cursor}, not ${source.gridColumn}.`);
      }
      const end = source.gridColumn + source.columnSpan;
      if (end > block.gridColumns) invalid(`cell ${rowIndex},${column} extends beyond gridColumns ${block.gridColumns}.`);
      cursor = end;
      const verticalMerge = String(source.verticalMerge || "none");
      const merge = verticalMerge === "restart"
        ? DocumentTableVerticalMerge.RESTART
        : verticalMerge === "continue" ? DocumentTableVerticalMerge.CONTINUE
          : verticalMerge === "none" ? DocumentTableVerticalMerge.UNSPECIFIED : undefined;
      if (merge === undefined) invalid(`cell ${rowIndex},${column} has unsupported verticalMerge ${verticalMerge}.`);
      const rowSpan = Number(source.rowSpan);
      if (!Number.isInteger(rowSpan) || rowSpan < 0 || rowSpan > 4_096) invalid(`cell ${rowIndex},${column} has invalid rowSpan.`);
      if (verticalMerge === "continue") {
        if (rowSpan !== 0 || String(values[column] ?? "") !== "") invalid(`continuation cell ${rowIndex},${column} must have rowSpan 0 and empty text.`);
      } else {
        if (rowSpan < 1 || source.editable === false) invalid(`origin cell ${rowIndex},${column} must have a positive rowSpan and remain editable.`);
        if (verticalMerge === "none" && rowSpan !== 1) invalid(`unmerged cell ${rowIndex},${column} must have rowSpan 1.`);
      }
      return {
        gridColumn: source.gridColumn,
        columnSpan: source.columnSpan,
        rowSpan,
        verticalMerge: merge,
        editable: verticalMerge !== "continue",
        textContentControl: source.contentControl
          ? wireDocumentTableCellContentControl(source.contentControl, contentControlNativeIds.get(source), `${block.id}/cell/${rowIndex}/${column}`, values[column])
          : undefined,
      };
    });
    const gridBefore = richCells[0].gridColumn;
    const gridAfter = block.gridColumns - cursor;
    return { cells: values.map((value) => String(value ?? "")), richCells, gridBefore, gridAfter };
  });
  if (records.size !== block.values.reduce((total, row) => total + row.length, 0)) invalid("contains geometry outside the physical value matrix.");

  let active = new Map();
  const finish = (group) => {
    if (group.seen !== group.expected) invalid(`merge origin ${group.row},${group.column} declares rowSpan ${group.expected} but spans ${group.seen} rows.`);
  };
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const continued = new Map();
    for (let column = 0; column < rows[rowIndex].richCells.length; column += 1) {
      const cell = rows[rowIndex].richCells[column];
      const key = `${cell.gridColumn}:${cell.columnSpan}`;
      if (cell.verticalMerge === DocumentTableVerticalMerge.CONTINUE) {
        const group = active.get(key);
        if (!group) invalid(`continuation cell ${rowIndex},${column} has no matching restart in the preceding row.`);
        group.seen += 1;
        continued.set(key, group);
      } else if (cell.verticalMerge === DocumentTableVerticalMerge.RESTART) {
        continued.set(key, { row: rowIndex, column, expected: cell.rowSpan, seen: 1 });
      }
    }
    for (const [key, group] of active) if (!continued.has(key) || continued.get(key) !== group) finish(group);
    active = continued;
  }
  for (const group of active.values()) finish(group);
  return { gridColumns: block.gridColumns, rows };
}

function defaultDocumentTableColumnWidths(columns, widthDxa = 9360) {
  const count = Math.max(1, Number(columns) || 1);
  const base = Math.floor(widthDxa / count);
  return Array.from({ length: count }, (_value, index) => base + (index < widthDxa - base * count ? 1 : 0));
}

function documentTableFormatting(block, logicalColumns) {
  const invalid = (message) => {
    throw new OfficeKitCodecError(`Document table ${block.id} ${message}`, [], { code: "invalid_document_table" });
  };
  const dxa = (value, name, { positive = false } = {}) => {
    if (!Number.isInteger(value) || value < (positive ? 1 : 0) || value > 1_000_000) {
      invalid(`${name} must be an integer from ${positive ? 1 : 0} through 1000000.`);
    }
    return value;
  };
  const widthDxa = dxa(block.widthDxa, "widthDxa", { positive: true });
  const indentDxa = dxa(block.indentDxa, "indentDxa");
  const horizontalAlignment = block.horizontalAlignment == null ? undefined : String(block.horizontalAlignment);
  const wireHorizontalAlignment = horizontalAlignment === undefined ? undefined
    : horizontalAlignment === "left" ? DocumentTableHorizontalAlignment.LEFT
      : horizontalAlignment === "center" ? DocumentTableHorizontalAlignment.CENTER
        : horizontalAlignment === "right" ? DocumentTableHorizontalAlignment.RIGHT : undefined;
  if (horizontalAlignment !== undefined && wireHorizontalAlignment === undefined) {
    invalid("horizontalAlignment must be left, center, or right when provided.");
  }
  if ((horizontalAlignment === "center" || horizontalAlignment === "right") && indentDxa !== 0) {
    invalid("center or right horizontalAlignment requires indentDxa 0.");
  }
  if (!Number.isInteger(logicalColumns) || logicalColumns < 1 || logicalColumns > 4_096) {
    invalid("requires between 1 and 4096 logical formatting columns.");
  }
  if (!Array.isArray(block.columnWidthsDxa) || block.columnWidthsDxa.length !== logicalColumns) {
    invalid(`columnWidthsDxa must contain one width for each of ${logicalColumns} logical grid columns.`);
  }
  const columnWidthsDxa = block.columnWidthsDxa.map((value, index) => dxa(value, `columnWidthsDxa[${index}]`, { positive: true }));
  if (columnWidthsDxa.reduce((sum, value) => sum + value, 0) !== widthDxa) {
    invalid("columnWidthsDxa must sum exactly to widthDxa.");
  }
  const margins = block.cellMarginsDxa;
  if (!margins || typeof margins !== "object") invalid("cellMarginsDxa must define top, bottom, start, and end margins.");
  const cellMarginsDxa = {
    top: dxa(margins.top, "cellMarginsDxa.top"),
    bottom: dxa(margins.bottom, "cellMarginsDxa.bottom"),
    start: dxa(margins.start, "cellMarginsDxa.start"),
    end: dxa(margins.end, "cellMarginsDxa.end"),
  };
  const borderColor = String(block.borderColor ?? "");
  const headerFill = String(block.headerFill ?? "");
  if (!/^[0-9A-F]{6}$/.test(borderColor)) invalid("borderColor must be a six-digit uppercase RGB value.");
  if (!/^[0-9A-F]{6}$/.test(headerFill)) invalid("headerFill must be a six-digit uppercase RGB value.");
  const borderSize = block.borderSize;
  if (!Number.isInteger(borderSize) || borderSize < 0 || borderSize > 96 || borderSize === 1) {
    invalid("borderSize must be zero or an integer from 2 through 96 eighths of a point.");
  }
  const verticalAlignment = block.verticalAlignment == null ? undefined : String(block.verticalAlignment);
  const wireVerticalAlignment = verticalAlignment === undefined ? undefined
    : verticalAlignment === "top" ? DocumentTableVerticalAlignment.TOP
      : verticalAlignment === "center" ? DocumentTableVerticalAlignment.CENTER
        : verticalAlignment === "bottom" ? DocumentTableVerticalAlignment.BOTTOM : undefined;
  if (verticalAlignment !== undefined && wireVerticalAlignment === undefined) {
    invalid("verticalAlignment must be top, center, or bottom when provided.");
  }
  return {
    widthDxa,
    indentDxa,
    columnWidthsDxa,
    cellMarginsDxa,
    borderColor,
    borderSize,
    headerFill,
    ...(wireHorizontalAlignment === undefined ? {} : { horizontalAlignment: wireHorizontalAlignment }),
    ...(wireVerticalAlignment === undefined ? {} : { verticalAlignment: wireVerticalAlignment }),
  };
}

function documentTableFormattingConfig(table) {
  const logicalColumns = table.gridColumns || Math.max(1, ...table.rows.map((row) => row.cells.length));
  const formatting = table.formatting;
  if (formatting) {
    return {
      widthDxa: formatting.widthDxa,
      indentDxa: formatting.indentDxa,
      columnWidthsDxa: [...formatting.columnWidthsDxa],
      cellMarginsDxa: { ...formatting.cellMarginsDxa },
      borderColor: formatting.borderColor,
      borderSize: formatting.borderSize,
      headerFill: formatting.headerFill,
      ...(formatting.horizontalAlignment === DocumentTableHorizontalAlignment.LEFT ? { horizontalAlignment: "left" }
        : formatting.horizontalAlignment === DocumentTableHorizontalAlignment.CENTER ? { horizontalAlignment: "center" }
          : formatting.horizontalAlignment === DocumentTableHorizontalAlignment.RIGHT ? { horizontalAlignment: "right" } : {}),
      ...(formatting.verticalAlignment === DocumentTableVerticalAlignment.TOP ? { verticalAlignment: "top" }
        : formatting.verticalAlignment === DocumentTableVerticalAlignment.CENTER ? { verticalAlignment: "center" }
          : formatting.verticalAlignment === DocumentTableVerticalAlignment.BOTTOM ? { verticalAlignment: "bottom" } : {}),
    };
  }
  return {
    widthDxa: 9360,
    indentDxa: 120,
    columnWidthsDxa: defaultDocumentTableColumnWidths(logicalColumns),
    cellMarginsDxa: { top: 80, bottom: 80, start: 120, end: 120 },
    borderColor: "D9D9D9",
    borderSize: 4,
    headerFill: "F2F4F7",
  };
}

function sameDocumentTableFormatting(block, table) {
  const expected = documentTableFormattingConfig(table);
  return block.widthDxa === expected.widthDxa && block.indentDxa === expected.indentDxa &&
    JSON.stringify(block.columnWidthsDxa) === JSON.stringify(expected.columnWidthsDxa) &&
    block.cellMarginsDxa?.top === expected.cellMarginsDxa.top &&
    block.cellMarginsDxa?.bottom === expected.cellMarginsDxa.bottom &&
    block.cellMarginsDxa?.start === expected.cellMarginsDxa.start &&
    block.cellMarginsDxa?.end === expected.cellMarginsDxa.end &&
    block.borderColor === expected.borderColor && block.borderSize === expected.borderSize &&
    block.headerFill === expected.headerFill && block.horizontalAlignment === expected.horizontalAlignment &&
    block.verticalAlignment === expected.verticalAlignment;
}

const DOCUMENT_PICTURE_BULLET_EMU_PER_POINT = 12_700;
const DOCUMENT_PICTURE_BULLET_MAX_BYTES = 8 * 1024 * 1024;

function documentPictureBulletAsset(dataUrl, assets, label) {
  const match = /^data:(image\/(?:png|jpeg|gif));base64,([A-Za-z0-9+/=]+)$/i.exec(String(dataUrl || ""));
  if (!match) throw new OfficeKitCodecError(`${label} must use a base64 PNG, JPEG, or GIF data URL.`, [], { code: "invalid_document_picture_bullet" });
  const contentType = match[1].toLowerCase();
  const encoded = match[2];
  if (encoded.length % 4 === 1 || /=[^=]|={3,}/.test(encoded)) {
    throw new OfficeKitCodecError(`${label} contains invalid base64.`, [], { code: "invalid_document_picture_bullet" });
  }
  const bytes = new Uint8Array(Buffer.from(encoded, "base64"));
  if (!bytes.length || bytes.length > DOCUMENT_PICTURE_BULLET_MAX_BYTES || Buffer.from(bytes).toString("base64") !== encoded) {
    throw new OfficeKitCodecError(`${label} must contain 1 through ${DOCUMENT_PICTURE_BULLET_MAX_BYTES} valid decoded bytes.`, [], { code: "invalid_document_picture_bullet" });
  }
  const matchesFormat = contentType === "image/png"
    ? Buffer.from(bytes.subarray(0, 8)).equals(Buffer.from("89504e470d0a1a0a", "hex"))
    : contentType === "image/jpeg"
      ? bytes.length >= 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff
      : bytes.length >= 6 && new Set(["GIF87a", "GIF89a"]).has(Buffer.from(bytes.subarray(0, 6)).toString("ascii"));
  if (!matchesFormat) throw new OfficeKitCodecError(`${label} bytes do not match ${contentType}.`, [], { code: "invalid_document_picture_bullet" });
  const sha256 = createHash("sha256").update(bytes).digest("hex");
  const assetId = `asset/document/image/${sha256}`;
  const extension = contentType === "image/jpeg" ? "jpg" : contentType.slice("image/".length);
  assets.set(assetId, { id: assetId, fileName: `picture-bullet-${sha256.slice(0, 16)}.${extension}`, contentType, data: bytes, sha256 });
  return assetId;
}

function documentPictureBulletUri(value, label) {
  const uri = String(value || "");
  if (!uri || uri.length > 4_096 || /[\u0000-\u001f\u007f]/.test(uri)) {
    throw new OfficeKitCodecError(`${label} URI must contain 1 through 4096 characters without controls.`, [], { code: "invalid_document_picture_bullet" });
  }
  let parsed;
  try { parsed = new URL(uri); } catch { parsed = undefined; }
  if (!parsed || !new Set(["http:", "https:"]).has(parsed.protocol)) {
    throw new OfficeKitCodecError(`${label} URI must be absolute http(s).`, [], { code: "invalid_document_picture_bullet" });
  }
  return uri;
}

function wireDocumentPictureBullet(value, assets, label) {
  if (!value) return undefined;
  if (typeof value !== "object" || Array.isArray(value) || Boolean(value.dataUrl) === Boolean(value.uri)) {
    throw new OfficeKitCodecError(`${label} requires exactly one embedded dataUrl or external uri.`, [], { code: "invalid_document_picture_bullet" });
  }
  const source = value.dataUrl
    ? { case: "assetId", value: documentPictureBulletAsset(value.dataUrl, assets, label) }
    : { case: "uri", value: documentPictureBulletUri(value.uri, label) };
  const widthEmu = Math.round(Number(value.widthPt) * DOCUMENT_PICTURE_BULLET_EMU_PER_POINT);
  const heightEmu = Math.round(Number(value.heightPt) * DOCUMENT_PICTURE_BULLET_EMU_PER_POINT);
  const altText = String(value.alt || "Picture bullet");
  if (!Number.isSafeInteger(widthEmu) || !Number.isSafeInteger(heightEmu) ||
      widthEmu < 4 * DOCUMENT_PICTURE_BULLET_EMU_PER_POINT || widthEmu > 72 * DOCUMENT_PICTURE_BULLET_EMU_PER_POINT ||
      heightEmu < 4 * DOCUMENT_PICTURE_BULLET_EMU_PER_POINT || heightEmu > 72 * DOCUMENT_PICTURE_BULLET_EMU_PER_POINT ||
      !altText || altText.length > 255 || /[\u0000-\u001f\u007f]/.test(altText)) {
    throw new OfficeKitCodecError(`${label} dimensions or alternative text are outside the bounded DOCX profile.`, [], { code: "invalid_document_picture_bullet" });
  }
  return create(DocumentPictureBulletSchema, { source, widthEmu, heightEmu, altText });
}

function publicDocumentPictureBullet(value, assets, label) {
  if (!value) return undefined;
  let dataUrl;
  let uri;
  if (value.source?.case === "assetId") {
    const asset = assets.get(value.source.value);
    if (!asset || !new Set(["image/png", "image/jpeg", "image/gif"]).has(asset.contentType)) {
      throw new OfficeKitCodecError(`${label} references a missing or unsupported image asset.`, [], { code: "invalid_document_asset" });
    }
    dataUrl = `data:${asset.contentType};base64,${Buffer.from(asset.data).toString("base64")}`;
  } else if (value.source?.case === "uri") {
    uri = documentPictureBulletUri(value.source.value, label);
  } else {
    throw new OfficeKitCodecError(`${label} has no image source.`, [], { code: "invalid_document_picture_bullet" });
  }
  const widthPt = Number(value.widthEmu) / DOCUMENT_PICTURE_BULLET_EMU_PER_POINT;
  const heightPt = Number(value.heightEmu) / DOCUMENT_PICTURE_BULLET_EMU_PER_POINT;
  if (!Number.isFinite(widthPt) || widthPt < 4 || widthPt > 72 || !Number.isFinite(heightPt) || heightPt < 4 || heightPt > 72) {
    throw new OfficeKitCodecError(`${label} dimensions are outside the bounded DOCX profile.`, [], { code: "invalid_document_picture_bullet" });
  }
  return { dataUrl, uri, widthPt, heightPt, alt: String(value.altText || "Picture bullet") };
}

function sameDocumentPictureBullet(left, right, assets, label) {
  const publicRight = publicDocumentPictureBullet(right, assets, label);
  if (!left || !publicRight) return !left && !publicRight;
  return (left.dataUrl || "") === (publicRight.dataUrl || "") && (left.uri || "") === (publicRight.uri || "") &&
    left.widthPt === publicRight.widthPt && left.heightPt === publicRight.heightPt && left.alt === publicRight.alt;
}

function documentPictureBulletDefinitionKey(value) {
  if (!value) return "none";
  return JSON.stringify([value.source?.case || "none", value.source?.value || "", value.widthEmu, value.heightEmu, value.altText]);
}

function sameDocumentNumbering(block, paragraph, assets) {
  const numbering = paragraph.numbering;
  if (!numbering || block.kind !== "listItem") return false;
  const numberFormat = numbering.numberFormat || "decimal";
  return block.text === paragraph.text &&
    block.listType === (numberFormat === "bullet" ? "bullet" : "number") &&
    block.numberFormat === numberFormat &&
    block.level === numbering.level &&
    block.start === (numbering.start || 1) &&
    block.levelText === (numbering.levelText || (numberFormat === "bullet" ? "•" : `%${numbering.level + 1}.`)) &&
    block.numberingId === numbering.numberingId &&
    block.abstractNumberingId === numbering.abstractNumberingId &&
    (block.numberingStyleId || "") === (numbering.numberingStyleId || "") &&
    sameDocumentPictureBullet(block.pictureBullet, numbering.pictureBullet, assets, `Document list item ${block.id} picture bullet`);
}

function sameDocumentNumberingIdentity(block, numbering) {
  return numbering && block.kind === "listItem" &&
    block.level === numbering.level &&
    block.numberingId === numbering.numberingId &&
    block.abstractNumberingId === numbering.abstractNumberingId &&
    (block.numberingStyleId || "") === (numbering.numberingStyleId || "");
}

function editedDocumentNumbering(block, source, assets) {
  if (!sameDocumentNumberingIdentity(block, source)) {
    throw new OfficeKitCodecError(`Document list item ${block.id} numbering identity, level, and style linkage are source-bound.`, [], { code: "unsupported_document_edit" });
  }
  const numberFormat = String(block.numberFormat ?? "");
  const levelText = String(block.levelText ?? "");
  const start = uint32(block.start, `Document list item ${block.id} start`);
  if (numberFormat.length > 128) {
    throw new OfficeKitCodecError(`Document list item ${block.id} numberFormat exceeds 128 characters.`, [], { code: "invalid_document_numbering" });
  }
  if (levelText.length > 1_024) {
    throw new OfficeKitCodecError(`Document list item ${block.id} levelText exceeds 1024 characters.`, [], { code: "invalid_document_numbering" });
  }
  const listType = numberFormat === "bullet" ? "bullet" : "number";
  if (block.listType !== listType) {
    throw new OfficeKitCodecError(`Document list item ${block.id} listType must be ${listType} for numberFormat ${numberFormat || "(empty)"}.`, [], { code: "invalid_document_numbering" });
  }
  const pictureBullet = wireDocumentPictureBullet(block.pictureBullet, assets, `Document list item ${block.id} picture bullet`);
  if (Boolean(pictureBullet) !== Boolean(source.pictureBullet) ||
      pictureBullet && pictureBullet.source?.case !== source.pictureBullet?.source?.case) {
    throw new OfficeKitCodecError(`Document list item ${block.id} picture-bullet source topology is source-bound.`, [], { code: "unsupported_document_edit" });
  }
  return { ...source, numberFormat, start, levelText, pictureBullet };
}

function directDocumentNumberingPlan(document, assets) {
  const groups = new Map();
  const usedNumberingIds = new Set();
  const usedAbstractIds = new Set();
  const result = new Map();
  const invalid = (message) => {
    throw new OfficeKitCodecError(message, [], { code: "invalid_document_numbering" });
  };
  const integer = (value, name, { positive = false } = {}) => {
    const normalized = typeof value === "string" && /^\d+$/.test(value) ? Number(value) : value;
    if (!Number.isInteger(normalized) || normalized < (positive ? 1 : 0) || normalized > 0x7fff_ffff) {
      invalid(`${name} must be ${positive ? "a positive" : "a non-negative"} WordprocessingML signed integer.`);
    }
    return normalized;
  };
  const sameDefinition = (left, right) => left.numberFormat === right.numberFormat && left.start === right.start && left.levelText === right.levelText &&
    documentPictureBulletDefinitionKey(left.pictureBullet) === documentPictureBulletDefinitionKey(right.pictureBullet);

  for (const block of document.blocks.filter((item) => item.kind === "listItem")) {
    if (block.numberingStyleId) {
      throw new OfficeKitCodecError(`The DOCX NativeAOT vertical slice cannot directly author style-linked numbering for list item ${block.id}.`, [], { code: "unsupported_document_features" });
    }
    const level = integer(block.level, `Document list item ${block.id} level`);
    if (level > 8) invalid(`Document list item ${block.id} level must be between 0 and 8.`);
    const start = integer(block.start, `Document list item ${block.id} start`, { positive: true });
    const numberFormat = String(block.numberFormat || "");
    const levelText = String(block.levelText || "");
    if (!numberFormat || numberFormat.length > 128) invalid(`Document list item ${block.id} numberFormat must contain 1 through 128 characters.`);
    if (!levelText || levelText.length > 1_024) invalid(`Document list item ${block.id} levelText must contain 1 through 1024 characters.`);
    const expectedListType = numberFormat === "bullet" ? "bullet" : "number";
    if (block.listType !== expectedListType) invalid(`Document list item ${block.id} listType must be ${expectedListType} for numberFormat ${numberFormat}.`);
    if (block.pictureBullet && (numberFormat !== "bullet" || [...levelText].length !== 1)) {
      invalid(`Document list item ${block.id} picture bullet requires bullet semantics and exactly one levelText character.`);
    }

    const explicitNumberingId = block.numberingId == null ? undefined : integer(block.numberingId, `Document list item ${block.id} numberingId`, { positive: true });
    const explicitAbstractId = block.abstractNumberingId == null ? undefined : integer(block.abstractNumberingId, `Document list item ${block.id} abstractNumberingId`);
    if (explicitNumberingId != null) usedNumberingIds.add(explicitNumberingId);
    if (explicitAbstractId != null) usedAbstractIds.add(explicitAbstractId);
    const key = explicitNumberingId == null ? `default:${block.listType}` : `native:${explicitNumberingId}`;
    if (!groups.has(key)) groups.set(key, { blocks: [], definitions: new Map(), explicitNumberingId, abstractIds: new Set() });
    const group = groups.get(key);
    if (group.explicitNumberingId !== explicitNumberingId) invalid(`Document numbering group ${key} has conflicting numbering IDs.`);
    if (explicitAbstractId != null) group.abstractIds.add(explicitAbstractId);
    const pictureBullet = wireDocumentPictureBullet(block.pictureBullet, assets, `Document list item ${block.id} picture bullet`);
    const definition = { numberFormat, start, levelText, pictureBullet };
    const existing = group.definitions.get(level);
    if (existing && !sameDefinition(existing, definition)) invalid(`Document numbering ${explicitNumberingId ?? key} level ${level} has conflicting definitions.`);
    group.definitions.set(level, definition);
    group.blocks.push({ block, level, definition });
  }

  const allocate = (used, start = 1) => {
    let candidate = start;
    while (used.has(candidate)) candidate += 1;
    if (candidate > 0x7fff_ffff) invalid("Document numbering ID space is exhausted.");
    used.add(candidate);
    return candidate;
  };
  const sharedDefinitions = new Map();
  for (const [key, group] of groups) {
    if (group.abstractIds.size > 1) invalid(`Document numbering ${group.explicitNumberingId ?? key} references conflicting abstract numbering IDs.`);
    const numberingId = group.explicitNumberingId ?? allocate(usedNumberingIds);
    const abstractNumberingId = group.abstractIds.size ? [...group.abstractIds][0] : allocate(usedAbstractIds);
    for (const [level, definition] of group.definitions) {
      const definitionKey = `${abstractNumberingId}:${level}`;
      const existing = sharedDefinitions.get(definitionKey);
      if (existing && !sameDefinition(existing, definition)) invalid(`Document abstract numbering ${abstractNumberingId} level ${level} has conflicting definitions.`);
      sharedDefinitions.set(definitionKey, definition);
    }
    for (const { block, level, definition } of group.blocks) {
      result.set(block, { numberingId, abstractNumberingId, level, ...definition });
    }
  }
  return result;
}

function sameDocumentHyperlink(block, source) {
  if (block.kind !== "hyperlink" || block.text !== source.text) return false;
  if (block.styleId !== (source.styleId || "Normal")) return false;
  if ((block.relationshipId || "") !== (source.relationshipId || "")) return false;
  if ((block.tooltip ?? undefined) !== source.tooltip) return false;
  if (block.history !== (source.history ?? true)) return false;
  if (source.target.case === "externalUri") return !block.anchor && block.url === source.target.value;
  if (source.target.case === "internalAnchor") return block.anchor === source.target.value && !block.url;
  return false;
}

function documentHyperlink(block, original) {
  const source = original?.content.case === "hyperlink" ? original.content.value : undefined;
  if (source && (block.relationshipId || "") !== (source.relationshipId || "")) {
    throw new OfficeKitCodecError(`Document hyperlink ${block.id} relationshipId is a source locator and cannot be edited directly.`, [], { code: "unsupported_document_edit" });
  }
  const text = String(block.text ?? "");
  if (text.length > 1_000_000) throw new OfficeKitCodecError(`Document hyperlink ${block.id} text exceeds 1,000,000 characters.`, [], { code: "invalid_document_hyperlink" });
  const anchor = String(block.anchor || "").trim();
  const url = String(block.url || "");
  let target;
  if (anchor) {
    if (anchor.length > 255 || [...anchor].some((character) => /[\u0000-\u001f\u007f]/.test(character))) {
      throw new OfficeKitCodecError(`Document hyperlink ${block.id} anchor must contain 1 through 255 characters without controls.`, [], { code: "invalid_document_hyperlink" });
    }
    target = { case: "internalAnchor", value: anchor };
  } else {
    let parsed;
    try { parsed = new URL(url); } catch { parsed = undefined; }
    if (!parsed || !new Set(["http:", "https:"]).has(parsed.protocol) || url.length > 4_096 || /[\u0000-\u001f\u007f]/.test(url)) {
      throw new OfficeKitCodecError(`Document hyperlink ${block.id} URI must be an absolute http(s) URI of at most 4096 characters without controls.`, [], { code: "invalid_document_hyperlink" });
    }
    target = { case: "externalUri", value: url };
  }
  if (block.tooltip != null && String(block.tooltip).length > 260) {
    throw new OfficeKitCodecError(`Document hyperlink ${block.id} tooltip exceeds 260 characters.`, [], { code: "invalid_document_hyperlink" });
  }
  const originalHistory = source?.history;
  const history = source && block.history === (originalHistory ?? true) ? originalHistory : block.history;
  return {
    text,
    target,
    relationshipId: source?.relationshipId || "",
    tooltip: block.tooltip == null ? undefined : String(block.tooltip),
    history,
  };
}

function documentField(block, original) {
  if (original?.source?.editable === false) {
    throw new OfficeKitCodecError(`Document field ${block.id} is source-preserved but its instruction or result topology is not editable.`, [], { code: "unsupported_document_edit" });
  }
  const instruction = String(block.instruction ?? "");
  const display = String(block.display ?? "");
  if (!instruction.trim() || instruction.length > 8_192 || /[\u0000-\u001f\u007f]/.test(instruction)) {
    throw new OfficeKitCodecError(`Document field ${block.id} instruction must contain 1 through 8192 characters without controls.`, [], { code: "invalid_document_field" });
  }
  const command = /^[A-Za-z]+/.exec(instruction.trimStart())?.[0]?.toUpperCase();
  const complex = Boolean(block.complex);
  if (!command || (complex ? command !== "TOC" : !DOCUMENT_FIELD_COMMANDS.has(command))) {
    throw new OfficeKitCodecError(`Document field ${block.id} command ${command || "(missing)"} is outside the bounded editable field catalog.`, [], { code: "invalid_document_field" });
  }
  if (complex && !/^TOC \\o "[1-9]-[1-9]"(?: \\h)?(?: \\z)?(?: \\u)?$/.test(instruction)) {
    throw new OfficeKitCodecError(`Document field ${block.id} complex TOC instruction is outside the canonical bounded profile.`, [], { code: "invalid_document_field" });
  }
  if (!complex && command === "BIBLIOGRAPHY" && !isCanonicalBibliographyFieldInstruction(instruction)) {
    throw new OfficeKitCodecError(`Document field ${block.id} BIBLIOGRAPHY instruction must not contain switches or arguments.`, [], { code: "invalid_document_field" });
  }
  const originalInstruction = original?.content?.case === "field" ? original.content.value?.instruction : undefined;
  if (isCanonicalBibliographyFieldInstruction(originalInstruction) && instruction !== originalInstruction) {
    throw new OfficeKitCodecError(`Imported document BIBLIOGRAPHY field ${block.id} may update only its cached display text.`, [], { code: "unsupported_document_edit" });
  }
  if (display.length > 1_000_000) throw new OfficeKitCodecError(`Document field ${block.id} display text exceeds 1,000,000 characters.`, [], { code: "invalid_document_field" });
  return { instruction, display, complex };
}

function documentCommentSnapshot(comment) {
  return {
    id: comment.id,
    targetId: comment.targetId,
    author: comment.author,
    initials: comment.initials,
    date: comment.date,
    text: comment.text,
    resolved: comment.resolved,
    parentId: comment.parentId,
    paraId: comment.paraId,
    durableId: comment.durableId,
    dateUtc: comment.dateUtc,
    person: comment.person,
    intelligentPlaceholder: comment.intelligentPlaceholder,
  };
}

const DOCUMENT_COMMENT_HEX_ID = /^[0-9A-F]{8}$/;

function validateDocumentCommentThreads(document) {
  const byId = new Map();
  for (const comment of document.comments) {
    const id = String(comment.id || "");
    if (!id || byId.has(id)) {
      throw new OfficeKitCodecError("Document comments require unique, non-empty IDs.", [], { code: "invalid_document_comment" });
    }
    byId.set(id, comment);
  }
  for (const comment of document.comments) {
    if (comment.parentId) {
      const parent = byId.get(String(comment.parentId));
      if (!parent) {
        throw new OfficeKitCodecError(`Document comment ${comment.id} references missing parent ${comment.parentId}.`, [], { code: "invalid_document_comment_thread" });
      }
      if (parent.parentId) {
        throw new OfficeKitCodecError(`Document comment ${comment.id} is a nested reply; OfficeKit supports roots plus direct replies only.`, [], { code: "unsupported_document_comment_thread" });
      }
      if (parent.targetId !== comment.targetId) {
        throw new OfficeKitCodecError(`Document comment ${comment.id} and root ${parent.id} must target the same block.`, [], { code: "invalid_document_comment_thread" });
      }
      if (comment.intelligentPlaceholder) {
        throw new OfficeKitCodecError(`Document reply ${comment.id} cannot be an intelligent placeholder.`, [], { code: "invalid_document_comment_thread" });
      }
    }
    for (const [name, value] of [["paraId", comment.paraId], ["durableId", comment.durableId]]) {
      if (value != null && value !== "" && !DOCUMENT_COMMENT_HEX_ID.test(String(value).toUpperCase())) {
        throw new OfficeKitCodecError(`Document comment ${comment.id} ${name} must contain exactly eight hexadecimal digits.`, [], { code: "invalid_document_comment" });
      }
    }
    if (comment.durableId) {
      const durableNumber = Number.parseInt(comment.durableId, 16);
      if (durableNumber <= 0 || durableNumber >= 0x7FFFFFFF) {
        throw new OfficeKitCodecError(`Document comment ${comment.id} durableId must be between 00000001 and 7FFFFFFE.`, [], { code: "invalid_document_comment" });
      }
    }
    if (comment.dateUtc != null) {
      const dateUtc = String(comment.dateUtc);
      if (!dateUtc || dateUtc.length > 64 || Number.isNaN(Date.parse(dateUtc))) {
        throw new OfficeKitCodecError(`Document comment ${comment.id} dateUtc must be an ISO 8601 date-time of at most 64 characters.`, [], { code: "invalid_document_comment" });
      }
    }
    if (comment.person) {
      const providerId = String(comment.person.providerId ?? "");
      const userId = String(comment.person.userId ?? "");
      if (!providerId || !userId || providerId.length > 100 || userId.length > 300) {
        throw new OfficeKitCodecError(`Document comment ${comment.id} person requires providerId of 1 through 100 characters and userId of 1 through 300 characters.`, [], { code: "invalid_document_comment" });
      }
    }
  }
  const commentsByAuthor = new Map();
  for (const comment of document.comments) {
    const author = String(comment.author || "");
    if (!commentsByAuthor.has(author)) commentsByAuthor.set(author, []);
    commentsByAuthor.get(author).push(comment);
  }
  for (const comments of commentsByAuthor.values()) {
    const profiles = new Set(comments.map((comment) => comment.person
      ? `${comment.person.providerId}\u0000${comment.person.userId}`
      : ""));
    if (profiles.size > 1) {
      throw new OfficeKitCodecError(`Document comment author ${comments[0].author} has inconsistent people metadata.`, [], { code: "invalid_document_comment" });
    }
  }
}

function documentBookmarkSnapshot(bookmark) {
  return {
    id: bookmark.id,
    name: bookmark.name,
    targetId: bookmark.targetId,
    endTargetId: bookmark.endTargetId,
    nativeId: bookmark.nativeId,
  };
}

function wireDocumentBibliographySource(source) {
  const tag = String(source?.tag || "");
  if (!DOCUMENT_CITATION_TAG.test(tag)) {
    throw new OfficeKitCodecError(`Document bibliography source ${source?.id || "(unknown)"} tag must contain 1 through 255 ASCII letters, digits, periods, underscores, colons, or hyphens.`, [], { code: "invalid_document_bibliography" });
  }
  const authors = (source.authors || []).map((author) => ({
    first: String(author?.first || ""),
    middle: String(author?.middle || ""),
    last: String(author?.last || ""),
  }));
  const fields = Object.fromEntries(DOCUMENT_BIBLIOGRAPHY_FIELD_KEYS.flatMap((key) =>
    source[key] === undefined || source[key] === null || source[key] === "" ? [] : [[key, String(source[key])]]));
  return {
    id: String(source.id || `bibliography/${tag}`),
    tag,
    sourceType: String(source.sourceType || "Misc"),
    authors,
    corporateAuthor: String(source.corporateAuthor || ""),
    fields,
  };
}

function publicDocumentBibliographySource(source) {
  return {
    id: source.id,
    tag: source.tag,
    sourceType: source.sourceType,
    authors: (source.authors || []).map((author) => ({ first: author.first, middle: author.middle, last: author.last })),
    corporateAuthor: source.corporateAuthor || undefined,
    ...(source.fields || {}),
  };
}

function wireDocumentBibliography(document, original) {
  const settings = {
    selectedStyle: String(document.bibliography?.selectedStyle || ""),
    styleName: String(document.bibliography?.styleName || ""),
    uri: String(document.bibliography?.uri || ""),
  };
  const sources = document.bibliographySources.map(wireDocumentBibliographySource);
  if (!sources.length && !Object.values(settings).some(Boolean)) return undefined;
  if (original) {
    if (sources.length !== original.sources.length) {
      throw new OfficeKitCodecError(`Source-preserving DOCX export requires the original ${original.sources.length}-source bibliography topology; the document contains ${sources.length} sources.`, [], { code: "document_bibliography_topology_changed" });
    }
    for (let index = 0; index < sources.length; index += 1) {
      if (sources[index].id !== original.sources[index].id || sources[index].tag !== original.sources[index].tag) {
        throw new OfficeKitCodecError(`Imported document bibliography source ${index} ID, tag, and order are source-bound.`, [], { code: "unsupported_document_bibliography_edit" });
      }
    }
  }
  return { ...settings, sources, source: original?.source };
}

function wireDocumentCitation(block, original) {
  const tag = String(block.metadata?.tag ?? block.metadata?.bibliographyTag ?? "");
  if (!DOCUMENT_CITATION_TAG.test(tag)) {
    throw new OfficeKitCodecError(`Document citation ${block.id} tag must contain 1 through 255 ASCII letters, digits, periods, underscores, colons, or hyphens.`, [], { code: "invalid_document_citation" });
  }
  if (original && tag !== original.tag) {
    throw new OfficeKitCodecError(`Imported document citation ${block.id} source tag is source-bound.`, [], { code: "unsupported_document_edit" });
  }
  const display = String(block.text ?? "");
  if (display.length > 1_000_000) {
    throw new OfficeKitCodecError(`Document citation ${block.id} display text exceeds 1,000,000 characters.`, [], { code: "invalid_document_citation" });
  }
  return { tag, display };
}

function documentNoteSnapshot(note) {
  return {
    id: note.id,
    kind: note.kind,
    targetId: note.targetId,
    paragraphs: note.paragraphs,
    text: note.text,
    nativeId: note.nativeId,
  };
}

function wireDocumentNoteKind(value) {
  if (value === "footnote") return DocumentNoteKind.FOOTNOTE;
  if (value === "endnote") return DocumentNoteKind.ENDNOTE;
  return DocumentNoteKind.UNSPECIFIED;
}

function publicDocumentNoteKind(value) {
  if (value === DocumentNoteKind.FOOTNOTE) return "footnote";
  if (value === DocumentNoteKind.ENDNOTE) return "endnote";
  throw new OfficeKitCodecError(`Document note kind ${value} is invalid.`, [], { code: "invalid_document_note" });
}

function documentNote(note, slot, document) {
  const kind = String(note.kind || "");
  const targetId = String(note.targetId || "");
  const body = documentNoteBody(note);
  const { paragraphs, text } = body;
  const nativeId = note.nativeId === undefined ? "" : String(note.nativeId);
  if (slot) {
    const original = slot.publicSnapshot;
    if (note.id !== original.id || kind !== original.kind || targetId !== original.targetId || note.nativeId !== original.nativeId) {
      throw new OfficeKitCodecError(`Imported document ${kind || "note"} ${note.id} identity, kind, target, and native ID are source-bound.`, [], { code: "unsupported_document_note_edit" });
    }
    if (JSON.stringify(paragraphs) === JSON.stringify(original.paragraphs)) return slot.wire;
    if (paragraphs.length !== original.paragraphs.length) {
      throw new OfficeKitCodecError(`Imported document ${kind} ${note.id} paragraph count is source-bound.`, [], { code: "unsupported_document_note_edit" });
    }
    if (slot.wire.source?.editable !== true) {
      throw new OfficeKitCodecError(`Imported document ${kind} ${note.id} body topology is preserved but not editable.`, [], { code: "unsupported_document_note_edit" });
    }
    return { ...slot.wire, text, paragraphs };
  }
  if (!new Set(["footnote", "endnote"]).has(kind)) {
    throw new OfficeKitCodecError(`Document note ${note.id} kind must be footnote or endnote.`, [], { code: "invalid_document_note" });
  }
  const target = document.blocks.find((block) => block.id === targetId);
  if (!target || !new Set(["paragraph", "listItem"]).has(target.kind)) {
    throw new OfficeKitCodecError(`Document ${kind} ${note.id} target must be a paragraph or list item.`, [], { code: "invalid_document_note" });
  }
  if (nativeId && (!/^\d+$/.test(nativeId) || Number(nativeId) < 1 || Number(nativeId) > 2_147_483_647)) {
    throw new OfficeKitCodecError(`Document ${kind} ${note.id} nativeId must be a positive 32-bit integer when present.`, [], { code: "invalid_document_note" });
  }
  return {
    id: String(note.id || ""),
    kind: wireDocumentNoteKind(kind),
    targetBlockId: targetId,
    text,
    nativeId,
    paragraphs: body.explicit ? paragraphs : undefined,
  };
}

function documentNoteBody(note) {
  const paragraphs = Array.isArray(note.paragraphs)
    ? note.paragraphs.map((paragraph) => String(paragraph ?? ""))
    : [String(note.text ?? "")];
  const text = paragraphs.join("\n");
  const explicit = note._paragraphsExplicit === true;
  validateDocumentNoteText(note, text, paragraphs, explicit);
  return { paragraphs, text, explicit };
}

function validateDocumentNoteText(note, text, paragraphs = [text], explicit = false) {
  if (!text.length || text.length > 1_000_000 || /[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/.test(text)) {
    throw new OfficeKitCodecError(`Document ${note.kind || "note"} ${note.id} text must contain 1 through 1,000,000 XML-safe characters.`, [], { code: "invalid_document_note" });
  }
  if (!explicit) {
    if (/[\r\n]/.test(text)) throw new OfficeKitCodecError(`Document ${note.kind || "note"} ${note.id} must use paragraphs for a multi-paragraph body.`, [], { code: "invalid_document_note" });
    return;
  }
  if (paragraphs.length < 1 || paragraphs.length > 16) {
    throw new OfficeKitCodecError(`Document ${note.kind || "note"} ${note.id} must contain 1 through 16 canonical note paragraphs.`, [], { code: "invalid_document_note" });
  }
  for (const [index, paragraph] of paragraphs.entries()) {
    if (!paragraph.length || paragraph.length > 1_000_000 || /[\r\n\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/.test(paragraph)) {
      throw new OfficeKitCodecError(`Document ${note.kind || "note"} ${note.id} paragraph ${index + 1} must contain 1 through 1,000,000 XML-safe characters without a line break.`, [], { code: "invalid_document_note" });
    }
  }
}

function documentBookmark(bookmark, slot, document) {
  if (slot) {
    if (JSON.stringify(documentBookmarkSnapshot(bookmark)) !== JSON.stringify(slot.publicSnapshot)) {
      throw new OfficeKitCodecError(`Imported document bookmark ${bookmark.id} identity, name, and target are source-bound in protocol 2.`, [], { code: "unsupported_document_bookmark_edit" });
    }
    return slot.wire;
  }
  const name = String(bookmark.name || "");
  if (!/^[A-Za-z][A-Za-z0-9_]{0,39}$/.test(name)) {
    throw new OfficeKitCodecError(`Document bookmark ${bookmark.id} name must start with an ASCII letter and contain only letters, digits, or underscores (maximum 40 characters).`, [], { code: "invalid_document_bookmark" });
  }
  if (!bookmark.targetId || bookmark.targetId !== bookmark.endTargetId) {
    throw new OfficeKitCodecError(`Document bookmark ${bookmark.id} must wrap exactly one block in protocol 2.`, [], { code: "invalid_document_bookmark" });
  }
  const target = document.blocks.find((block) => block.id === bookmark.targetId);
  if (!target || !new Set(["paragraph", "hyperlink", "field", "citation", "change", "image"]).has(target.kind)) {
    throw new OfficeKitCodecError(`Document bookmark ${bookmark.id} target must be a paragraph, hyperlink, field, citation, tracked change, or image block.`, [], { code: "invalid_document_bookmark" });
  }
  let nativeId = "";
  if (bookmark.nativeId !== undefined) {
    const value = Number(bookmark.nativeId);
    if (!Number.isInteger(value) || value < 0 || value > 4_294_967_295) {
      throw new OfficeKitCodecError(`Document bookmark ${bookmark.id} nativeId must be an unsigned 32-bit integer when present.`, [], { code: "invalid_document_bookmark" });
    }
    nativeId = String(value);
  }
  return {
    id: String(bookmark.id || ""),
    name,
    targetBlockId: bookmark.targetId,
    endTargetBlockId: bookmark.endTargetId,
    nativeId,
  };
}

function documentComment(comment, slot) {
  if (slot && (comment.id !== slot.wire.id || comment.targetId !== slot.wire.targetBlockId)) {
    throw new OfficeKitCodecError(`Document comment ${comment.id} identity and target are source-bound.`, [], { code: "unsupported_document_comment_edit" });
  }
  if (slot) {
    const immutable = {
      parentId: slot.wire.parentCommentId || undefined,
      paraId: slot.wire.paragraphId || undefined,
      durableId: slot.wire.durableId || undefined,
      dateUtc: slot.wire.dateUtc,
      person: slot.wire.person ? { providerId: slot.wire.person.providerId, userId: slot.wire.person.userId } : undefined,
      intelligentPlaceholder: Boolean(slot.wire.intelligentPlaceholder),
    };
    const requested = {
      parentId: comment.parentId || undefined,
      paraId: comment.paraId || undefined,
      durableId: comment.durableId || undefined,
      dateUtc: comment.dateUtc,
      person: comment.person ? { providerId: String(comment.person.providerId ?? ""), userId: String(comment.person.userId ?? "") } : undefined,
      intelligentPlaceholder: Boolean(comment.intelligentPlaceholder),
    };
    if (JSON.stringify(requested) !== JSON.stringify(immutable)) {
      throw new OfficeKitCodecError(`Document comment ${comment.id} parent, paragraph/durable identity, UTC/person metadata, and intelligent-placeholder state are source-bound.`, [], { code: "unsupported_document_comment_edit" });
    }
  }
  if (slot && JSON.stringify(documentCommentSnapshot(comment)) === JSON.stringify(slot.publicSnapshot)) return slot.wire;
  const author = String(comment.author ?? "");
  const initials = slot && comment.initials === slot.publicSnapshot.initials
    ? slot.wire.initials
    : comment.initials == null ? undefined : String(comment.initials);
  const text = String(comment.text ?? "");
  if (!author || author.length > 255) throw new OfficeKitCodecError(`Document comment ${comment.id} author must contain 1 through 255 characters.`, [], { code: "invalid_document_comment" });
  if (initials !== undefined && (!initials || initials.length > 9)) throw new OfficeKitCodecError(`Document comment ${comment.id} initials must contain 1 through 9 characters when present.`, [], { code: "invalid_document_comment" });
  if (text.length > 1_000_000) throw new OfficeKitCodecError(`Document comment ${comment.id} text exceeds 1,000,000 characters.`, [], { code: "invalid_document_comment" });
  let createdAt;
  if (comment.date != null) {
    createdAt = String(comment.date);
    if (createdAt.length > 64 || Number.isNaN(Date.parse(createdAt))) throw new OfficeKitCodecError(`Document comment ${comment.id} date must be an ISO 8601 date-time of at most 64 characters.`, [], { code: "invalid_document_comment" });
  }
  const modern = Boolean(comment.parentId || comment.paraId || comment.durableId || comment.dateUtc || comment.person || comment.intelligentPlaceholder || comment._resolvedSpecified || slot?.wire.resolved !== undefined);
  return {
    id: slot?.wire.id || comment.id,
    targetBlockId: comment.targetId,
    author,
    text,
    initials,
    createdAt,
    source: slot?.wire.source,
    parentCommentId: comment.parentId || "",
    resolved: modern ? Boolean(comment.resolved) : undefined,
    paragraphId: comment.paraId ? String(comment.paraId).toUpperCase() : "",
    durableId: comment.durableId ? String(comment.durableId).toUpperCase() : "",
    dateUtc: comment.dateUtc == null ? undefined : String(comment.dateUtc),
    person: comment.person ? { providerId: String(comment.person.providerId), userId: String(comment.person.userId) } : undefined,
    intelligentPlaceholder: modern && comment.intelligentPlaceholder ? true : undefined,
  };
}

function documentStyleType(value) {
  if (value === "character") return DocumentStyleType.CHARACTER;
  if (value === "table") return DocumentStyleType.TABLE;
  return DocumentStyleType.PARAGRAPH;
}

function publicDocumentStyleType(value) {
  if (value === DocumentStyleType.CHARACTER) return "character";
  if (value === DocumentStyleType.TABLE) return "table";
  return "paragraph";
}

function wireDocumentStyle(style) {
  const runSource = Object.fromEntries([...DOCUMENT_RUN_STYLE_KEYS].filter((key) => key !== "runStyleId" && Object.hasOwn(style, key)).map((key) => [key, style[key]]));
  return {
    id: String(style.id || ""),
    name: String(style.name || style.id || ""),
    type: documentStyleType(style.type),
    basedOn: String(style.basedOn || style.parent || style.extends || ""),
    runFormat: documentRunFormatting(runSource, `Document style ${style.id || "(unnamed)"}`),
    paragraphFormat: documentParagraphFormatting({ id: style.id || "(unnamed)", paragraphFormat: style.paragraphFormat || style }),
  };
}

function publicDocumentStyle(style) {
  return {
    id: style.id,
    name: style.name || style.id,
    type: publicDocumentStyleType(style.type),
    ...(style.basedOn ? { basedOn: style.basedOn } : {}),
    ...publicDocumentRunFormatting(style.runFormat),
    ...(publicDocumentParagraphFormatting(style.paragraphFormat) || {}),
  };
}

function headerFooterReference(value) {
  if (value === "first") return DocumentHeaderFooterReference.FIRST;
  if (value === "even") return DocumentHeaderFooterReference.EVEN;
  return DocumentHeaderFooterReference.DEFAULT;
}

function publicHeaderFooterReference(value) {
  if (value === DocumentHeaderFooterReference.FIRST) return "first";
  if (value === DocumentHeaderFooterReference.EVEN) return "even";
  return "default";
}

function documentHeaderFooterSnapshot(block) {
  return {
    id: String(block.id || ""),
    name: String(block.name || block.kind || ""),
    styleId: String(block.styleId || "Normal"),
    text: String(block.text || ""),
    referenceType: block.referenceType === "first" || block.referenceType === "even" ? block.referenceType : "default",
    sectionIndex: block.sectionIndex == null ? undefined : Number(block.sectionIndex),
    relationshipId: String(block.relationshipId || ""),
    partPath: String(block.partPath || ""),
    variantActive: block.variantActive == null ? undefined : Boolean(block.variantActive),
    fieldInstruction: String(block.fieldInstruction || block.field || ""),
    segments: Array.isArray(block.segments) ? structuredClone(block.segments) : [],
  };
}

function wireDocumentHeaderFooterSegments(segments, label) {
  if (!segments.length) return [];
  if (segments.length < 2 || segments.length > 32) {
    throw new OfficeKitCodecError(`Document ${label} structured segments require 2 through 32 items.`, [], { code: "invalid_document_header_footer" });
  }
  let display = "";
  let fieldCount = 0;
  const wire = segments.map((segment, index) => {
    if (!segment || typeof segment !== "object" || Array.isArray(segment)) {
      throw new OfficeKitCodecError(`Document ${label} structured segment ${index + 1} must be an object.`, [], { code: "invalid_document_header_footer" });
    }
    const keys = Object.keys(segment);
    if (keys.length !== 1 || (keys[0] !== "text" && keys[0] !== "field")) {
      throw new OfficeKitCodecError(`Document ${label} structured segment ${index + 1} must contain exactly one text or field property.`, [], { code: "invalid_document_header_footer" });
    }
    if (keys[0] === "text") {
      const text = String(segment.text ?? "");
      if (!text || text.length > 1_000_000 || !isXmlSafeText(text)) {
        throw new OfficeKitCodecError(`Document ${label} structured text segment ${index + 1} is invalid.`, [], { code: "invalid_document_header_footer" });
      }
      display += text;
      return { content: { case: "text", value: text } };
    }
    const field = segment.field;
    if (!field || typeof field !== "object" || Array.isArray(field) || Object.keys(field).some((key) => key !== "instruction" && key !== "display")) {
      throw new OfficeKitCodecError(`Document ${label} structured field segment ${index + 1} must define instruction and display.`, [], { code: "invalid_document_header_footer" });
    }
    const instruction = String(field.instruction ?? "").trim();
    const command = instruction.split(/\s+/, 1)[0]?.toUpperCase();
    const fieldDisplay = String(field.display ?? "");
    if (!instruction || instruction.length > 8192 || /[\u0000-\u001f\u007f]/.test(instruction) || !DOCUMENT_HEADER_FOOTER_FIELD_COMMANDS.has(command) || fieldDisplay.length > 1_000_000 || !isXmlSafeText(fieldDisplay)) {
      throw new OfficeKitCodecError(`Document ${label} structured field segment ${index + 1} is outside the bounded simple-field profile.`, [], { code: "invalid_document_header_footer" });
    }
    fieldCount += 1;
    display += fieldDisplay;
    return { content: { case: "field", value: { instruction, display: fieldDisplay, complex: false } } };
  });
  if (!fieldCount || display.length > 1_000_000) {
    throw new OfficeKitCodecError(`Document ${label} structured segments require a field and at most 1,000,000 display characters.`, [], { code: "invalid_document_header_footer" });
  }
  return wire;
}

function publicDocumentHeaderFooterSegments(segments = []) {
  if (!segments.length) return undefined;
  return segments.map((segment) => {
    if (segment.content?.case === "text") return { text: segment.content.value };
    if (segment.content?.case === "field") {
      const field = segment.content.value;
      return { field: { instruction: field.instruction, display: field.display } };
    }
    throw new OfficeKitCodecError("Imported DOCX header/footer contains an unsupported structured segment.", [], { code: "unsupported_document_header_footer_preserved" });
  });
}

function documentHeaderFooterSegmentsEqual(left = [], right = []) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function wireHeaderFooter(block, slot) {
  const snapshot = documentHeaderFooterSnapshot(block);
  const segments = wireDocumentHeaderFooterSegments(snapshot.segments, `${block.kind} ${block.id}`);
  const instruction = segments.length ? "" : snapshot.fieldInstruction;
  if (segments.length && snapshot.text !== segments.map((segment) => segment.content.case === "text" ? segment.content.value : segment.content.value.display).join("")) {
    throw new OfficeKitCodecError(`Document ${block.kind} ${block.id} structured segment display must exactly match text.`, [], { code: "invalid_document_header_footer" });
  }
  if (instruction && !DOCUMENT_HEADER_FOOTER_FIELD_COMMANDS.has(instruction.trim().split(/\s+/)[0].toUpperCase())) throw new OfficeKitCodecError(`Document ${block.kind} ${block.id} uses unsupported field ${instruction}.`, [], { code: "invalid_document_field" });
  if (slot) {
    const source = slot.publicSnapshot;
    if (snapshot.id !== source.id || snapshot.name !== source.name || snapshot.styleId !== source.styleId ||
      snapshot.referenceType !== source.referenceType || snapshot.sectionIndex !== source.sectionIndex ||
      snapshot.relationshipId !== source.relationshipId || snapshot.partPath !== source.partPath ||
      snapshot.variantActive !== source.variantActive || snapshot.fieldInstruction !== source.fieldInstruction ||
      !documentHeaderFooterSegmentsEqual(snapshot.segments, source.segments)) {
      throw new OfficeKitCodecError(`Imported document ${block.kind} ${snapshot.id} has fixed source identity, section scope, style, and field topology.`, [], { code: "unsupported_document_header_footer_edit" });
    }
    if (snapshot.text !== source.text && slot.wire.source?.editable !== true) {
      throw new OfficeKitCodecError(`Imported document ${block.kind} ${snapshot.id} is source-bound and cannot replace its text in this codec profile.`, [], { code: "unsupported_document_header_footer_edit" });
    }
  }
  return {
    id: snapshot.id,
    name: snapshot.name,
    styleId: snapshot.styleId,
    text: snapshot.text,
    reference: headerFooterReference(snapshot.referenceType),
    sectionIndex: snapshot.sectionIndex == null ? undefined : uint32(snapshot.sectionIndex, `Document ${block.kind} ${block.id} sectionIndex`),
    relationshipId: snapshot.relationshipId,
    partPath: snapshot.partPath,
    variantActive: snapshot.variantActive,
    fieldInstruction: instruction,
    segments,
    source: slot?.wire.source,
  };
}

function publicHeaderFooter(block) {
  return {
    id: block.id || undefined,
    name: block.name || undefined,
    styleId: block.styleId || "Normal",
    text: block.text || "",
    referenceType: publicHeaderFooterReference(block.reference),
    sectionIndex: block.sectionIndex,
    relationshipId: block.relationshipId || undefined,
    partPath: block.partPath || undefined,
    variantActive: block.variantActive,
    fieldInstruction: block.fieldInstruction || undefined,
    segments: publicDocumentHeaderFooterSegments(block.segments || []),
    sourceBound: Boolean(block.source),
    editable: block.source ? block.source.editable === true : undefined,
  };
}

function wireDocumentHeaderFooters(blocks, slots, kind) {
  if (!slots) return blocks.map((block) => wireHeaderFooter(block));
  if (blocks.length !== slots.length) {
    throw new OfficeKitCodecError(`Source-preserving DOCX export requires the original ${kind} topology.`, [], { code: "document_header_footer_topology_changed" });
  }
  return blocks.map((block, index) => wireHeaderFooter(block, slots[index]));
}

function documentWatermarkSnapshot(watermark) {
  return {
    id: watermark.id,
    text: watermark.text,
    referenceType: watermark.referenceType,
    sectionIndex: watermark.sectionIndex,
  };
}

function wireDocumentWatermark(watermark, slot) {
  const snapshot = documentWatermarkSnapshot(watermark);
  if (!snapshot.id || snapshot.id.length > 512 || /[\u0000-\u001f\u007f]/.test(snapshot.id)) {
    throw new OfficeKitCodecError("Document watermarks require unique IDs of 1 through 512 characters without controls.", [], { code: "invalid_document_watermark" });
  }
  if (!snapshot.text.trim() || snapshot.text.length > 256 || !isXmlSafeText(snapshot.text)) {
    throw new OfficeKitCodecError(`Document watermark ${snapshot.id} text must contain 1 through 256 XML-safe characters and cannot be blank.`, [], { code: "invalid_document_watermark" });
  }
  if (slot) {
    if (snapshot.id !== slot.wire.id || snapshot.referenceType !== publicHeaderFooterReference(slot.wire.reference) || snapshot.sectionIndex !== slot.wire.sectionIndex) {
      throw new OfficeKitCodecError(`Document watermark ${snapshot.id} source identity, section, and header reference are fixed after import.`, [], { code: "unsupported_document_watermark_edit" });
    }
    if (snapshot.text !== slot.publicSnapshot.text && slot.wire.source?.editable !== true) {
      throw new OfficeKitCodecError(`Document watermark ${snapshot.id} is source-bound and read-only.`, [], { code: "unsupported_document_watermark_edit" });
    }
  }
  return {
    id: snapshot.id,
    text: snapshot.text,
    reference: headerFooterReference(snapshot.referenceType),
    sectionIndex: uint32(snapshot.sectionIndex, `Document watermark ${snapshot.id} sectionIndex`),
    source: slot?.wire.source,
  };
}

function wireDocumentWatermarks(document, state) {
  if (!state) return document.watermarks.map((watermark) => wireDocumentWatermark(watermark));
  const slots = state.watermarkSlots || [];
  const byId = new Map(slots.map((slot, index) => [slot.wire.id, { slot, index }]));
  let previousIndex = -1;
  const retained = document.watermarks.map((watermark) => {
    const match = byId.get(watermark.id);
    if (!match) {
      throw new OfficeKitCodecError("Source-preserving DOCX export cannot add a watermark to an imported document; only recognized existing watermarks may be edited or removed.", [], { code: "document_watermark_topology_changed" });
    }
    if (match.index <= previousIndex) {
      throw new OfficeKitCodecError("Source-preserving DOCX export cannot reorder imported watermarks.", [], { code: "document_watermark_topology_changed" });
    }
    previousIndex = match.index;
    return wireDocumentWatermark(watermark, match.slot);
  });
  const retainedIds = new Set(retained.map((watermark) => watermark.id));
  for (const slot of slots) {
    if (!retainedIds.has(slot.wire.id) && slot.wire.source?.editable !== true) {
      throw new OfficeKitCodecError(`Document watermark ${slot.wire.id} is source-bound and cannot be removed.`, [], { code: "unsupported_document_watermark_edit" });
    }
  }
  return retained;
}

function documentSectionBreak(value) {
  if (value === "continuous") return DocumentSectionBreak.CONTINUOUS;
  if (value === "evenPage") return DocumentSectionBreak.EVEN_PAGE;
  if (value === "oddPage") return DocumentSectionBreak.ODD_PAGE;
  return DocumentSectionBreak.NEXT_PAGE;
}

function publicDocumentSectionBreak(value) {
  if (value === DocumentSectionBreak.CONTINUOUS) return "continuous";
  if (value === DocumentSectionBreak.EVEN_PAGE) return "evenPage";
  if (value === DocumentSectionBreak.ODD_PAGE) return "oddPage";
  return "nextPage";
}

function documentSectionPageNumberFormat(value) {
  if (value === "decimal") return DocumentSectionPageNumberFormat.DECIMAL;
  if (value === "upperRoman") return DocumentSectionPageNumberFormat.UPPER_ROMAN;
  if (value === "lowerRoman") return DocumentSectionPageNumberFormat.LOWER_ROMAN;
  if (value === "upperLetter") return DocumentSectionPageNumberFormat.UPPER_LETTER;
  if (value === "lowerLetter") return DocumentSectionPageNumberFormat.LOWER_LETTER;
  throw new TypeError(`Unsupported document section page-number format ${value || "(empty)"}.`);
}

function publicDocumentSectionPageNumberFormat(value) {
  if (value === DocumentSectionPageNumberFormat.DECIMAL) return "decimal";
  if (value === DocumentSectionPageNumberFormat.UPPER_ROMAN) return "upperRoman";
  if (value === DocumentSectionPageNumberFormat.LOWER_ROMAN) return "lowerRoman";
  if (value === DocumentSectionPageNumberFormat.UPPER_LETTER) return "upperLetter";
  if (value === DocumentSectionPageNumberFormat.LOWER_LETTER) return "lowerLetter";
  throw new OfficeKitCodecError("OfficeKit returned an unsupported document section page-number format.", [], { code: "invalid_document_section" });
}

function publicDocumentSectionPageNumbering(value) {
  if (!value) return undefined;
  return {
    ...(value.start === undefined ? {} : { start: value.start }),
    ...(value.format === DocumentSectionPageNumberFormat.UNSPECIFIED ? {} : { format: publicDocumentSectionPageNumberFormat(value.format) }),
  };
}

function documentSectionLineNumberRestart(value) {
  if (value === "newPage") return DocumentSectionLineNumberRestart.NEW_PAGE;
  if (value === "newSection") return DocumentSectionLineNumberRestart.NEW_SECTION;
  if (value === "continuous") return DocumentSectionLineNumberRestart.CONTINUOUS;
  throw new TypeError(`Unsupported document section line-number restart ${value || "(empty)"}.`);
}

function publicDocumentSectionLineNumberRestart(value) {
  if (value === DocumentSectionLineNumberRestart.NEW_PAGE) return "newPage";
  if (value === DocumentSectionLineNumberRestart.NEW_SECTION) return "newSection";
  if (value === DocumentSectionLineNumberRestart.CONTINUOUS) return "continuous";
  throw new OfficeKitCodecError("OfficeKit returned an unsupported document section line-number restart.", [], { code: "invalid_document_section" });
}

function publicDocumentSectionLineNumbering(value) {
  if (!value) return undefined;
  return {
    countBy: value.countBy,
    ...(value.start === undefined ? {} : { start: value.start }),
    ...(value.distanceTwips === undefined ? {} : { distance: value.distanceTwips }),
    ...(value.restart === DocumentSectionLineNumberRestart.UNSPECIFIED ? {} : { restart: publicDocumentSectionLineNumberRestart(value.restart) }),
  };
}

function documentChangeType(value) {
  if (value === "insert") return DocumentChangeType.INSERT;
  if (value === "delete") return DocumentChangeType.DELETE;
  throw new OfficeKitCodecError(`Document tracked-change type ${value || "(empty)"} must be insert or delete.`, [], { code: "invalid_document_change" });
}

function publicDocumentChangeType(value) {
  if (value === DocumentChangeType.INSERT) return "insert";
  if (value === DocumentChangeType.DELETE) return "delete";
  throw new OfficeKitCodecError("Document tracked-change wire type must be insert or delete.", [], { code: "invalid_document_change" });
}

function wireDocumentChange(block) {
  const text = String(block.text ?? "");
  const author = String(block.author ?? "");
  const date = block.date == null || block.date === "" ? undefined : String(block.date);
  if (text.length > 1_000_000) throw new OfficeKitCodecError(`Document tracked change ${block.id} text exceeds 1,000,000 characters.`, [], { code: "invalid_document_change" });
  if (!author.trim() || author.length > 255 || /[\u0000-\u001f\u007f]/.test(author)) throw new OfficeKitCodecError(`Document tracked change ${block.id} requires an author of at most 255 characters without controls.`, [], { code: "invalid_document_change" });
  if (date !== undefined && Number.isNaN(Date.parse(date))) throw new OfficeKitCodecError(`Document tracked change ${block.id} date must be an ISO 8601 timestamp.`, [], { code: "invalid_document_change" });
  return { type: documentChangeType(block.changeType), text, author, date };
}

function documentImage(block, assets) {
  if (!block.dataUrl) throw new OfficeKitCodecError(`Document image ${block.id} requires embedded PNG or JPEG data.`, [], { code: "unsupported_document_image" });
  const match = /^data:(image\/(?:png|jpeg));base64,([A-Za-z0-9+/=\s]+)$/i.exec(String(block.dataUrl));
  if (!match) throw new OfficeKitCodecError(`Document image ${block.id} must use a base64 PNG/JPEG data URL.`, [], { code: "unsupported_document_image" });
  const bytes = new Uint8Array(Buffer.from(match[2].replace(/\s/g, ""), "base64"));
  if (!bytes.length) throw new OfficeKitCodecError(`Document image ${block.id} contains no image bytes.`, [], { code: "invalid_document_image" });
  const contentType = match[1].toLowerCase();
  const sha256 = createHash("sha256").update(bytes).digest("hex");
  const assetId = `asset/document/image/${sha256}`;
  assets.set(assetId, { id: assetId, fileName: `${sha256}.${contentType === "image/png" ? "png" : "jpg"}`, contentType, data: bytes, sha256 });
  const widthEmu = Math.round(Number(block.widthPx) * 9_525);
  const heightEmu = Math.round(Number(block.heightPx) * 9_525);
  if (!Number.isSafeInteger(widthEmu) || !Number.isSafeInteger(heightEmu) || widthEmu <= 0 || heightEmu <= 0) throw new OfficeKitCodecError(`Document image ${block.id} dimensions must be positive bounded pixels.`, [], { code: "invalid_document_image" });
  return { assetId, altText: String(block.alt ?? ""), widthEmu, heightEmu, floating: wireDocumentFloatingImagePlacement(block) };
}

function wireDocumentImageHorizontalReference(value) {
  if (value === "margin") return DocumentImageHorizontalRelativeFrom.MARGIN;
  if (value === "page") return DocumentImageHorizontalRelativeFrom.PAGE;
  if (value === "column") return DocumentImageHorizontalRelativeFrom.COLUMN;
  throw new OfficeKitCodecError(`Document floating image horizontal relativeTo ${value || "(empty)"} is unsupported.`, [], { code: "invalid_document_image" });
}

function wireDocumentImageVerticalReference(value) {
  if (value === "margin") return DocumentImageVerticalRelativeFrom.MARGIN;
  if (value === "page") return DocumentImageVerticalRelativeFrom.PAGE;
  if (value === "paragraph") return DocumentImageVerticalRelativeFrom.PARAGRAPH;
  throw new OfficeKitCodecError(`Document floating image vertical relativeTo ${value || "(empty)"} is unsupported.`, [], { code: "invalid_document_image" });
}

function wireDocumentImageWrap(value) {
  if (value === "square") return DocumentImageWrapMode.SQUARE;
  if (value === "topAndBottom") return DocumentImageWrapMode.TOP_AND_BOTTOM;
  throw new OfficeKitCodecError(`Document floating image wrap ${value || "(empty)"} is unsupported.`, [], { code: "invalid_document_image" });
}

function wireDocumentImageWrapSide(value) {
  if (value === undefined) return DocumentImageWrapSide.UNSPECIFIED;
  if (value === "bothSides") return DocumentImageWrapSide.BOTH_SIDES;
  if (value === "left") return DocumentImageWrapSide.LEFT;
  if (value === "right") return DocumentImageWrapSide.RIGHT;
  if (value === "largest") return DocumentImageWrapSide.LARGEST;
  throw new OfficeKitCodecError(`Document floating image wrapSide ${value || "(empty)"} is unsupported.`, [], { code: "invalid_document_image" });
}

function documentImageEmu(value, label, { unsigned = false } = {}) {
  const number = Number(value);
  const emu = Math.round(number * 9_525);
  if (!Number.isFinite(number) || !Number.isSafeInteger(emu) || (unsigned ? emu < 0 || emu > 0xffff_ffff : Math.abs(emu) > 95_250_000)) {
    throw new OfficeKitCodecError(`${label} is outside the bounded pixel range.`, [], { code: "invalid_document_image" });
  }
  return emu;
}

function wireDocumentFloatingImagePlacement(block) {
  const placement = typeof block?.toProto === "function" ? block.toProto().placement : block?.placement;
  if (!placement) return undefined;
  if (placement.type !== "floating") throw new OfficeKitCodecError(`Document image ${block.id} placement must be floating when present.`, [], { code: "invalid_document_image" });
  const distance = placement.distanceFromTextPx || {};
  return {
    horizontalRelativeFrom: wireDocumentImageHorizontalReference(placement.horizontal?.relativeTo),
    horizontalOffsetEmu: documentImageEmu(placement.horizontal?.offsetPx, `Document image ${block.id} horizontal offset`),
    verticalRelativeFrom: wireDocumentImageVerticalReference(placement.vertical?.relativeTo),
    verticalOffsetEmu: documentImageEmu(placement.vertical?.offsetPx, `Document image ${block.id} vertical offset`),
    wrapMode: wireDocumentImageWrap(placement.wrap),
    wrapSide: wireDocumentImageWrapSide(placement.wrapSide),
    distanceTopEmu: documentImageEmu(distance.top, `Document image ${block.id} top text distance`, { unsigned: true }),
    distanceRightEmu: documentImageEmu(distance.right, `Document image ${block.id} right text distance`, { unsigned: true }),
    distanceBottomEmu: documentImageEmu(distance.bottom, `Document image ${block.id} bottom text distance`, { unsigned: true }),
    distanceLeftEmu: documentImageEmu(distance.left, `Document image ${block.id} left text distance`, { unsigned: true }),
  };
}

function publicDocumentImageHorizontalReference(value) {
  if (value === DocumentImageHorizontalRelativeFrom.MARGIN) return "margin";
  if (value === DocumentImageHorizontalRelativeFrom.PAGE) return "page";
  if (value === DocumentImageHorizontalRelativeFrom.COLUMN) return "column";
  throw new OfficeKitCodecError("OfficeKit returned an unsupported document floating-image horizontal reference.", [], { code: "invalid_document_image" });
}

function publicDocumentImageVerticalReference(value) {
  if (value === DocumentImageVerticalRelativeFrom.MARGIN) return "margin";
  if (value === DocumentImageVerticalRelativeFrom.PAGE) return "page";
  if (value === DocumentImageVerticalRelativeFrom.PARAGRAPH) return "paragraph";
  throw new OfficeKitCodecError("OfficeKit returned an unsupported document floating-image vertical reference.", [], { code: "invalid_document_image" });
}

function publicDocumentImageWrap(value) {
  if (value === DocumentImageWrapMode.SQUARE) return "square";
  if (value === DocumentImageWrapMode.TOP_AND_BOTTOM) return "topAndBottom";
  throw new OfficeKitCodecError("OfficeKit returned an unsupported document floating-image wrap mode.", [], { code: "invalid_document_image" });
}

function publicDocumentImageWrapSide(value, mode) {
  if (mode === DocumentImageWrapMode.TOP_AND_BOTTOM && value === DocumentImageWrapSide.UNSPECIFIED) return undefined;
  if (value === DocumentImageWrapSide.BOTH_SIDES) return "bothSides";
  if (value === DocumentImageWrapSide.LEFT) return "left";
  if (value === DocumentImageWrapSide.RIGHT) return "right";
  if (value === DocumentImageWrapSide.LARGEST) return "largest";
  throw new OfficeKitCodecError("OfficeKit returned an unsupported document floating-image wrap side.", [], { code: "invalid_document_image" });
}

function publicDocumentFloatingImagePlacement(value) {
  if (!value) return undefined;
  return {
    type: "floating",
    horizontal: { relativeTo: publicDocumentImageHorizontalReference(value.horizontalRelativeFrom), offsetPx: Number(value.horizontalOffsetEmu) / 9_525 },
    vertical: { relativeTo: publicDocumentImageVerticalReference(value.verticalRelativeFrom), offsetPx: Number(value.verticalOffsetEmu) / 9_525 },
    wrap: publicDocumentImageWrap(value.wrapMode),
    wrapSide: publicDocumentImageWrapSide(value.wrapSide, value.wrapMode),
    distanceFromTextPx: {
      top: value.distanceTopEmu / 9_525,
      right: value.distanceRightEmu / 9_525,
      bottom: value.distanceBottomEmu / 9_525,
      left: value.distanceLeftEmu / 9_525,
    },
  };
}

function wireDocumentSection(block) {
  const page = block.pageSize || {};
  const margins = block.margins || {};
  const columns = block.columns;
  const pageNumbering = block.pageNumbering;
  const lineNumbering = block.lineNumbering;
  if (columns && typeof columns.separator !== "boolean") {
    throw new TypeError(`Document section ${block.id} column separator must be boolean.`);
  }
  return {
    breakType: documentSectionBreak(block.breakType),
    pageWidthTwips: uint32(Math.round(Number(page.widthTwips)), `Document section ${block.id} page width`),
    pageHeightTwips: uint32(Math.round(Number(page.heightTwips)), `Document section ${block.id} page height`),
    landscape: block.orientation === "landscape",
    marginTopTwips: uint32(Math.round(Number(margins.top)), `Document section ${block.id} top margin`),
    marginRightTwips: uint32(Math.round(Number(margins.right)), `Document section ${block.id} right margin`),
    marginBottomTwips: uint32(Math.round(Number(margins.bottom)), `Document section ${block.id} bottom margin`),
    marginLeftTwips: uint32(Math.round(Number(margins.left)), `Document section ${block.id} left margin`),
    marginGutterTwips: uint32(Math.round(Number(margins.gutter ?? 0)), `Document section ${block.id} gutter margin`),
    columns: columns ? (Object.hasOwn(columns, "definitions") ? {
      separator: columns.separator,
      definitions: columns.definitions.map((definition, index) => ({
        widthTwips: uint32(Number(definition?.width), `Document section ${block.id} custom column ${index} width`),
        spacingAfterTwips: uint32(Number(definition?.spacing), `Document section ${block.id} custom column ${index} spacing`),
      })),
    } : {
      count: uint32(Number(columns.count), `Document section ${block.id} column count`),
      spacingTwips: uint32(Number(columns.spacing), `Document section ${block.id} column spacing`),
      separator: columns.separator,
    }) : undefined,
    pageNumbering: pageNumbering ? {
      ...(Object.hasOwn(pageNumbering, "start") ? { start: uint32(Number(pageNumbering.start), `Document section ${block.id} page-number start`) } : {}),
      ...(Object.hasOwn(pageNumbering, "format") ? { format: documentSectionPageNumberFormat(pageNumbering.format) } : {}),
    } : undefined,
    lineNumbering: lineNumbering ? {
      countBy: uint32(Number(lineNumbering.countBy), `Document section ${block.id} line-number countBy`),
      ...(Object.hasOwn(lineNumbering, "start") ? { start: uint32(Number(lineNumbering.start), `Document section ${block.id} line-number start`) } : {}),
      ...(Object.hasOwn(lineNumbering, "distance") ? { distanceTwips: uint32(Number(lineNumbering.distance), `Document section ${block.id} line-number distance`) } : {}),
      ...(Object.hasOwn(lineNumbering, "restart") ? { restart: documentSectionLineNumberRestart(lineNumbering.restart) } : {}),
    } : undefined,
  };
}

function unchangedSourceBlock(block, original, assets) {
  switch (original.content.case) {
    case "paragraph": {
      if (original.content.value.numbering) {
        return block.styleId === (original.styleId || "Normal") && sameDocumentNumbering(block, original.content.value, assets);
      }
      if (block.kind !== "paragraph" || block.text !== original.content.value.text || block.styleId !== (original.styleId || "Normal")) return false;
      if (original.source?.editable !== false) return false;
      return block.runs.every((run) => Object.keys(run.style || {}).length === 0);
    }
    case "table": {
      if (block.kind !== "table" || block.textPatches?.length || !sameTableValues(block, original) ||
          !sameDocumentTableGeometry(block, original.content.value) ||
          !sameDocumentTableContentControls(block, original.content.value) ||
          !sameDocumentTableHeaderRows(block, original.content.value) ||
          !sameDocumentTableKeepTogetherRows(block, original.content.value) ||
          !sameDocumentTableMinimumRowHeights(block, original.content.value) ||
          !sameDocumentTableAccessibility(block, original.content.value) ||
          !sameDocumentTableFormatting(block, original.content.value)) return false;
      return block.styleId === original.styleId || (!original.styleId && block.styleId === "TableGrid");
    }
    case "hyperlink":
      return sameDocumentHyperlink(block, original.content.value);
    case "field":
      return block.kind === "field" && block.styleId === (original.styleId || "Normal") && block.instruction === original.content.value.instruction && block.display === original.content.value.display && Boolean(block.complex) === Boolean(original.content.value.complex);
    case "citation":
      return block.kind === "citation" && block.styleId === (original.styleId || "Normal") &&
        String(block.metadata?.tag || "") === original.content.value.tag && block.text === original.content.value.display;
    case "change": {
      if (block.kind !== "change" || block.styleId !== (original.styleId || "Normal")) return false;
      const value = original.content.value;
      return block.changeType === publicDocumentChangeType(value.type) && block.text === value.text && block.author === value.author && (block.date || undefined) === value.date;
    }
    case "section": {
      if (block.kind !== "section") return false;
      const value = wireDocumentSection(block);
      return JSON.stringify(value) === JSON.stringify(original.content.value);
    }
    case "opaque":
      return block.kind === "paragraph" && block.text === original.content.value.text && block.runs.every((run) => Object.keys(run.style || {}).length === 0);
    default:
      return false;
  }
}

function documentBlockSnapshot(block) {
  return JSON.stringify({
    proto: typeof block?.toProto === "function" ? block.toProto() : {
      id: block?.id,
      name: block?.name,
      kind: block?.kind,
      styleId: block?.styleId,
      text: block?.text,
    },
    runs: Array.isArray(block?.runs)
      ? block.runs.map((run) => ({ text: String(run?.text ?? ""), style: { ...(run?.style || {}) }, contentControl: run?.contentControl ? { ...run.contentControl } : undefined, inlineField: run?.inlineField ? { ...run.inlineField } : undefined }))
      : undefined,
  });
}

function patchedSourceParagraphBlock(block, original) {
  if (block.kind !== "paragraph") return undefined;
  const patches = Array.isArray(block.textPatches) ? block.textPatches : [];
  if (!patches.length) return undefined;
  if (original?.content.case !== "paragraph" || original.source?.textPatchable !== true || original.source?.editable !== false) {
    throw new OfficeKitCodecError(`Document paragraph ${block.id} text patches require a non-editable imported paragraph that advertises textPatchable.`, [], { code: "unsupported_document_edit" });
  }
  if (patches.length > 10_000) throw new OfficeKitCodecError(`Document paragraph ${block.id} exceeds 10,000 source text patches.`, [], { code: "invalid_document_text_patch" });
  let expected = String(original.content.value.text ?? "");
  const sourceTextSha256 = createHash("sha256").update(expected, "utf8").digest("hex");
  const wirePatches = patches.map((patch) => {
    const search = String(patch.search ?? "");
    const replacement = String(patch.replacement ?? "");
    if (!search || search.length > 1_000_000 || replacement.length > 1_000_000 || !isXmlSafeText(search) || !isXmlSafeText(replacement)) {
      throw new OfficeKitCodecError(`Document paragraph ${block.id} text patch requires bounded XML-safe strings.`, [], { code: "invalid_document_text_patch" });
    }
    const first = expected.indexOf(search);
    if (first < 0 || expected.indexOf(search, first + 1) >= 0) {
      throw new OfficeKitCodecError(`Document paragraph ${block.id} text patch requires exactly one visible match.`, [], { code: "unsupported_document_edit" });
    }
    expected = expected.replace(search, replacement);
    return { search, replacement, sourceTextSha256 };
  });
  const baselineFormat = publicDocumentParagraphFormatting(original.content.value.formatting) || {};
  const plainSyntheticRuns = block.runs.length === (expected ? 1 : 0) && block.runs.every((run) =>
    !run.contentControl && !run.inlineField && Object.keys(run.style || {}).length === 0 && run.text === expected);
  if (block.kind !== "paragraph" || block.id !== original.id || block.name !== (original.name || "") ||
      block.styleId !== (original.styleId || "Normal") || block.text !== expected || !plainSyntheticRuns ||
      JSON.stringify(block.paragraphFormat || {}) !== JSON.stringify(baselineFormat)) {
    throw new OfficeKitCodecError(`Document paragraph ${block.id} cannot combine a native text patch with other semantic or formatting edits.`, [], { code: "unsupported_document_edit" });
  }
  const { $typeName: _typeName, ...sourceBlock } = original;
  return { ...sourceBlock, textPatches: wirePatches };
}

function documentBlock(block, original, directNumbering, assets, contentControlNativeIds) {
  const patchedParagraph = patchedSourceParagraphBlock(block, original);
  if (patchedParagraph) return patchedParagraph;
  if (original && unchangedSourceBlock(block, original, assets)) return original;
  const common = {
    id: original?.id || block.id,
    name: block.name || original?.name || "",
    styleId: block.styleId || original?.styleId || "",
    source: original?.source,
  };
  if (block.kind === "paragraph") {
    assertDocumentContentControlTopology(block, original);
    assertDocumentInlineFieldTopology(block, original);
    return {
      ...common,
      content: {
        case: "paragraph",
        value: {
          text: block.text,
          runs: block.runs.map((run) => documentRun(run, block.id, contentControlNativeIds.get(run))),
          formatting: documentParagraphFormatting(block),
          blockContentControl: block.blockContentControl
            ? wireDocumentContentControl(block.blockContentControl, contentControlNativeIds.get(block), block.id)
            : undefined,
        },
      },
    };
  }
  if (block.kind === "listItem") {
    const source = original?.content.case === "paragraph" ? original.content.value : undefined;
    if (!source?.numbering) {
      if (!directNumbering) {
        throw new OfficeKitCodecError(`The DOCX NativeAOT vertical slice could not plan a numbering-definition graph for list item ${block.id}.`, [], { code: "invalid_document_numbering" });
      }
      const text = String(block.text ?? "");
      if (text.length > 1_000_000) throw new OfficeKitCodecError(`Document list item ${block.id} text exceeds 1,000,000 characters.`, [], { code: "invalid_document_numbering" });
      return {
        ...common,
        content: { case: "paragraph", value: { text, numbering: directNumbering } },
      };
    }
    if (original.source?.editable === false) {
      throw new OfficeKitCodecError(`Document list item ${block.id} is source-preserved but its paragraph topology is not editable.`, [], { code: "unsupported_document_edit" });
    }
    const numbering = editedDocumentNumbering(block, source.numbering, assets);
    const text = String(block.text ?? "");
    if (text.length > 1_000_000) throw new OfficeKitCodecError(`Document list item ${block.id} text exceeds 1,000,000 characters.`, [], { code: "invalid_document_numbering" });
    return {
      ...common,
      content: {
        case: "paragraph",
        value: {
          text,
          runs: source.runs.map((run) => ({ ...run, text })),
          numbering,
        },
      },
    };
  }
  if (block.kind === "table") {
    const source = original?.content.case === "table" ? original.content.value : undefined;
    const authored = !source && Array.isArray(block.cells) ? authoredDocumentTableGeometry(block, contentControlNativeIds) : undefined;
    const headerRowCount = documentTableHeaderRowCount(block, source?.rows.length ?? block.values.length);
    const keepTogetherRows = documentTableKeepTogetherRows(block, source?.rows.length ?? block.values.length);
    const minimumRowHeightsDxa = documentTableMinimumRowHeights(block, source?.rows.length ?? block.values.length);
    const accessibility = documentTableAccessibility(block);
    if (source && !sameDocumentTableContentControlTopology(block, source)) {
      throw new OfficeKitCodecError(`Document table ${block.id} content-control topology is source-bound.`, [], { code: "document_content_control_topology_changed" });
    }
    if (source && !sameDocumentTableGeometry(block, source)) {
      throw new OfficeKitCodecError(`Document table ${block.id} grid, span, merge, and per-cell editability metadata are source-bound.`, [], { code: "unsupported_document_edit" });
    }
    const formattingChanged = source && !sameDocumentTableFormatting(block, source);
    if (formattingChanged && !source.formatting) {
      throw new OfficeKitCodecError(`Document table ${block.id} direct formatting can change only when OfficeKit recognized the complete bounded profile during import.`, [], { code: "unsupported_document_edit" });
    }
    if (source) {
      for (let rowIndex = 0; rowIndex < source.rows.length; rowIndex += 1) {
        for (let cellIndex = 0; cellIndex < source.rows[rowIndex].cells.length; cellIndex += 1) {
          if (String(block.values?.[rowIndex]?.[cellIndex] ?? "") !== source.rows[rowIndex].cells[cellIndex] &&
              source.rows[rowIndex].richCells[cellIndex]?.editable === false) {
            throw new OfficeKitCodecError(`Document table ${block.id} cell ${rowIndex},${cellIndex} is a vertical continuation or complex source cell and cannot be edited.`, [], { code: "unsupported_document_edit" });
          }
        }
      }
    }
    const textPatches = wireDocumentTableTextPatches(block, source);
    return {
      ...common,
      content: {
        case: "table",
        value: {
          ...(source ? { gridColumns: source.gridColumns } : authored ? { gridColumns: authored.gridColumns } : {}),
          headerRowCount,
          keepTogetherRows,
          minimumRowHeightsDxa,
          ...accessibility,
          ...(source ? (source.formatting ? {
            formatting: formattingChanged
              ? documentTableFormatting(block, source.gridColumns || Math.max(1, ...source.rows.map((row) => row.cells.length)))
              : { ...source.formatting, columnWidthsDxa: [...source.formatting.columnWidthsDxa], cellMarginsDxa: { ...source.formatting.cellMarginsDxa } },
          } : {}) : {
            formatting: documentTableFormatting(block, authored?.gridColumns || Math.max(1, block.columns)),
          }),
          rows: authored?.rows || (block.values || []).map((cells, rowIndex) => ({
            cells: cells.map((value) => String(value ?? "")),
            ...(source ? {
              richCells: source.rows[rowIndex]?.richCells.map((cell, column) => {
                const { $typeName: _typeName, ...cellValue } = cell;
                const requestedCell = block.cells?.find((candidate) => candidate.row === rowIndex && candidate.column === column);
                return {
                  ...cellValue,
                  textContentControl: requestedCell?.contentControl
                    ? wireDocumentTableCellContentControl(requestedCell.contentControl, contentControlNativeIds.get(requestedCell), `${block.id}/cell/${rowIndex}/${column}`, cells[column])
                    : undefined,
                };
              }) || [],
              gridBefore: source.rows[rowIndex]?.gridBefore || 0,
              gridAfter: source.rows[rowIndex]?.gridAfter || 0,
            } : {}),
          })),
          textPatches,
        },
      },
    };
  }
  if (block.kind === "hyperlink") {
    return {
      ...common,
      content: { case: "hyperlink", value: documentHyperlink(block, original) },
    };
  }
  if (block.kind === "field") {
    return {
      ...common,
      content: { case: "field", value: documentField(block, original) },
    };
  }
  if (block.kind === "citation") {
    return {
      ...common,
      content: { case: "citation", value: wireDocumentCitation(block, original?.content.case === "citation" ? original.content.value : undefined) },
    };
  }
  if (block.kind === "change") {
    return {
      ...common,
      content: { case: "change", value: wireDocumentChange(block) },
    };
  }
  if (block.kind === "image") {
    return {
      ...common,
      content: { case: "image", value: documentImage(block, assets) },
    };
  }
  if (block.kind === "section") {
    return {
      ...common,
      content: { case: "section", value: wireDocumentSection(block) },
    };
  }
  throw new OfficeKitCodecError(`The DOCX NativeAOT vertical slice cannot author document block kind ${block.kind}.`, [], { code: "unsupported_document_features" });
}

function wireDocumentProtection(value) {
  if (value == null) return undefined;
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new OfficeKitCodecError("Document protection must be a normalized settings object.", [], { code: "invalid_document_protection" });
  }
  const mode = value.edit === "none"
    ? DocumentProtectionMode.NONE
    : value.edit === "readOnly" ? DocumentProtectionMode.READ_ONLY
      : value.edit === "comments" ? DocumentProtectionMode.COMMENTS
        : value.edit === "trackedChanges" ? DocumentProtectionMode.TRACKED_CHANGES
          : value.edit === "forms" ? DocumentProtectionMode.FORMS : undefined;
  if (mode === undefined || typeof value.enforcement !== "boolean" || typeof value.formatting !== "boolean") {
    throw new OfficeKitCodecError("Document protection requires a canonical mode plus boolean enforcement and formatting flags.", [], { code: "invalid_document_protection" });
  }
  return { mode, enforcement: value.enforcement, formatting: value.formatting };
}

function publicDocumentProtection(value) {
  if (!value) return null;
  const edit = value.mode === DocumentProtectionMode.NONE
    ? "none"
    : value.mode === DocumentProtectionMode.READ_ONLY ? "readOnly"
      : value.mode === DocumentProtectionMode.COMMENTS ? "comments"
        : value.mode === DocumentProtectionMode.TRACKED_CHANGES ? "trackedChanges"
          : value.mode === DocumentProtectionMode.FORMS ? "forms" : undefined;
  if (!edit) throw new OfficeKitCodecError("OfficeKit returned an unsupported document-protection mode.", [], { code: "invalid_document_protection" });
  return { edit, enforcement: Boolean(value.enforcement), formatting: Boolean(value.formatting) };
}

function documentEnvelope(document) {
  if (!(document instanceof DocumentModel)) throw new TypeError("exportDocxWithOfficeKit expects a DocumentModel instance.");
  const state = document[DOCUMENT_STATE];
  assertTrustedImportedState(state, "DOCX");
  if (state && state.blocks.length !== document.blocks.length) {
    throw new OfficeKitCodecError(`Source-preserving DOCX export requires the original ${state.blocks.length}-block topology; the document contains ${document.blocks.length} blocks.`, [], { code: "document_topology_changed" });
  }
  if (state && state.comments.length !== document.comments.length) {
    throw new OfficeKitCodecError(`Source-preserving DOCX export requires the original ${state.comments.length}-comment topology; the document contains ${document.comments.length} comments.`, [], { code: "document_comment_topology_changed" });
  }
  validateDocumentCommentThreads(document);
  if (state && state.bookmarks.length !== document.bookmarks.length) {
    throw new OfficeKitCodecError(`Source-preserving DOCX export requires the original ${state.bookmarks.length}-bookmark topology; the document contains ${document.bookmarks.length} bookmarks.`, [], { code: "document_bookmark_topology_changed" });
  }
  if (state && state.notes.length !== document.notes.length) {
    throw new OfficeKitCodecError(`Source-preserving DOCX export requires the original ${state.notes.length}-note topology; the document contains ${document.notes.length} notes.`, [], { code: "document_note_topology_changed" });
  }
  if (state && Boolean(state.bibliography) !== Boolean(document.bibliographySources.length || Object.values(document.bibliography || {}).some(Boolean))) {
    throw new OfficeKitCodecError("Source-preserving DOCX export cannot add or remove the modeled bibliography catalog.", [], { code: "document_bibliography_topology_changed" });
  }
  if (state && ((state.headerSlots || state.headers).length !== document.headers.length || (state.footerSlots || state.footers).length !== document.footers.length)) {
    throw new OfficeKitCodecError("Source-preserving DOCX export requires the original header/footer topology.", [], { code: "document_header_footer_topology_changed" });
  }
  for (const slot of state?.readOnlyBlockSlots || []) {
    if (document.blocks[slot.index] !== slot.block || documentBlockSnapshot(slot.block) !== slot.publicSnapshot) {
      throw new OfficeKitCodecError(`Imported document block ${slot.wire.id} is source-bound and read-only in OfficeKit 0.2.`, [], { code: "unsupported_document_edit" });
    }
  }
  const assets = new Map((state?.assets || []).map((asset) => [asset.id, asset]));
  const directNumbering = state ? undefined : directDocumentNumberingPlan(document, assets);
  const contentControlNativeIds = planDocumentContentControls(document);
  const defaultRunSource = Object.fromEntries([...DOCUMENT_RUN_STYLE_KEYS].filter((key) => key !== "runStyleId" && Object.hasOwn(document.defaultRunStyle || {}, key)).map((key) => [key, document.defaultRunStyle[key]]));
  const blocks = document.blocks.map((block, index) => documentBlock(block, state?.blocks[index], directNumbering?.get(block), assets, contentControlNativeIds));
  return {
    protocolVersion: OFFICE_KIT_PROTOCOL_VERSION,
    family: ArtifactFamily.DOCUMENT,
    source: state?.source,
    opaqueOpc: state?.opaqueOpc,
    assets: [...assets.values()],
    diagnostics: state?.diagnostics || [],
    payload: {
      case: "document",
      value: {
        id: document.id,
        name: document.name,
        blocks,
        comments: document.comments.map((comment, index) => documentComment(comment, state?.comments[index])),
        bookmarks: document.bookmarks.map((bookmark, index) => documentBookmark(bookmark, state?.bookmarks[index], document)),
        notes: document.notes.map((note, index) => documentNote(note, state?.notes[index], document)),
        styles: document.styles.values().map(wireDocumentStyle),
        defaultRunStyle: documentRunFormatting(defaultRunSource, "Document default run style"),
        headers: wireDocumentHeaderFooters(document.headers, state?.headerSlots, "header"),
        footers: wireDocumentHeaderFooters(document.footers, state?.footerSlots, "footer"),
        watermarks: wireDocumentWatermarks(document, state),
        evenAndOddHeaders: Boolean(document.settings?.evenAndOddHeaders),
        mirrorMargins: Boolean(document.settings?.mirrorMargins),
        gutterAtTop: Boolean(document.settings?.gutterAtTop),
        updateFields: Boolean(document.settings?.updateFields),
        trackRevisions: Boolean(document.settings?.trackRevisions),
        documentProtection: wireDocumentProtection(document.settings?.documentProtection),
        sectionSettings: (document.sectionSettings || []).map((settings) => ({
          sectionIndex: uint32(settings.sectionIndex, "Document section settings index"),
          differentFirstPage: settings.differentFirstPage == null ? undefined : Boolean(settings.differentFirstPage),
        })),
        bibliography: wireDocumentBibliography(document, state?.bibliography),
      },
    },
  };
}

export async function exportDocxWithOfficeKit(document, options = {}) {
  assertCodecOptions(options, new Set(["limits"]), "exportDocxWithOfficeKit");
  const response = await invokeOfficeKitLazy(() => ({
    protocolVersion: OFFICE_KIT_PROTOCOL_VERSION,
    operation: CodecOperation.EXPORT_DOCX,
    family: ArtifactFamily.DOCUMENT,
    artifact: documentEnvelope(document),
    limits: codecLimits(options.limits),
  }));
  return new FileBlob(response.file, {
    type: DOCX_MIME,
    metadata: { artifactKind: "document", codec: "office-kit", diagnostics: response.diagnostics },
  });
}

export async function finalizeDocxRevisionsWithOfficeKit(input, options = {}) {
  assertCodecOptions(options, new Set(["mode", "keepTracking", "expectedSourceSha256", "limits"]), "finalizeDocxRevisionsWithOfficeKit");
  const mode = String(options.mode || "").trim().toLowerCase();
  const wireMode = mode === "accept"
    ? DocumentRevisionFinalizationMode.ACCEPT
    : mode === "reject"
      ? DocumentRevisionFinalizationMode.REJECT
      : undefined;
  if (wireMode == null) throw new TypeError("finalizeDocxRevisionsWithOfficeKit mode must be accept or reject.");
  if (options.keepTracking != null && typeof options.keepTracking !== "boolean") {
    throw new TypeError("finalizeDocxRevisionsWithOfficeKit keepTracking must be a boolean.");
  }
  const expectedSourceSha256 = String(options.expectedSourceSha256 || "").trim().toLowerCase();
  if (!/^[0-9a-f]{64}$/.test(expectedSourceSha256)) {
    throw new TypeError("finalizeDocxRevisionsWithOfficeKit expectedSourceSha256 must be a 64-character SHA-256 hex digest.");
  }
  const limits = codecLimits(options.limits);
  const file = await boundedInputBytes(input, limits, "DOCX");
  const actualSourceSha256 = createHash("sha256").update(file).digest("hex");
  if (actualSourceSha256 !== expectedSourceSha256) {
    throw new OfficeKitCodecError("DOCX revision finalization source bytes do not match expectedSourceSha256.", [], { code: "document_source_hash_mismatch" });
  }
  const response = await invokeOfficeKit({
    protocolVersion: OFFICE_KIT_PROTOCOL_VERSION,
    operation: CodecOperation.FINALIZE_DOCX_REVISIONS,
    family: ArtifactFamily.DOCUMENT,
    file,
    limits,
    revisionFinalization: {
      mode: wireMode,
      keepTracking: options.keepTracking === true,
      expectedSourceSha256,
    },
  });
  const result = response.revisionFinalization;
  const outputSha256 = createHash("sha256").update(response.file).digest("hex");
  const changedParts = result ? [...result.changedParts] : [];
  const allowedChangedParts = new Set(["word/document.xml", "word/settings.xml"]);
  if (!result ||
      result.mode !== wireMode ||
      result.sourceSha256 !== expectedSourceSha256 ||
      result.outputSha256 !== outputSha256 ||
      result.insertionCount + result.deletionCount === 0 ||
      result.trackingAfter !== (options.keepTracking === true && result.trackingBefore) ||
      changedParts.length !== new Set(changedParts).size ||
      !changedParts.includes("word/document.xml") ||
      changedParts.some((part) => !allowedChangedParts.has(part))) {
    throw new OfficeKitCodecError("OfficeKit returned an invalid DOCX revision-finalization audit result.", [], { code: "invalid_office_kit_response" });
  }
  return new FileBlob(response.file, {
    type: DOCX_MIME,
    metadata: {
      artifactKind: "document",
      codec: "office-kit",
      operation: "finalize-revisions",
      diagnostics: response.diagnostics,
      revisionFinalization: {
        mode,
        sourceSha256: result.sourceSha256,
        outputSha256,
        insertionCount: result.insertionCount,
        deletionCount: result.deletionCount,
        trackingBefore: result.trackingBefore,
        trackingAfter: result.trackingAfter,
        changedParts,
      },
    },
  });
}

export async function addDocxTrackedReplacementWithOfficeKit(input, options = {}) {
  assertCodecOptions(options, new Set([
    "target", "targetBlockIndex", "expectedText", "search", "replacement", "author", "date", "expectedSourceSha256", "limits",
  ]), "addDocxTrackedReplacementWithOfficeKit");
  const { target, wireTarget } = documentTrackedReplacementTarget(options);
  const expectedText = typeof options.expectedText === "string" ? options.expectedText : "";
  const search = typeof options.search === "string" ? options.search : "";
  const replacement = typeof options.replacement === "string" ? options.replacement : "";
  for (const [label, value] of [["expectedText", expectedText], ["search", search], ["replacement", replacement]]) {
    if (!value || value.length > 1_000_000 || !isXmlSafeText(value)) {
      throw new TypeError(`addDocxTrackedReplacementWithOfficeKit ${label} must contain 1 through 1,000,000 XML-safe characters.`);
    }
  }
  const author = typeof options.author === "string" ? options.author : "";
  if (!author.trim() || author.length > 255 || /[\u0000-\u001f\u007f]/.test(author)) {
    throw new TypeError("addDocxTrackedReplacementWithOfficeKit author must contain 1 through 255 characters without controls.");
  }
  const date = options.date == null || options.date === "" ? undefined : String(options.date);
  if (date !== undefined && (date.length > 64 || Number.isNaN(Date.parse(date)))) {
    throw new TypeError("addDocxTrackedReplacementWithOfficeKit date must be an ISO 8601 timestamp of at most 64 characters.");
  }
  const expectedSourceSha256 = String(options.expectedSourceSha256 || "").trim().toLowerCase();
  if (!/^[0-9a-f]{64}$/.test(expectedSourceSha256)) {
    throw new TypeError("addDocxTrackedReplacementWithOfficeKit expectedSourceSha256 must be a 64-character SHA-256 hex digest.");
  }
  const limits = codecLimits(options.limits);
  const file = await boundedInputBytes(input, limits, "DOCX");
  const actualSourceSha256 = createHash("sha256").update(file).digest("hex");
  if (actualSourceSha256 !== expectedSourceSha256) {
    throw new OfficeKitCodecError("DOCX tracked replacement source bytes do not match expectedSourceSha256.", [], { code: "document_source_hash_mismatch" });
  }
  const response = await invokeOfficeKit({
    protocolVersion: OFFICE_KIT_PROTOCOL_VERSION,
    operation: CodecOperation.ADD_DOCX_TRACKED_REPLACEMENT,
    family: ArtifactFamily.DOCUMENT,
    file,
    limits,
    trackedReplacement: {
      expectedSourceSha256,
      targetBlockIndex: target.blockIndex,
      expectedParagraphText: expectedText,
      search,
      replacement,
      author,
      date,
      target: wireTarget,
    },
  });
  const result = response.trackedReplacement;
  const outputSha256 = createHash("sha256").update(response.file).digest("hex");
  const deletedTextSha256 = createHash("sha256").update(search).digest("hex");
  const insertedTextSha256 = createHash("sha256").update(replacement).digest("hex");
  const changedParts = result ? [...result.changedParts] : [];
  const returnedTarget = publicDocumentTrackedReplacementTarget(result?.target);
  if (!result ||
      result.sourceSha256 !== expectedSourceSha256 ||
      result.outputSha256 !== outputSha256 ||
      outputSha256 === expectedSourceSha256 ||
      result.targetBlockIndex !== target.blockIndex ||
      !sameDocumentTrackedReplacementTarget(returnedTarget, target) ||
      !Number.isInteger(result.targetBodyIndex) ||
      !/^[0-9a-f]{64}$/.test(result.sourceElementSha256) ||
      !/^[0-9a-f]{64}$/.test(result.outputElementSha256) ||
      result.sourceElementSha256 === result.outputElementSha256 ||
      result.deletedTextSha256 !== deletedTextSha256 ||
      result.insertedTextSha256 !== insertedTextSha256 ||
      result.deletedTextChars !== search.length ||
      result.insertedTextChars !== replacement.length ||
      !Number.isInteger(result.matchedSourceRunCount) ||
      result.matchedSourceRunCount < 1 ||
      result.matchedSourceRunCount > search.length ||
      !/^\d+$/.test(result.deletionNativeRevisionId) ||
      !/^\d+$/.test(result.insertionNativeRevisionId) ||
      result.deletionNativeRevisionId === result.insertionNativeRevisionId ||
      changedParts.length !== 1 || changedParts[0] !== "word/document.xml") {
    throw new OfficeKitCodecError("OfficeKit returned an invalid DOCX tracked-replacement audit result.", [], { code: "invalid_office_kit_response" });
  }
  return new FileBlob(response.file, {
    type: DOCX_MIME,
    metadata: {
      artifactKind: "document",
      codec: "office-kit",
      operation: "add-tracked-replacement",
      diagnostics: response.diagnostics,
      trackedReplacement: {
        sourceSha256: result.sourceSha256,
        outputSha256,
        target: returnedTarget,
        targetBlockIndex: result.targetBlockIndex,
        targetBodyIndex: result.targetBodyIndex,
        sourceElementSha256: result.sourceElementSha256,
        outputElementSha256: result.outputElementSha256,
        deletedTextSha256,
        insertedTextSha256,
        deletedTextChars: result.deletedTextChars,
        insertedTextChars: result.insertedTextChars,
        matchedSourceRunCount: result.matchedSourceRunCount,
        deletionNativeRevisionId: result.deletionNativeRevisionId,
        insertionNativeRevisionId: result.insertionNativeRevisionId,
        changedParts,
      },
    },
  });
}

function documentTrackedReplacementTarget(options) {
  if (options.target === undefined) {
    if (!Number.isInteger(options.targetBlockIndex) || options.targetBlockIndex < 0 || options.targetBlockIndex > 0xffff_ffff) {
      throw new TypeError("addDocxTrackedReplacementWithOfficeKit targetBlockIndex must be an unsigned 32-bit integer from document.inspect().");
    }
    const target = { kind: "paragraph", blockIndex: options.targetBlockIndex };
    return {
      target,
      wireTarget: { blockIndex: target.blockIndex, location: { case: "bodyParagraph", value: {} } },
    };
  }
  if (options.targetBlockIndex !== undefined) {
    throw new TypeError("addDocxTrackedReplacementWithOfficeKit accepts either target or targetBlockIndex, not both.");
  }
  if (!options.target || typeof options.target !== "object" || Array.isArray(options.target)) {
    throw new TypeError("addDocxTrackedReplacementWithOfficeKit target must be a paragraph or tableCell selector object.");
  }
  const kind = String(options.target.kind || "");
  const blockIndex = options.target.blockIndex;
  if (!Number.isInteger(blockIndex) || blockIndex < 0 || blockIndex > 0xffff_ffff) {
    throw new TypeError("addDocxTrackedReplacementWithOfficeKit target.blockIndex must be an unsigned 32-bit integer from document.inspect().");
  }
  if (kind === "paragraph") {
    if (options.target.row !== undefined || options.target.column !== undefined) {
      throw new TypeError("addDocxTrackedReplacementWithOfficeKit paragraph target cannot include row or column.");
    }
    const target = { kind, blockIndex };
    return { target, wireTarget: { blockIndex, location: { case: "bodyParagraph", value: {} } } };
  }
  if (kind !== "tableCell") {
    throw new TypeError("addDocxTrackedReplacementWithOfficeKit target.kind must be paragraph or tableCell.");
  }
  const row = options.target.row;
  const column = options.target.column;
  if (!Number.isInteger(row) || row < 0 || row > 0xffff_ffff ||
      !Number.isInteger(column) || column < 0 || column > 0xffff_ffff) {
    throw new TypeError("addDocxTrackedReplacementWithOfficeKit tableCell target row and column must be unsigned 32-bit physical indexes from document.inspect().");
  }
  const target = { kind, blockIndex, row, column };
  return {
    target,
    wireTarget: { blockIndex, location: { case: "tableCell", value: { row, column } } },
  };
}

function publicDocumentTrackedReplacementTarget(target) {
  if (!target || !Number.isInteger(target.blockIndex)) return undefined;
  if (target.location.case === "bodyParagraph") return { kind: "paragraph", blockIndex: target.blockIndex };
  if (target.location.case === "tableCell" && target.location.value &&
      Number.isInteger(target.location.value.row) && Number.isInteger(target.location.value.column)) {
    return {
      kind: "tableCell",
      blockIndex: target.blockIndex,
      row: target.location.value.row,
      column: target.location.value.column,
    };
  }
  return undefined;
}

function sameDocumentTrackedReplacementTarget(left, right) {
  return left?.kind === right.kind && left.blockIndex === right.blockIndex &&
    (right.kind !== "tableCell" || left.row === right.row && left.column === right.column);
}

function publicDocumentContentControl(control) {
  if (!control) return undefined;
  return {
    id: control.id,
    tag: control.tag,
    alias: control.alias,
    nativeId: control.nativeId,
    controlType: control.controlType === DocumentContentControlType.CHECKBOX
      ? "checkbox"
      : control.controlType === DocumentContentControlType.DROP_DOWN
        ? "dropdown"
        : control.controlType === DocumentContentControlType.COMBO_BOX ? "comboBox"
          : control.controlType === DocumentContentControlType.DATE ? "date" : "text",
    ...(control.controlType === DocumentContentControlType.CHECKBOX ? { checked: control.checked === true } : {}),
    ...(control.controlType === DocumentContentControlType.DROP_DOWN ? {
      choices: control.choices.map((choice) => ({ displayText: choice.displayText, value: choice.value })),
      selectedValue: control.selectedValue,
    } : {}),
    ...(control.controlType === DocumentContentControlType.COMBO_BOX ? {
      choices: control.choices.map((choice) => ({ displayText: choice.displayText, value: choice.value })),
      value: control.value,
    } : {}),
    ...(control.controlType === DocumentContentControlType.DATE ? { dateValue: control.dateValue } : {}),
  };
}

function documentFromEnvelope(envelope) {
  if (envelope.family !== ArtifactFamily.DOCUMENT || envelope.payload.case !== "document") {
    throw new OfficeKitCodecError("OfficeKit response does not contain a document artifact.", [], { code: "invalid_document_artifact" });
  }
  const source = envelope.payload.value;
  const assets = new Map((envelope.assets || []).map((asset) => [asset.id, asset]));
  const styles = Object.fromEntries((source.styles || []).map((style) => [style.id, publicDocumentStyle(style)]));
  if (!(source.styles || []).length) for (const block of source.blocks) {
    if (block.styleId) styles[block.styleId] = { id: block.styleId, name: block.styleId, type: block.content.case === "table" ? "table" : "paragraph" };
    if (block.content.case === "paragraph") for (const run of block.content.value.runs) if (run.styleId) styles[run.styleId] = { id: run.styleId, name: run.styleId, type: "character" };
  }
  const blocks = source.blocks.map((block) => {
    switch (block.content.case) {
      case "paragraph":
        if (block.content.value.numbering) {
          const paragraph = block.content.value;
          const numbering = paragraph.numbering;
          const numberFormat = numbering.numberFormat || "decimal";
          return {
            kind: "listItem",
            id: block.id,
            name: block.name,
            styleId: block.styleId || "Normal",
            text: paragraph.text,
            listType: numberFormat === "bullet" ? "bullet" : "number",
            numberFormat,
            level: numbering.level,
            start: numbering.start || 1,
            levelText: numbering.levelText || (numberFormat === "bullet" ? "•" : `%${numbering.level + 1}.`),
            numberingId: numbering.numberingId,
            abstractNumberingId: numbering.abstractNumberingId,
            numberingStyleId: numbering.numberingStyleId || undefined,
            pictureBullet: publicDocumentPictureBullet(numbering.pictureBullet, assets, `Document list item ${block.id} picture bullet`),
          };
        }
        return {
          kind: "paragraph",
          id: block.id,
          name: block.name,
          styleId: block.styleId || "Normal",
          textEditable: block.source?.editable !== false,
          textPatchable: block.source?.textPatchable === true,
          textPatches: [],
          text: block.content.value.text,
          paragraphFormat: publicDocumentParagraphFormatting(block.content.value.formatting),
          blockContentControl: publicDocumentContentControl(block.content.value.blockContentControl),
          runs: block.content.value.runs.length ? block.content.value.runs.map((run) => ({
            text: run.text,
            style: {
              ...(run.styleId ? { runStyleId: run.styleId } : {}),
              ...publicDocumentRunFormatting(run.formatting),
              ...(!run.formatting && run.bold ? { bold: true } : {}),
              ...(!run.formatting && run.italic ? { italic: true } : {}),
              ...(!run.formatting && run.underline ? { underline: true } : {}),
            },
            ...(run.textContentControl ? { contentControl: publicDocumentContentControl(run.textContentControl) } : {}),
            ...(run.inlineField ? { inlineField: {
              instruction: run.inlineField.instruction,
              ...(run.inlineField.bookmarkName ? { bookmarkName: run.inlineField.bookmarkName } : {}),
              ...(run.inlineField.bookmarkNativeId !== "" ? { bookmarkNativeId: Number(run.inlineField.bookmarkNativeId) } : {}),
            } } : {}),
          })) : undefined,
        };
      case "table":
        {
          const formatting = documentTableFormattingConfig(block.content.value);
          const accessibility = {};
          if (block.content.value.accessibilityTitle !== undefined) accessibility.title = block.content.value.accessibilityTitle;
          if (block.content.value.accessibilityDescription !== undefined) accessibility.description = block.content.value.accessibilityDescription;
        return {
          kind: "table",
          id: block.id,
          name: block.name,
          styleId: block.styleId || "TableGrid",
          sourceBound: Boolean(block.source),
          values: block.content.value.rows.map((row) => [...row.cells]),
          gridColumns: block.content.value.gridColumns,
          headerRowCount: Number(block.content.value.headerRowCount || 0),
          keepTogetherRows: (block.content.value.keepTogetherRows || []).map((value) => Number(value)),
          minimumRowHeightsDxa: (block.content.value.minimumRowHeightsDxa || []).length === block.content.value.rows.length
            ? block.content.value.minimumRowHeightsDxa.map((value) => Number(value))
            : Array.from({ length: block.content.value.rows.length }, () => 0),
          ...(Object.keys(accessibility).length ? { accessibility } : {}),
          cells: documentTableCells(block.content.value),
          textPatches: [],
          ...formatting,
        };
        }
      case "hyperlink": {
        const hyperlink = block.content.value;
        return {
          kind: "hyperlink",
          id: block.id,
          name: block.name,
          styleId: block.styleId || "Normal",
          text: hyperlink.text,
          url: hyperlink.target.case === "externalUri" ? hyperlink.target.value : undefined,
          anchor: hyperlink.target.case === "internalAnchor" ? hyperlink.target.value : undefined,
          relationshipId: hyperlink.relationshipId || undefined,
          tooltip: hyperlink.tooltip,
          history: hyperlink.history,
        };
      }
      case "field":
        return {
          kind: "field",
          id: block.id,
          name: block.name,
          styleId: block.styleId || "Normal",
          instruction: block.content.value.instruction,
          display: block.content.value.display,
          complex: Boolean(block.content.value.complex),
        };
      case "citation":
        return {
          kind: "citation",
          id: block.id,
          name: block.name,
          styleId: block.styleId || "Normal",
          text: block.content.value.display,
          metadata: { tag: block.content.value.tag },
          _restore: true,
        };
      case "change": {
        const change = block.content.value;
        return {
          kind: "change",
          id: block.id,
          name: block.name,
          styleId: block.styleId || "Normal",
          changeType: publicDocumentChangeType(change.type),
          text: change.text,
          author: change.author,
          date: change.date,
          _restore: true,
        };
      }
      case "image": {
        const image = block.content.value;
        const asset = assets.get(image.assetId);
        if (!asset || !new Set(["image/png", "image/jpeg"]).has(asset.contentType)) throw new OfficeKitCodecError(`Document image ${block.id} references a missing or unsupported asset.`, [], { code: "invalid_document_asset" });
        return {
          kind: "image",
          id: block.id,
          name: block.name,
          styleId: block.styleId || "Normal",
          dataUrl: `data:${asset.contentType};base64,${Buffer.from(asset.data).toString("base64")}`,
          alt: image.altText,
          widthPx: Number(image.widthEmu) / 9_525,
          heightPx: Number(image.heightEmu) / 9_525,
          placement: publicDocumentFloatingImagePlacement(image.floating),
        };
      }
      case "section": {
        const section = block.content.value;
        return {
          kind: "section",
          id: block.id,
          name: block.name,
          editable: block.source?.editable !== false,
          breakType: publicDocumentSectionBreak(section.breakType),
          orientation: section.landscape ? "landscape" : "portrait",
          pageSize: { widthTwips: section.pageWidthTwips, heightTwips: section.pageHeightTwips },
          margins: { top: section.marginTopTwips, right: section.marginRightTwips, bottom: section.marginBottomTwips, left: section.marginLeftTwips, gutter: section.marginGutterTwips },
          columns: section.columns ? (section.columns.definitions?.length ? {
            definitions: section.columns.definitions.map((definition) => ({ width: definition.widthTwips, spacing: definition.spacingAfterTwips })),
            separator: section.columns.separator,
          } : {
            count: section.columns.count,
            spacing: section.columns.spacingTwips,
            separator: section.columns.separator,
          }) : undefined,
          pageNumbering: publicDocumentSectionPageNumbering(section.pageNumbering),
          lineNumbering: publicDocumentSectionLineNumbering(section.lineNumbering),
        };
      }
      case "opaque":
        return {
          kind: "paragraph",
          id: block.id,
          name: block.name || `Preserved ${block.content.value.elementName}`,
          styleId: "Normal",
          textEditable: false,
          textPatchable: false,
          text: block.content.value.text,
        };
      default:
        throw new OfficeKitCodecError(`Document block ${block.id} has no supported wire content.`, [], { code: "invalid_document_artifact" });
    }
  });
  const comments = source.comments.map((comment) => ({
    id: comment.id,
    targetId: comment.targetBlockId,
    parentId: comment.parentCommentId || undefined,
    author: comment.author,
    initials: comment.initials,
    date: comment.createdAt,
    text: comment.text,
    resolved: comment.resolved,
    paraId: comment.paragraphId || undefined,
    durableId: comment.durableId || undefined,
    dateUtc: comment.dateUtc,
    person: comment.person ? { providerId: comment.person.providerId, userId: comment.person.userId } : undefined,
    intelligentPlaceholder: comment.intelligentPlaceholder ?? false,
  }));
  const bookmarks = (source.bookmarks || []).map((bookmark) => ({
    id: bookmark.id,
    name: bookmark.name,
    targetId: bookmark.targetBlockId,
    endTargetId: bookmark.endTargetBlockId,
    nativeId: bookmark.nativeId === "" ? undefined : Number(bookmark.nativeId),
  }));
  const notes = (source.notes || []).map((note) => ({
    id: note.id,
    kind: publicDocumentNoteKind(note.kind),
    targetId: note.targetBlockId,
    text: note.text,
    paragraphs: note.paragraphs?.length ? [...note.paragraphs] : [note.text],
    nativeId: note.nativeId === "" ? undefined : Number(note.nativeId),
  }));
  const document = DocumentModel.create({
    name: source.name || "Imported document",
    styles,
    defaultRunStyle: publicDocumentRunFormatting(source.defaultRunStyle),
    blocks,
    comments,
    bookmarks,
    notes,
    bibliography: source.bibliography ? {
      selectedStyle: source.bibliography.selectedStyle,
      styleName: source.bibliography.styleName,
      uri: source.bibliography.uri,
    } : undefined,
    bibliographySources: (source.bibliography?.sources || []).map(publicDocumentBibliographySource),
    headers: (source.headers || []).map(publicHeaderFooter),
    footers: (source.footers || []).map(publicHeaderFooter),
    watermarks: (source.watermarks || []).map((watermark) => ({
      id: watermark.id,
      text: watermark.text,
      referenceType: publicHeaderFooterReference(watermark.reference),
      sectionIndex: watermark.sectionIndex,
      editable: watermark.source?.editable === true,
      sourceBound: Boolean(watermark.source),
    })),
    settings: {
      evenAndOddHeaders: Boolean(source.evenAndOddHeaders),
      mirrorMargins: Boolean(source.mirrorMargins),
      gutterAtTop: Boolean(source.gutterAtTop),
      updateFields: Boolean(source.updateFields),
      trackRevisions: Boolean(source.trackRevisions),
      documentProtection: publicDocumentProtection(source.documentProtection),
    },
    sectionSettings: (source.sectionSettings || []).map((settings) => ({ sectionIndex: settings.sectionIndex, differentFirstPage: settings.differentFirstPage })),
  });
  document.id = source.id || document.id;
  const commentSlots = source.comments.map((wire, index) => ({
    wire,
    publicSnapshot: documentCommentSnapshot(document.comments[index]),
  }));
  const bookmarkSlots = (source.bookmarks || []).map((wire, index) => ({
    wire,
    publicSnapshot: documentBookmarkSnapshot(document.bookmarks[index]),
  }));
  const noteSlots = (source.notes || []).map((wire, index) => ({
    wire,
    publicSnapshot: documentNoteSnapshot(document.notes[index]),
  }));
  const watermarkSlots = (source.watermarks || []).map((wire, index) => ({
    wire,
    publicSnapshot: documentWatermarkSnapshot(document.watermarks[index]),
  }));
  const headerSlots = (source.headers || []).map((wire, index) => ({
    wire,
    publicSnapshot: documentHeaderFooterSnapshot(document.headers[index]),
  }));
  const footerSlots = (source.footers || []).map((wire, index) => ({
    wire,
    publicSnapshot: documentHeaderFooterSnapshot(document.footers[index]),
  }));
  const readOnlyBlockSlots = source.blocks.flatMap((wire, index) => {
    if (wire.content.case !== "opaque" && (wire.source?.editable !== false || wire.source?.textPatchable === true)) return [];
    const block = document.blocks[index];
    return [{ wire, index, block, publicSnapshot: documentBlockSnapshot(block) }];
  });
  Object.defineProperty(document, DOCUMENT_STATE, {
    configurable: true,
    value: { source: envelope.source, opaqueOpc: envelope.opaqueOpc, diagnostics: envelope.diagnostics, assets: envelope.assets || [], blocks: source.blocks, readOnlyBlockSlots, comments: commentSlots, bookmarks: bookmarkSlots, notes: noteSlots, watermarkSlots, headerSlots, footerSlots, bibliography: source.bibliography, headers: source.headers || [], footers: source.footers || [] },
    writable: true,
  });
  return document;
}

export async function importDocxWithOfficeKit(input, options = {}) {
  assertCodecOptions(options, new Set(["limits"]), "importDocxWithOfficeKit");
  const limits = codecLimits(options.limits);
  return invokeOfficeKit({
    protocolVersion: OFFICE_KIT_PROTOCOL_VERSION,
    operation: CodecOperation.IMPORT_DOCX,
    family: ArtifactFamily.DOCUMENT,
    file: await boundedInputBytes(input, limits, "DOCX"),
    limits,
  }, { consumeResponse: (response) => documentFromEnvelope(response.artifact) });
}
