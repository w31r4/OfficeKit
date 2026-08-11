import crypto from "node:crypto";

import mupdf from "mupdf";

import { inspectCanonicalXmpMetadata, patchCanonicalXmpMetadata } from "./xmp-metadata.mjs";

const METADATA_KEYS = Object.freeze({
  author: mupdf.Document.META_INFO_AUTHOR,
  title: mupdf.Document.META_INFO_TITLE,
  subject: mupdf.Document.META_INFO_SUBJECT,
  keywords: mupdf.Document.META_INFO_KEYWORDS,
  creator: mupdf.Document.META_INFO_CREATOR,
  producer: mupdf.Document.META_INFO_PRODUCER,
  creationDate: mupdf.Document.META_INFO_CREATIONDATE,
  modificationDate: mupdf.Document.META_INFO_MODIFICATIONDATE,
});

const PDF_INFO_KEYS = Object.freeze({
  author: "Author",
  title: "Title",
  subject: "Subject",
  keywords: "Keywords",
  creator: "Creator",
  producer: "Producer",
  creationDate: "CreationDate",
  modificationDate: "ModDate",
});

const DOCUMENT_METADATA_ID = "mupdf-document-info";
const SNAPSHOT_FIELDS = Object.freeze([
  "values",
  "infoPresent",
  "infoObject",
  "infoEntries",
  "xmpPresent",
  "xmpObject",
  "xmpByteLength",
  "xmpSha256",
  "xmpProfile",
  "xmpValues",
  "xmpMutableFields",
]);
const OPERATION_FIELDS = new Set(["type", "sourceSha256", "metadataId", "expected", "patch"]);

function sha256(value) {
  return crypto.createHash("sha256").update(value).digest("hex");
}

function destroyPdfObject(object) {
  if (object && object !== mupdf.PDFObject.Null) object.destroy();
}

function pdfObjectXrefOrNull(object) {
  if (!object || object === mupdf.PDFObject.Null || !object.isIndirect()) return null;
  const value = object.asIndirect();
  return Number.isSafeInteger(value) && value > 0 ? value : null;
}

function copyMuPdfBuffer(value) {
  const bytes = value?.asUint8Array ? value.asUint8Array() : value;
  return Buffer.from(bytes).subarray(0);
}

function xmpStreamProfile(reference) {
  const empty = {
    byteLength: null,
    sha256: null,
    profile: "none",
    values: {},
    mutableFields: [],
    issues: [],
    bytes: null,
  };
  if (!reference || reference === mupdf.PDFObject.Null || reference.isNull()) return empty;
  const issues = [];
  let type;
  let subtype;
  let stream;
  try {
    if (!reference.isIndirect()) issues.push("catalog-xmp-metadata-is-not-indirect");
    if (!reference.isStream()) issues.push("catalog-xmp-metadata-is-not-a-stream");
    type = reference.get("Type");
    subtype = reference.get("Subtype");
    if (!type.isName() || type.asName() !== "Metadata") issues.push("catalog-xmp-type-is-not-metadata");
    if (!subtype.isName() || subtype.asName() !== "XML") issues.push("catalog-xmp-subtype-is-not-xml");
    if (issues.length) return { ...empty, profile: "unsupported", issues };
    stream = reference.readStream();
    const bytes = copyMuPdfBuffer(stream);
    const profile = inspectCanonicalXmpMetadata(bytes);
    return {
      ...profile,
      bytes,
      issues: profile.issues.map((issue) => `xmp-profile:${issue}`),
    };
  } catch (error) {
    return { ...empty, profile: "unsupported", issues: [`xmp-read:${String(error?.message || error)}`] };
  } finally {
    stream?.destroy();
    destroyPdfObject(subtype);
    destroyPdfObject(type);
  }
}

function currentXmpReference(document) {
  let trailer;
  let root;
  try {
    trailer = document.getTrailer();
    root = trailer.get("Root");
    return root.get("Metadata");
  } finally {
    destroyPdfObject(root);
    destroyPdfObject(trailer);
  }
}

export function metadataFor(document) {
  return Object.fromEntries(Object.entries(METADATA_KEYS)
    .map(([name, key]) => [name, document.getMetaData(key)])
    .filter(([, value]) => value !== undefined && value !== ""));
}

function normalizedValues(value, label) {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new Error(`${label} must be an object.`);
  for (const name of Object.keys(value)) {
    if (!Object.hasOwn(METADATA_KEYS, name)) throw new Error(`${label} contains unsupported metadata key: ${name}.`);
  }
  const normalized = {};
  for (const name of Object.keys(METADATA_KEYS)) {
    if (!Object.hasOwn(value, name)) continue;
    if (typeof value[name] !== "string" || !value[name].length) throw new Error(`${label}.${name} must be a non-empty string.`);
    normalized[name] = value[name];
  }
  return normalized;
}

function dictionaryEntryFingerprints(object) {
  const entries = [];
  if (!object || object === mupdf.PDFObject.Null || !object.isDictionary()) return entries;
  object.forEach((value, name) => {
    try {
      entries.push({ name: String(name), fingerprint: sha256(Buffer.from(value.toString(true, true), "utf8")) });
    } finally {
      destroyPdfObject(value);
    }
  });
  return entries.sort((left, right) => left.name < right.name ? -1 : left.name > right.name ? 1 : 0);
}

export function documentMetadataProfile(document) {
  const issues = [];
  let trailer;
  let infoReference;
  let info;
  let root;
  let xmpReference;
  try {
    trailer = document.getTrailer();
    infoReference = trailer.get("Info");
    const infoPresent = !infoReference.isNull();
    const infoObject = pdfObjectXrefOrNull(infoReference);
    let infoEntries = [];
    if (infoPresent) {
      if (!infoReference.isDictionary()) issues.push("document-info-is-not-a-dictionary");
      else {
        info = infoReference.resolve();
        infoEntries = dictionaryEntryFingerprints(info);
      }
    }

    root = trailer.get("Root");
    if (!root.isDictionary()) issues.push("catalog-root-is-not-a-dictionary");
    xmpReference = root.isDictionary() ? root.get("Metadata") : mupdf.PDFObject.Null;
    const xmpPresent = !xmpReference.isNull();
    const xmp = xmpStreamProfile(xmpReference);
    issues.push(...xmp.issues);
    if (xmpPresent && !xmp.mutableFields.length && !xmp.issues.length) issues.push("xmp-has-no-supported-standard-properties");

    const snapshot = {
      values: metadataFor(document),
      infoPresent,
      infoObject,
      infoEntries,
      xmpPresent,
      xmpObject: pdfObjectXrefOrNull(xmpReference),
      xmpByteLength: xmp.byteLength,
      xmpSha256: xmp.sha256,
      xmpProfile: xmp.profile,
      xmpValues: xmp.values,
      xmpMutableFields: xmp.mutableFields,
    };
    return {
      snapshot,
      record: {
        kind: "mupdfDocumentMetadata",
        id: DOCUMENT_METADATA_ID,
        snapshot,
        updateCapability: issues.length
          ? { supported: false, reasons: issues }
          : {
              supported: true,
              sourceBound: true,
              savePolicies: ["rewrite", "incremental"],
              xmpSynchronized: xmpPresent,
              mutableFields: xmpPresent ? xmp.mutableFields : Object.keys(METADATA_KEYS),
            },
      },
    };
  } finally {
    destroyPdfObject(xmpReference);
    destroyPdfObject(root);
    destroyPdfObject(info);
    destroyPdfObject(infoReference);
    destroyPdfObject(trailer);
  }
}

function expectedSnapshot(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new Error("set_metadata expected must be the complete snapshot from one mupdfDocumentMetadata record.");
  }
  for (const field of Object.keys(value)) {
    if (!SNAPSHOT_FIELDS.includes(field)) throw new Error(`set_metadata expected contains unsupported snapshot field: ${field}.`);
  }
  for (const field of SNAPSHOT_FIELDS) {
    if (!Object.hasOwn(value, field)) throw new Error(`set_metadata expected must include snapshot field: ${field}.`);
  }
  const values = normalizedValues(value.values, "set_metadata expected.values");
  for (const name of ["infoPresent", "xmpPresent"]) {
    if (typeof value[name] !== "boolean") throw new Error(`set_metadata expected.${name} must be a boolean.`);
  }
  for (const name of ["infoObject", "xmpObject"]) {
    if (value[name] !== null && (!Number.isSafeInteger(value[name]) || value[name] < 1)) {
      throw new Error(`set_metadata expected.${name} must be a positive safe integer or null.`);
    }
  }
  if (value.xmpByteLength !== null && (!Number.isSafeInteger(value.xmpByteLength) || value.xmpByteLength < 0)) {
    throw new Error("set_metadata expected.xmpByteLength must be a nonnegative safe integer or null.");
  }
  if (value.xmpSha256 !== null && !/^[a-f0-9]{64}$/u.test(String(value.xmpSha256))) {
    throw new Error("set_metadata expected.xmpSha256 must be a SHA-256 hex digest or null.");
  }
  if (!["none", "canonical-simple-v1", "unsupported"].includes(value.xmpProfile)) {
    throw new Error("set_metadata expected.xmpProfile is unsupported.");
  }
  const xmpValues = normalizedValues(value.xmpValues, "set_metadata expected.xmpValues");
  if (!Array.isArray(value.xmpMutableFields)) throw new Error("set_metadata expected.xmpMutableFields must be an array.");
  const xmpMutableFields = value.xmpMutableFields.map((field, index) => {
    if (typeof field !== "string" || !Object.hasOwn(METADATA_KEYS, field)) {
      throw new Error(`set_metadata expected.xmpMutableFields[${index}] is unsupported.`);
    }
    return field;
  });
  if (new Set(xmpMutableFields).size !== xmpMutableFields.length) {
    throw new Error("set_metadata expected.xmpMutableFields must contain unique fields.");
  }
  if (!Array.isArray(value.infoEntries)) throw new Error("set_metadata expected.infoEntries must be an array.");
  const names = new Set();
  const infoEntries = value.infoEntries.map((entry, index) => {
    if (!entry || typeof entry !== "object" || Array.isArray(entry) || Object.keys(entry).some((name) => name !== "name" && name !== "fingerprint")) {
      throw new Error(`set_metadata expected.infoEntries[${index}] must be an inspect-derived {name,fingerprint} entry.`);
    }
    if (typeof entry.name !== "string" || !entry.name || names.has(entry.name)) {
      throw new Error("set_metadata expected.infoEntries must contain unique non-empty names.");
    }
    if (!/^[a-f0-9]{64}$/u.test(String(entry.fingerprint || ""))) {
      throw new Error(`set_metadata expected.infoEntries[${index}].fingerprint must be a SHA-256 hex digest.`);
    }
    names.add(entry.name);
    return { name: entry.name, fingerprint: entry.fingerprint };
  }).sort((left, right) => left.name < right.name ? -1 : left.name > right.name ? 1 : 0);
  return {
    values,
    infoPresent: value.infoPresent,
    infoObject: value.infoObject,
    infoEntries,
    xmpPresent: value.xmpPresent,
    xmpObject: value.xmpObject,
    xmpByteLength: value.xmpByteLength,
    xmpSha256: value.xmpSha256,
    xmpProfile: value.xmpProfile,
    xmpValues,
    xmpMutableFields,
  };
}

function metadataPatch(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new Error("set_metadata patch must be an object with at least one mutable field.");
  }
  for (const name of Object.keys(value)) {
    if (!Object.hasOwn(METADATA_KEYS, name)) throw new Error(`set_metadata patch contains unsupported metadata key: ${name}.`);
  }
  const patch = {};
  for (const name of Object.keys(METADATA_KEYS)) {
    if (!Object.hasOwn(value, name)) continue;
    if (value[name] !== null && (typeof value[name] !== "string" || !value[name].length)) {
      throw new Error(`set_metadata patch.${name} must be a non-empty string or null to clear.`);
    }
    patch[name] = value[name];
  }
  if (!Object.keys(patch).length) throw new Error("set_metadata patch must include at least one metadata field.");
  return patch;
}

function equal(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function updateXmpStream(document, snapshot, patch) {
  let reference;
  let stream;
  try {
    reference = currentXmpReference(document);
    if (reference.isNull() || !reference.isIndirect() || reference.asIndirect() !== snapshot.xmpObject || !reference.isStream()) {
      throw new Error("set_metadata XMP stream identity changed after inspection; refusing to write.");
    }
    stream = reference.readStream();
    const bytes = copyMuPdfBuffer(stream);
    const profile = inspectCanonicalXmpMetadata(bytes);
    if (profile.sha256 !== snapshot.xmpSha256 || profile.byteLength !== snapshot.xmpByteLength || profile.profile !== snapshot.xmpProfile) {
      throw new Error("set_metadata XMP stream no longer matches the inspected snapshot; refusing to write.");
    }
    const updated = patchCanonicalXmpMetadata(bytes, profile, patch);
    reference.writeStream(updated);
    return updated;
  } finally {
    stream?.destroy();
    destroyPdfObject(reference);
  }
}

export function applySourceBoundMetadataUpdate(document, operation, context = {}) {
  for (const name of Object.keys(operation)) {
    if (!OPERATION_FIELDS.has(name)) throw new Error(`set_metadata contains unsupported field: ${name}.`);
  }
  if (!/^[a-f0-9]{64}$/u.test(String(operation.sourceSha256 || "")) || operation.sourceSha256 !== context.sourceSha256) {
    throw new Error("set_metadata sourceSha256 must exactly match PdfFile.inspectPdf(...).summary.sourceSha256 for the current input bytes.");
  }
  if (operation.metadataId !== DOCUMENT_METADATA_ID) {
    throw new Error(`set_metadata metadataId must be ${DOCUMENT_METADATA_ID} from PdfFile.inspectPdf.`);
  }
  const expected = expectedSnapshot(operation.expected);
  const before = documentMetadataProfile(document);
  if (!equal(before.snapshot, expected)) {
    throw new Error("set_metadata precondition did not match the current Info/XMP metadata snapshot; re-inspect the exact input bytes.");
  }
  if (!before.record.updateCapability.supported) throw new Error(`set_metadata is unsupported for this metadata graph: ${before.record.updateCapability.reasons.join(", ")}. Preserve the PDF unchanged or use an explicit provider that supports this XMP profile.`);
  const patch = metadataPatch(operation.patch);
  if (before.snapshot.xmpPresent) {
    const unsupportedFields = Object.keys(patch).filter((field) => !before.snapshot.xmpMutableFields.includes(field));
    if (unsupportedFields.length) {
      throw new Error(`set_metadata cannot synchronize XMP field${unsupportedFields.length === 1 ? "" : "s"} ${unsupportedFields.join(", ")} because the canonical packet does not contain ${unsupportedFields.length === 1 ? "that property" : "those properties"}.`);
    }
  }
  if (!Object.entries(patch).some(([name, value]) => {
    const requested = value === null ? undefined : value;
    return requested !== before.snapshot.values[name] || (before.snapshot.xmpPresent && requested !== before.snapshot.xmpValues[name]);
  })) {
    throw new Error("set_metadata patch would not change the inspected Info/XMP metadata values.");
  }

  for (const [name, value] of Object.entries(patch)) document.setMetaData(METADATA_KEYS[name], value ?? "");
  const expectedXmpBytes = before.snapshot.xmpPresent ? updateXmpStream(document, before.snapshot, patch) : null;
  const after = documentMetadataProfile(document);
  if (!after.record.updateCapability.supported) throw new Error("MuPDF produced an unsupported metadata graph; refusing to save.");

  for (const name of Object.keys(METADATA_KEYS)) {
    const expectedValue = Object.hasOwn(patch, name) ? patch[name] ?? undefined : before.snapshot.values[name];
    if (after.snapshot.values[name] !== expectedValue) {
      throw new Error(`MuPDF did not preserve the requested set_metadata ${name} value; refusing to save.`);
    }
  }
  const patchedInfoKeys = new Set(Object.keys(patch).map((name) => PDF_INFO_KEYS[name]));
  const beforeResidual = before.snapshot.infoEntries.filter((entry) => !patchedInfoKeys.has(entry.name));
  const afterResidual = after.snapshot.infoEntries.filter((entry) => !patchedInfoKeys.has(entry.name));
  if (!equal(beforeResidual, afterResidual)) {
    throw new Error("MuPDF changed a non-target Document Info entry while applying set_metadata; refusing to save.");
  }
  if (before.snapshot.xmpPresent) {
    for (const [name, value] of Object.entries(patch)) {
      const expectedValue = value ?? undefined;
      if (after.snapshot.xmpValues[name] !== expectedValue) {
        throw new Error(`MuPDF did not preserve the requested synchronized XMP ${name} value; refusing to save.`);
      }
    }
    if (after.snapshot.xmpSha256 !== sha256(expectedXmpBytes)) {
      throw new Error("MuPDF changed bytes outside the requested canonical XMP metadata values; refusing to save.");
    }
  }
  return {
    type: "set_metadata",
    metadataId: DOCUMENT_METADATA_ID,
    matched: before.snapshot,
    patch,
    updated: after.snapshot,
    xmpSynchronized: before.snapshot.xmpPresent,
  };
}
