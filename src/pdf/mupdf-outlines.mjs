import crypto from "node:crypto";

const OUTLINE_ID_PREFIX = "mupdf-outline";
const OUTLINE_MAX_DEPTH = 256;
const OUTLINE_MAX_TITLE_LENGTH = 4_096;
const OUTLINE_SNAPSHOT_FIELDS = Object.freeze([
  "path",
  "title",
  "uri",
  "open",
  "page",
  "childCount",
]);
const UPDATE_OPERATION_FIELDS = new Set([
  "type",
  "sourceSha256",
  "outlineId",
  "expected",
  "patch",
]);
const UPDATE_PATCH_FIELDS = new Set(["title", "open"]);

function digest(value) {
  return crypto.createHash("sha256").update(JSON.stringify(value)).digest("hex");
}

function sameValue(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function pathLabel(path) {
  return path.map((index) => index + 1).join(".");
}

function resolvedPageFor(document, uri, pageCount) {
  if (typeof uri !== "string") return null;
  try {
    const page = document.resolveLink(uri);
    return Number.isInteger(page) && page >= 0 && page < pageCount ? page + 1 : null;
  } catch {
    return null;
  }
}

function snapshotFor(document, item, path, pageCount) {
  const uri = typeof item.uri === "string" ? item.uri : null;
  return {
    path: [...path],
    title: typeof item.title === "string" ? item.title : null,
    uri,
    open: Boolean(item.open),
    page: resolvedPageFor(document, uri, pageCount),
    childCount: 0,
  };
}

function outlineId(snapshot) {
  return `${OUTLINE_ID_PREFIX}-${pathLabel(snapshot.path)}-${digest(snapshot)}`;
}

function updateCapability(snapshot) {
  const mutableFields = ["title"];
  const blockedFields = [];
  if (snapshot.childCount > 0) mutableFields.push("open");
  else blockedFields.push({ field: "open", reason: "leaf outline has no child expansion state" });
  return { supported: true, mutableFields, blockedFields };
}

export function documentOutlineProfile(document, limits = {}) {
  const maxOutlines = Number(limits.maxOutlines ?? 100_000);
  if (!Number.isSafeInteger(maxOutlines) || maxOutlines <= 0) {
    throw new Error("PDF limit maxOutlines must be a positive safe integer.");
  }
  const iterator = document.outlineIterator();
  const snapshots = [];
  try {
    if (!iterator.item()) return { count: 0, records: [] };
    const pageCount = document.countPages();
    const parentIndexes = [];
    let path = [0];
    while (true) {
      const item = iterator.item();
      if (!item) throw new Error("MuPDF outline iterator lost its current item during bounded inspection.");
      if (path.length > OUTLINE_MAX_DEPTH) {
        throw new Error(`PDF outline depth exceeds the ${OUTLINE_MAX_DEPTH}-level inspection budget.`);
      }
      if (snapshots.length >= maxOutlines) {
        throw new Error(`PDF exceeds maxOutlines (${snapshots.length + 1} > ${maxOutlines}).`);
      }
      if (parentIndexes.length) snapshots[parentIndexes.at(-1)].childCount += 1;
      const currentIndex = snapshots.push(snapshotFor(document, item, path, pageCount)) - 1;

      const down = iterator.down();
      if (down === 0) {
        parentIndexes.push(currentIndex);
        path = [...path, 0];
        continue;
      }
      if (down === 1 && iterator.up() !== 0) {
        throw new Error("MuPDF outline iterator could not restore a leaf after probing its children.");
      }

      let advanced = false;
      while (true) {
        const next = iterator.next();
        if (next === 0) {
          path[path.length - 1] += 1;
          advanced = true;
          break;
        }
        if (next === 1 && iterator.prev() !== 0) {
          throw new Error("MuPDF outline iterator could not restore an item after probing its siblings.");
        }
        if (path.length === 1) break;
        if (iterator.up() !== 0) {
          throw new Error("MuPDF outline iterator could not return to a parent during bounded inspection.");
        }
        parentIndexes.pop();
        path = path.slice(0, -1);
      }
      if (!advanced) break;
    }
  } finally {
    iterator.destroy();
  }
  const records = snapshots.map((snapshot) => ({
    kind: "mupdfOutline",
    id: outlineId(snapshot),
    depth: snapshot.path.length - 1,
    ...snapshot,
    updateCapability: updateCapability(snapshot),
    snapshot,
  }));
  return { count: records.length, records };
}

function validateOperationFields(operation) {
  for (const field of Object.keys(operation)) {
    if (!UPDATE_OPERATION_FIELDS.has(field)) throw new Error(`update_outline contains unsupported field: ${field}.`);
  }
}

function expectedSnapshot(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new Error("update_outline expected must be the complete snapshot from one mupdfOutline record.");
  }
  for (const field of Object.keys(value)) {
    if (!OUTLINE_SNAPSHOT_FIELDS.includes(field)) throw new Error(`update_outline expected contains unsupported snapshot field: ${field}.`);
  }
  for (const field of OUTLINE_SNAPSHOT_FIELDS) {
    if (!Object.hasOwn(value, field)) throw new Error(`update_outline expected must include snapshot field: ${field}.`);
  }
  if (!Array.isArray(value.path) || !value.path.length || value.path.length > OUTLINE_MAX_DEPTH
      || !value.path.every((index) => Number.isSafeInteger(index) && index >= 0)) {
    throw new Error(`update_outline expected.path must contain 1 through ${OUTLINE_MAX_DEPTH} non-negative safe integer indexes.`);
  }
  for (const field of ["title", "uri"]) {
    if (value[field] !== null && typeof value[field] !== "string") {
      throw new Error(`update_outline expected.${field} must be a string or null.`);
    }
  }
  if (typeof value.open !== "boolean") throw new Error("update_outline expected.open must be a boolean.");
  if (value.page !== null && (!Number.isSafeInteger(value.page) || value.page < 1)) {
    throw new Error("update_outline expected.page must be a positive 1-based page number or null.");
  }
  if (!Number.isSafeInteger(value.childCount) || value.childCount < 0) {
    throw new Error("update_outline expected.childCount must be a non-negative safe integer.");
  }
  return {
    path: [...value.path],
    title: value.title,
    uri: value.uri,
    open: value.open,
    page: value.page,
    childCount: value.childCount,
  };
}

function outlineTitle(value) {
  if (typeof value !== "string" || !value.trim()) {
    throw new Error("update_outline patch.title must be a non-empty string.");
  }
  if (value.length > OUTLINE_MAX_TITLE_LENGTH) {
    throw new Error(`update_outline patch.title exceeds ${OUTLINE_MAX_TITLE_LENGTH} UTF-16 code units.`);
  }
  if (/[\u0000-\u001f\u007f]/u.test(value)) {
    throw new Error("update_outline patch.title must not contain control characters.");
  }
  return value;
}

function outlinePatch(value, capability) {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new Error("update_outline patch must be a non-empty object.");
  }
  const fields = Object.keys(value);
  if (!fields.length) throw new Error("update_outline patch must be a non-empty object.");
  for (const field of fields) {
    if (!UPDATE_PATCH_FIELDS.has(field)) throw new Error(`update_outline patch contains unsupported field: ${field}.`);
    if (!capability.mutableFields.includes(field)) {
      const blocked = capability.blockedFields.find((entry) => entry.field === field);
      throw new Error(`update_outline cannot change ${field}: ${blocked?.reason || "field is not mutable"}.`);
    }
  }
  const patch = {};
  if (Object.hasOwn(value, "title")) patch.title = outlineTitle(value.title);
  if (Object.hasOwn(value, "open")) {
    if (typeof value.open !== "boolean") throw new Error("update_outline patch.open must be a boolean.");
    patch.open = value.open;
  }
  return patch;
}

function samePath(left, right) {
  return left.length === right.length && left.every((value, index) => value === right[index]);
}

function iteratorAtPath(document, path) {
  const iterator = document.outlineIterator();
  try {
    if (!iterator.item()) throw new Error(`update_outline could not resolve outline path ${pathLabel(path)}.`);
    for (let depth = 0; depth < path.length; depth += 1) {
      for (let index = 0; index < path[depth]; index += 1) {
        if (iterator.next() !== 0) throw new Error(`update_outline could not resolve outline path ${pathLabel(path)}.`);
      }
      if (depth < path.length - 1 && iterator.down() !== 0) {
        throw new Error(`update_outline could not resolve outline path ${pathLabel(path)}.`);
      }
    }
    return iterator;
  } catch (error) {
    iterator.destroy();
    throw error;
  }
}

function assertGraphTransition(before, after, targetPath, expectedAfter) {
  if (before.length !== after.length) throw new Error("MuPDF changed outline count during update_outline.");
  for (let index = 0; index < before.length; index += 1) {
    const prior = before[index];
    const current = after[index];
    if (!samePath(prior.path, current.path)) throw new Error("MuPDF changed outline order or topology during update_outline.");
    if (samePath(prior.path, targetPath)) {
      if (!sameValue(current.snapshot, expectedAfter)) {
        throw new Error("MuPDF did not preserve the requested fixed-topology outline update.");
      }
    } else if (!sameValue(prior.snapshot, current.snapshot)) {
      throw new Error(`MuPDF changed non-target outline ${prior.id} during update_outline.`);
    }
  }
}

export function applySourceBoundOutlineUpdate(document, operation, context = {}) {
  validateOperationFields(operation);
  if (!/^[a-f0-9]{64}$/u.test(String(operation.sourceSha256 || "")) || operation.sourceSha256 !== context.sourceSha256) {
    throw new Error("update_outline sourceSha256 must exactly match PdfFile.inspectPdf(...).summary.sourceSha256 for the current input bytes.");
  }
  const expected = expectedSnapshot(operation.expected);
  const expectedId = outlineId(expected);
  if (operation.outlineId !== expectedId) {
    throw new Error("update_outline outlineId must exactly match the inspect-derived expected snapshot.");
  }
  const beforeProfile = documentOutlineProfile(document, context.limits);
  const matched = beforeProfile.records.find((record) => samePath(record.path, expected.path));
  if (!matched) {
    throw new Error(`update_outline could not find source-bound outline ${operation.outlineId}; re-inspect the current source PDF before retrying.`);
  }
  if (matched.id !== operation.outlineId || !sameValue(matched.snapshot, expected)) {
    throw new Error(`update_outline precondition did not match ${operation.outlineId}; refusing a stale or ambiguous mutation.`);
  }
  const patch = outlinePatch(operation.patch, matched.updateCapability);
  const expectedAfter = { ...expected, ...patch };
  if (sameValue(expected, expectedAfter)) throw new Error("update_outline patch would not change the inspected outline.");

  const iterator = iteratorAtPath(document, expected.path);
  try {
    const item = iterator.item();
    const iteratorSnapshot = {
      title: typeof item?.title === "string" ? item.title : null,
      uri: typeof item?.uri === "string" ? item.uri : null,
      open: Boolean(item?.open),
    };
    if (!sameValue(iteratorSnapshot, { title: expected.title, uri: expected.uri, open: expected.open })) {
      throw new Error(`update_outline iterator state did not match ${operation.outlineId}; refusing an ambiguous mutation.`);
    }
    iterator.update({
      title: expectedAfter.title ?? undefined,
      uri: expected.uri ?? undefined,
      open: expectedAfter.open,
    });
  } finally {
    iterator.destroy();
  }

  const afterProfile = documentOutlineProfile(document, context.limits);
  assertGraphTransition(beforeProfile.records, afterProfile.records, expected.path, expectedAfter);
  const updated = afterProfile.records.find((record) => samePath(record.path, expected.path));
  return {
    type: "update_outline",
    outlineId: matched.id,
    updatedOutlineId: updated.id,
    path: [...expected.path],
    patch,
    matched: expected,
    updated: updated.snapshot,
  };
}
