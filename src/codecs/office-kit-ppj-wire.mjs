import { BinaryReader, BinaryWriter, WireType } from "@bufbuild/protobuf/wire";

const EMPTY_BYTES = new Uint8Array();

export function encodePpjCodecRequest(request, { omitFile = false } = {}) {
  const writer = new BinaryWriter();
  writeUint32(writer, 1, request.protocolVersion);
  writeUint32(writer, 2, request.operation);
  writeUint32(writer, 3, request.family);
  if (!omitFile) writeBytes(writer, 5, request.file);
  writeMessage(writer, 6, request.limits, writeLimits);
  writeMessage(writer, 12, request.presentationProgram, writeProgramRequest);
  return writer.finish();
}

export function decodePpjCodecResponse(bytes) {
  const response = {
    protocolVersion: 0,
    ok: false,
    file: EMPTY_BYTES,
    diagnostics: [],
    presentationProgram: undefined,
  };
  readMessage(bytes, (field, wire, reader) => {
    if (field === 1) response.protocolVersion = readUint32(reader, wire);
    else if (field === 2) response.ok = readBool(reader, wire);
    else if (field === 3) response.file = readBytes(reader, wire);
    else if (field === 5) response.diagnostics.push(readDiagnostic(readBytes(reader, wire)));
    else if (field === 9) response.presentationProgram = readProgramResult(readBytes(reader, wire));
    else return false;
    return true;
  });
  return response;
}

function writeLimits(writer, limits) {
  writeUint64(writer, 1, limits.maxInputBytes);
  writeUint64(writer, 2, limits.maxUncompressedBytes);
  writeUint32(writer, 3, limits.maxParts);
  writeUint32(writer, 4, limits.maxSheets);
  writeUint64(writer, 5, limits.maxCells);
  writeUint32(writer, 6, limits.maxCompressionRatio);
}

function writeProgramRequest(writer, program) {
  writeBytes(writer, 1, program.programJson);
  for (const asset of program.assets ?? []) writeMessage(writer, 2, asset, writeAsset);
  writeBool(writer, 3, program.includeNodeMap);
  writeString(writer, 4, program.sourceUri);
  writeString(writer, 5, program.assetRootUri);
  writeBool(writer, 6, program.validationOnly);
}

function writeAsset(writer, asset) {
  writeString(writer, 1, asset.id);
  writeString(writer, 2, asset.fileName);
  writeString(writer, 3, asset.contentType);
  writeBytes(writer, 4, asset.data);
  writeString(writer, 5, asset.sha256);
}

function readProgramResult(bytes) {
  const program = {
    programJson: EMPTY_BYTES,
    programSha256: "",
    nodeMapJson: EMPTY_BYTES,
    sourceSha256: "",
    outputSha256: "",
    changedParts: [],
    assets: [],
    restoredEmbeddedProgram: false,
    sourceBound: false,
    expandedElementCount: 0,
    changedNodeIds: [],
    originalProgramJson: EMPTY_BYTES,
  };
  readMessage(bytes, (field, wire, reader) => {
    if (field === 1) program.programJson = readBytes(reader, wire);
    else if (field === 2) program.programSha256 = readString(reader, wire);
    else if (field === 3) program.nodeMapJson = readBytes(reader, wire);
    else if (field === 4) program.sourceSha256 = readString(reader, wire);
    else if (field === 5) program.outputSha256 = readString(reader, wire);
    else if (field === 6) program.changedParts.push(readString(reader, wire));
    else if (field === 7) program.assets.push(readAsset(readBytes(reader, wire)));
    else if (field === 8) program.restoredEmbeddedProgram = readBool(reader, wire);
    else if (field === 9) program.sourceBound = readBool(reader, wire);
    else if (field === 10) program.expandedElementCount = readUint32(reader, wire);
    else if (field === 11) program.changedNodeIds.push(readString(reader, wire));
    else if (field === 12) program.originalProgramJson = readBytes(reader, wire);
    else return false;
    return true;
  });
  return program;
}

function readAsset(bytes) {
  const asset = { id: "", fileName: "", contentType: "", data: EMPTY_BYTES, sha256: "" };
  readMessage(bytes, (field, wire, reader) => {
    if (field === 1) asset.id = readString(reader, wire);
    else if (field === 2) asset.fileName = readString(reader, wire);
    else if (field === 3) asset.contentType = readString(reader, wire);
    else if (field === 4) asset.data = readBytes(reader, wire);
    else if (field === 5) asset.sha256 = readString(reader, wire);
    else return false;
    return true;
  });
  return asset;
}

function readDiagnostic(bytes) {
  const diagnostic = { severity: 0, code: "", message: "", sourcePath: "", sourceIdentity: "" };
  readMessage(bytes, (field, wire, reader) => {
    if (field === 1) diagnostic.severity = readUint32(reader, wire);
    else if (field === 2) diagnostic.code = readString(reader, wire);
    else if (field === 3) diagnostic.message = readString(reader, wire);
    else if (field === 4) diagnostic.sourcePath = readString(reader, wire);
    else if (field === 5) diagnostic.sourceIdentity = readString(reader, wire);
    else return false;
    return true;
  });
  return diagnostic;
}

function readMessage(bytes, consume) {
  const reader = new BinaryReader(bytes);
  while (reader.pos < reader.len) {
    const [field, wire] = reader.tag();
    if (!consume(field, wire, reader)) reader.skip(wire, field);
  }
  if (reader.pos !== reader.len) throw new RangeError("Protobuf message ended outside its declared boundary.");
}

function writeMessage(writer, field, value, write) {
  if (value == null) return;
  writer.tag(field, WireType.LengthDelimited).fork();
  write(writer, value);
  writer.join();
}

function writeUint32(writer, field, value) {
  if (!value) return;
  writer.tag(field, WireType.Varint).uint32(value);
}

function writeUint64(writer, field, value) {
  if (value == null || value === 0 || value === 0n) return;
  writer.tag(field, WireType.Varint).uint64(value);
}

function writeBool(writer, field, value) {
  if (!value) return;
  writer.tag(field, WireType.Varint).bool(value);
}

function writeBytes(writer, field, value) {
  if (!(value instanceof Uint8Array) || value.byteLength === 0) return;
  writer.tag(field, WireType.LengthDelimited).bytes(value);
}

function writeString(writer, field, value) {
  if (typeof value !== "string" || value.length === 0) return;
  writer.tag(field, WireType.LengthDelimited).string(value);
}

function readUint32(reader, wire) {
  requireWire(wire, WireType.Varint);
  return reader.uint32();
}

function readBool(reader, wire) {
  requireWire(wire, WireType.Varint);
  return reader.bool();
}

function readBytes(reader, wire) {
  requireWire(wire, WireType.LengthDelimited);
  return reader.bytes();
}

function readString(reader, wire) {
  requireWire(wire, WireType.LengthDelimited);
  return reader.string(true);
}

function requireWire(actual, expected) {
  if (actual !== expected) throw new TypeError(`Unexpected protobuf wire type ${actual}; expected ${expected}.`);
}
