import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

const root = path.resolve(import.meta.dirname, "..");
const script = path.join(root, "skills", "pdf", "skills", "pdf", "scripts", "pypdf_edit.py");
const python = process.env.OFFICE_KIT_PDF_PROVIDER_PYTHON;

if (!python) {
  console.log("pypdf static-form smoke skipped (set OFFICE_KIT_PDF_PROVIDER_PYTHON)");
  process.exit(0);
}

function run(args, { status = 0 } = {}) {
  const result = spawnSync(python, args, {
    cwd: root,
    encoding: "utf8",
    env: { ...process.env, OFFICE_KIT_PDF_PROVIDER_PYTHON: python },
    maxBuffer: 4 * 1024 * 1024,
  });
  if (result.error) throw result.error;
  assert.equal(
    result.status,
    status,
    `Expected ${path.basename(args[0])} to exit ${status}, got ${result.status}.\nstdout:\n${result.stdout}\nstderr:\n${result.stderr}`,
  );
  return result;
}

function jsonResult(result) {
  return JSON.parse(result.stdout);
}

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

const temporary = await fs.mkdtemp(path.join(os.tmpdir(), "office-kit-pypdf-flatten-"));
try {
  const source = path.join(temporary, "interactive-source.pdf");
  const flattened = path.join(temporary, "static-output.pdf");
  const incremental = path.join(temporary, "invalid-incremental.pdf");
  const orphanSource = path.join(temporary, "orphan-widget-source.pdf");
  const orphanOutput = path.join(temporary, "orphan-widget-output.pdf");
  run(["-c", [
    "from pathlib import Path",
    "from reportlab.pdfgen import canvas",
    "from pypdf import PdfReader, PdfWriter",
    "from pypdf.annotations import Text",
    "import sys",
    "source = Path(sys.argv[1])",
    "base = source.with_suffix('.base.pdf')",
    "document = canvas.Canvas(str(base))",
    "document.drawString(72, 720, 'Static delivery form')",
    "form = document.acroForm",
    "form.textfield(name='sender.city', tooltip='City', x=72, y=670, width=180, height=24, value='')",
    "form.textfield(name='prior.reference', tooltip='Existing reference', x=72, y=640, width=180, height=24, value='Preserve this value')",
    "form.radio(name='company_type', value='LLC', selected=False, x=72, y=620, buttonStyle='circle')",
    "form.radio(name='company_type', value='Corporation', selected=False, x=140, y=620, buttonStyle='circle')",
    "form.checkbox(name='terms_ack', checked=False, x=72, y=570, buttonStyle='check')",
    "document.save()",
    "writer = PdfWriter(clone_from=PdfReader(str(base), strict=True))",
    "writer.add_annotation(0, Text(rect=(360, 660, 390, 690), text='Retain this review note'))",
    "writer.write(str(source))",
    "base.unlink()",
  ].join("\n"), source]);
  const sourceBytes = await fs.readFile(source);
  const sourceHash = sha256(sourceBytes);

  run(["-c", [
    "from pypdf import PdfReader, PdfWriter",
    "import sys",
    "reader = PdfReader(sys.argv[1], strict=True)",
    "writer = PdfWriter(clone_from=reader)",
    "acroform = writer.root_object['/AcroForm'].get_object()",
    "fields = acroform['/Fields']",
    "for index, reference in enumerate(fields):",
    "    if str(reference.get_object().get('/T', '')) == 'prior.reference':",
    "        del fields[index]",
    "        break",
    "else:",
    "    raise RuntimeError('fixture field not found')",
    "writer.write(sys.argv[2])",
  ].join("\n"), source, orphanSource]);
  const orphanRejected = run([
    script, "fill-form", orphanSource, orphanOutput,
    "--strategy", "rewrite", "--flatten", "--field", "sender.city=Shanghai",
  ], { status: 2 });
  assert.match(orphanRejected.stderr, /orphan or unmodeled Widget annotation/);
  await assert.rejects(fs.access(orphanOutput));

  const rejected = run([
    script, "fill-form", source, incremental,
    "--strategy", "incremental", "--flatten", "--field", "sender.city=Shanghai",
  ], { status: 2 });
  assert.match(rejected.stderr, /flatten requires rewrite/);
  await assert.rejects(fs.access(incremental));

  const result = jsonResult(run([
    script, "fill-form", source, flattened,
    "--strategy", "rewrite", "--flatten",
    "--field", "sender.city=Shanghai",
    "--field", "company_type=LLC",
    "--field", "terms_ack=Yes",
  ]));
  assert.equal(result.operation.flatten, true);
  assert.equal(result.formValidation.mode, "static");
  assert.equal(result.formValidation.source.acroFormPresent, true);
  assert.equal(result.formValidation.source.widgetCount, 5);
  assert.deepEqual(result.operation.staticPaintedFields, ["company_type", "prior.reference", "sender.city", "terms_ack"]);
  assert.deepEqual(result.formValidation.output, {
    acroFormPresent: false,
    fieldCount: 0,
    fieldTreeRoots: 0,
    widgetCount: 0,
  });
  assert.equal(result.formValidation.allWidgetsRemoved, true);
  assert.equal(result.formValidation.fieldTreeRemoved, true);
  assert.equal(sha256(await fs.readFile(source)), sourceHash, "flattening must not mutate the source PDF");

  const staticEvidence = jsonResult(run(["-c", [
    "import json, sys",
    "from pypdf import PdfReader",
    "reader = PdfReader(sys.argv[1], strict=True)",
    "root = reader.trailer['/Root']",
    "subtypes = []",
    "for page in reader.pages:",
    "    for reference in page.get('/Annots', []) or []:",
    "        subtypes.append(str(reference.get_object().get('/Subtype', '')))",
    "print(json.dumps({",
    "  'acroFormPresent': '/AcroForm' in root,",
    "  'fields': sorted((reader.get_fields() or {}).keys()),",
    "  'widgetCount': subtypes.count('/Widget'),",
    "  'textNotes': subtypes.count('/Text'),",
    "  'text': reader.pages[0].extract_text() or '',",
    "}))",
  ].join("\n"), flattened]));
  assert.equal(staticEvidence.acroFormPresent, false);
  assert.deepEqual(staticEvidence.fields, []);
  assert.equal(staticEvidence.widgetCount, 0);
  assert.equal(staticEvidence.textNotes, 1, "flattening must retain non-Widget annotations");
  assert.match(staticEvidence.text, /Shanghai/);
  assert.match(staticEvidence.text, /Preserve this value/, "flattening must paint untouched field values before removing Widgets");

  const inspection = jsonResult(run([script, "inspect", flattened]));
  assert.equal(inspection.summary.acroFormPresent, false);
  assert.equal(inspection.summary.fieldTreeRoots, 0);
  assert.equal(inspection.summary.widgets, 0);
  assert.deepEqual(inspection.formStructure, result.formValidation.output);
} finally {
  await fs.rm(temporary, { force: true, recursive: true });
}

console.log("pypdf static-form flatten smoke ok");
