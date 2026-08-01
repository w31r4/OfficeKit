import assert from "node:assert/strict";
import crypto from "node:crypto";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import sharp from "sharp";

const repoRoot = path.resolve(import.meta.dirname, "..");
const skillRoot = path.join(repoRoot, "skills", "pdf", "skills", "pdf");
const verifier = path.join(skillRoot, "scripts", "pyhanko_provider.py");
const filler = path.join(skillRoot, "scripts", "pyhanko_certified_form_fill.py");
const mupdf = path.join(skillRoot, "scripts", "mupdf.mjs");

function run(executable, args, options = {}) {
  const result = spawnSync(executable, args, {
    cwd: options.cwd || repoRoot,
    encoding: "utf8",
    env: { ...process.env, PYTHONDONTWRITEBYTECODE: "1", ...options.env },
    maxBuffer: 24 * 1024 * 1024,
  });
  if (options.status !== undefined) {
    assert.equal(result.status, options.status, `${executable} ${args.join(" ")}\nSTDOUT:\n${result.stdout}\nSTDERR:\n${result.stderr}`);
  }
  return result;
}

function jsonResult(result, stream = "stdout") {
  const text = result[stream]?.trim();
  assert.ok(text, `expected JSON on ${stream}\nSTDOUT:\n${result.stdout}\nSTDERR:\n${result.stderr}`);
  return JSON.parse(text);
}

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

async function staticAppearanceBounds(pngPath) {
  const { data, info } = await sharp(pngPath).removeAlpha().raw().toBuffer({ resolveWithObject: true });
  const scale = info.width / 612;
  const left = Math.round(220 * scale) + 4;
  const right = Math.round(370 * scale) - 4;
  const top = Math.round((792 - 684) * scale) + 4;
  const bottom = Math.round((792 - 660) * scale) - 4;
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -1;
  let maxY = -1;
  for (let y = top; y <= bottom; y += 1) {
    for (let x = left; x <= right; x += 1) {
      const offset = (y * info.width + x) * 3;
      if (data[offset] < 90 && data[offset + 1] < 90 && data[offset + 2] < 90) {
        minX = Math.min(minX, x);
        minY = Math.min(minY, y);
        maxX = Math.max(maxX, x);
        maxY = Math.max(maxY, y);
      }
    }
  }
  return { left, right, top, bottom, minX, minY, maxX, maxY };
}

function supportedPyHanko(executable) {
  if (!executable) return false;
  return run(executable, [
    "-c",
    "from importlib.metadata import version; v=tuple(int(x) for x in version('pyHanko').split('.')[:3]); raise SystemExit(0 if (0,35,0) <= v < (0,36,0) else 1)",
  ]).status === 0;
}

const configuredPython = process.env.OFFICE_KIT_PYHANKO_TEST_PYTHON;
if (configuredPython) {
  assert.ok(supportedPyHanko(configuredPython), `OFFICE_KIT_PYHANKO_TEST_PYTHON must provide pyHanko 0.35.x: ${configuredPython}`);
}
const managedPython = configuredPython || (supportedPyHanko("python3") ? "python3" : null);

if (!managedPython) {
  console.log("pyHanko certified form fill smoke ok (real provider skipped: set OFFICE_KIT_PYHANKO_TEST_PYTHON)");
  process.exit(0);
}

const managedEnv = {
  OFFICE_KIT_PDF_PROVIDER_PYTHON: managedPython,
  PYTHONNOUSERSITE: "1",
};

const configuredFoundationPython = process.env.OFFICE_KIT_PDF_PROVIDER_PYTHON;
function supportsFoundation(executable) {
  if (!executable) return false;
  return run(executable, ["-c", "import pypdf, reportlab"]).status === 0;
}

if (configuredPython) {
  assert.ok(
    configuredFoundationPython && supportsFoundation(configuredFoundationPython),
    "OFFICE_KIT_PDF_PROVIDER_PYTHON must provide the pypdf/reportlab foundation runtime when pyHanko is enabled",
  );
}
const foundationPython = configuredFoundationPython || (supportsFoundation(managedPython) ? managedPython : null);
if (!foundationPython) {
  console.log("pyHanko certified form fill smoke ok (real provider skipped: set OFFICE_KIT_PDF_PROVIDER_PYTHON and OFFICE_KIT_PYHANKO_TEST_PYTHON)");
  process.exit(0);
}

function runManaged(args, options = {}) {
  return run(managedPython, args, { ...options, env: { ...managedEnv, ...options.env } });
}

function runFoundation(args, options = {}) {
  return run(foundationPython, args, { ...options, env: { PYTHONNOUSERSITE: "1", ...options.env } });
}

const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), "office-kit-pyhanko-certified-form-"));
try {
  const sourceBuilder = path.join(tempRoot, "build_source.py");
  await fs.writeFile(sourceBuilder, String.raw`
from pathlib import Path
import sys

from pypdf import PdfReader, PdfWriter
from pypdf.generic import DictionaryObject, NameObject, NumberObject, TextStringObject
from reportlab.pdfgen import canvas

root = Path(sys.argv[1])
base = root / "base.pdf"
source = root / "source.pdf"
document = canvas.Canvas(str(base), pagesize=(612, 792), invariant=1)
document.setTitle("OfficeKit certified P=2 form fixture")
document.setAuthor("OfficeKit test")
document.setFont("Helvetica-Bold", 16)
document.drawString(72, 716, "Certification-approved amount")
document.setFont("Helvetica", 10)
document.drawString(72, 676, "Approved amount (USD):")
document.acroForm.textfield(
    name="ApprovedAmount", x=220, y=660, width=150, height=24,
    value="", fontName="Helvetica", fontSize=10,
    borderWidth=1, forceBorder=True,
)
document.drawString(72, 622, "Locked reference: LOCKED-9000")
document.save()

reader = PdfReader(str(base), strict=True)
writer = PdfWriter()
writer.clone_document_from_reader(reader)
acroform = writer._root_object["/AcroForm"]
if hasattr(acroform, "get_object"):
    acroform = acroform.get_object()
fields = acroform["/Fields"]
if hasattr(fields, "get_object"):
    fields = fields.get_object()
locked = DictionaryObject({
    NameObject("/FT"): NameObject("/Tx"),
    NameObject("/T"): TextStringObject("LockedAmount"),
    NameObject("/V"): TextStringObject("LOCKED-9000"),
    NameObject("/Ff"): NumberObject(1),
})
fields.append(writer._add_object(locked))
with source.open("wb") as handle:
    writer.write(handle)
`, "utf8");
  runFoundation([sourceBuilder, tempRoot], { status: 0 });

  const signerBuilder = path.join(tempRoot, "sign_source.py");
  await fs.writeFile(signerBuilder, String.raw`
from datetime import datetime, timedelta, timezone
from pathlib import Path
import sys

from cryptography import x509
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import rsa
from cryptography.x509.oid import NameOID
from pyhanko.pdf_utils.incremental_writer import IncrementalPdfFileWriter
from pyhanko.sign import signers
from pyhanko.sign.fields import FieldMDPAction, FieldMDPSpec, MDPPerm, SigFieldSpec

root = Path(sys.argv[1])
now = datetime.now(timezone.utc)
key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
name = x509.Name([
    x509.NameAttribute(NameOID.COMMON_NAME, "OfficeKit P=2 Test Root"),
    x509.NameAttribute(NameOID.ORGANIZATION_NAME, "OfficeKit"),
    x509.NameAttribute(NameOID.COUNTRY_NAME, "US"),
])
certificate = (
    x509.CertificateBuilder()
    .subject_name(name).issuer_name(name).public_key(key.public_key())
    .serial_number(x509.random_serial_number())
    .not_valid_before(now - timedelta(days=1)).not_valid_after(now + timedelta(days=3650))
    .add_extension(x509.BasicConstraints(ca=True, path_length=None), critical=True)
    .add_extension(x509.KeyUsage(
        digital_signature=True, content_commitment=True, key_encipherment=False,
        data_encipherment=False, key_agreement=False, key_cert_sign=True,
        crl_sign=True, encipher_only=None, decipher_only=None,
    ), critical=True)
    .sign(key, hashes.SHA256())
)
key_path = root / "signer-key.pem"
cert_path = root / "root.pem"
key_path.write_bytes(key.private_bytes(
    serialization.Encoding.PEM, serialization.PrivateFormat.PKCS8, serialization.NoEncryption()
))
cert_path.write_bytes(certificate.public_bytes(serialization.Encoding.PEM))
signer = signers.SimpleSigner.load(key_path, cert_path)
with (root / "source.pdf").open("rb") as source, (root / "certified.pdf").open("wb") as output:
    writer = IncrementalPdfFileWriter(source, strict=True)
    signers.sign_pdf(
        writer,
        signers.PdfSignatureMetadata(
            field_name="Certification", certify=True, docmdp_permissions=MDPPerm.FILL_FORMS,
        ),
        signer=signer,
        new_field_spec=SigFieldSpec(
            sig_field_name="Certification",
            field_mdp_spec=FieldMDPSpec(FieldMDPAction.INCLUDE, fields=["LockedAmount"]),
        ),
        output=output,
    )
`, "utf8");
  runManaged([signerBuilder, tempRoot], { status: 0 });

  const source = path.join(tempRoot, "certified.pdf");
  const rootCertificate = path.join(tempRoot, "root.pem");
  const sourceBytes = await fs.readFile(source);
  const sourceHash = sha256(sourceBytes);
  const output = path.join(tempRoot, "approved-amount.pdf");

  const probe = jsonResult(runManaged([filler, "probe"], { status: 0 }));
  assert.equal(probe.schema, "office-kit.pyhanko-certified-form-fill.v1");
  assert.equal(probe.operation, "fill-certified-docmdp-p2-text-field");
  assert.equal(probe.supported.docmdpPermission, "fill-forms");
  assert.equal(probe.silentFallback, false);

  const baseline = jsonResult(runManaged([
    verifier, "verify", source, "--expected-sha256", sourceHash,
    "--trust-policy", "explicit-roots", "--trust-root", rootCertificate,
    "--require-signature", "--require-all-integrity-valid", "--require-all-trusted",
    "--require-docmdp-compliant", "--require-all-bottom-line",
  ], { status: 0 }));
  assert.equal(baseline.signatures[0].docMDP.permission, "fill-forms");
  assert.equal(baseline.signatures[0].fieldMDP.present, true);
  assert.equal(baseline.signatures[0].fieldMDP.action, "include");
  assert.deepEqual(baseline.signatures[0].fieldMDP.fields, ["LockedAmount"]);
  assert.equal(baseline.signatures[0].coverage, "entire-file");

  const fill = jsonResult(runManaged([
    filler, "fill", source, output,
    "--expected-source-sha256", sourceHash,
    "--trust-root", rootCertificate,
    "--field", "ApprovedAmount", "--value", "12500.00",
    "--expected-signature-field", "Certification",
    "--expected-locked-field", "LockedAmount", "--expected-locked-value", "LOCKED-9000",
    "--caller-isolated",
  ], { status: 0 }));
  assert.equal(fill.ok, true);
  assert.equal(fill.operationCompleted, true);
  assert.equal(fill.savePolicy.strategy, "incremental");
  assert.equal(fill.savePolicy.sourcePrefixPreserved, true);
  assert.equal(fill.savePolicy.revisionsAfter, fill.savePolicy.revisionsBefore + 1);
  assert.equal(fill.field.target, "ApprovedAmount");
  assert.equal(fill.field.value, "12500.00");
  assert.deepEqual(fill.field.appearance, { mode: "static", verticalAlignment: "middle", innerMarginPoints: 3 });
  assert.equal(fill.field.after.readOnly, true);
  assert.equal(fill.field.after.hasNormalAppearance, true);
  assert.equal(fill.field.after.hasDefaultAppearance, false);
  assert.equal(fill.field.nonTargetFieldsUnchanged, true);
  assert.equal(fill.signature.postflight.modificationLevel, "form-filling");
  assert.deepEqual(fill.signature.postflight.changedFormFields, ["ApprovedAmount"]);
  assert.equal(fill.signature.postflight.docMDPCompliant, true);
  assert.equal(fill.transaction.noReplace, true);
  assert.deepEqual(await fs.readFile(source), sourceBytes, "source bytes must remain immutable");
  const outputBytes = await fs.readFile(output);
  assert.deepEqual(outputBytes.subarray(0, sourceBytes.length), sourceBytes, "output must preserve every source byte as a prefix");

  const outputValidation = jsonResult(runManaged([
    verifier, "verify", output, "--expected-sha256", sha256(outputBytes),
    "--trust-policy", "explicit-roots", "--trust-root", rootCertificate,
    "--require-signature", "--require-all-integrity-valid", "--require-all-trusted",
    "--require-docmdp-compliant", "--require-all-bottom-line",
  ], { status: 0 }));
  assert.equal(outputValidation.signatures[0].coverage, "entire-revision");
  assert.equal(outputValidation.signatures[0].modificationLevel, "form-filling");
  assert.deepEqual(outputValidation.signatures[0].changedFormFields, ["ApprovedAmount"]);

  const pypdfEvidence = runFoundation(["-c", String.raw`
from pypdf import PdfReader
import json, sys
reader = PdfReader(sys.argv[1], strict=True)
fields = reader.get_fields()
target = fields["ApprovedAmount"]
locked = fields["LockedAmount"]
target_widget = next(
  (
    annotation.get_object()
    for page in reader.pages
    for annotation in page.get("/Annots", [])
    if str(annotation.get_object().get("/T", "")) == "ApprovedAmount"
  ),
  None,
)
print(json.dumps({
  "target": str(target.get("/V")),
  "targetReadOnly": bool(int(target.get("/Ff", 0)) & 1),
  "targetAppearance": target_widget is not None and target_widget.get("/AP") is not None,
  "locked": str(locked.get("/V")),
  "lockedReadOnly": bool(int(locked.get("/Ff", 0)) & 1),
}, sort_keys=True))
`, output], { status: 0 });
  assert.deepEqual(JSON.parse(pypdfEvidence.stdout), {
    locked: "LOCKED-9000",
    lockedReadOnly: true,
    target: "12500.00",
    targetAppearance: true,
    targetReadOnly: true,
  });
  const extracted = run("pdftotext", [output, "-"]);
  if (extracted.status === 0) {
    assert.match(extracted.stdout, /12500\.00/, "static field appearance must be extractable/visible to a native PDF consumer");
  } else {
    assert.equal(extracted.error?.code, "ENOENT", `pdftotext must either render the static field or be unavailable: ${extracted.stderr}`);
  }
  const rendered = path.join(tempRoot, "approved-amount.png");
  run(process.execPath, [mupdf, "render", output, rendered, "--page", "1", "--dpi", "144"], { status: 0 });
  const appearanceBounds = await staticAppearanceBounds(rendered);
  assert.ok(appearanceBounds.maxX >= appearanceBounds.minX, `static field appearance must paint text: ${JSON.stringify(appearanceBounds)}`);
  assert.ok(appearanceBounds.minY >= appearanceBounds.top + 5, `static field appearance must not touch its top border: ${JSON.stringify(appearanceBounds)}`);
  assert.ok(appearanceBounds.maxY <= appearanceBounds.bottom - 5, `static field appearance must not touch its bottom border: ${JSON.stringify(appearanceBounds)}`);

  const collision = runManaged([
    filler, "fill", source, output,
    "--expected-source-sha256", sourceHash,
    "--trust-root", rootCertificate,
    "--field", "ApprovedAmount", "--value", "12500.00",
    "--expected-signature-field", "Certification",
    "--expected-locked-field", "LockedAmount", "--expected-locked-value", "LOCKED-9000",
    "--trusted-input",
  ], { status: 2 });
  assert.match(jsonResult(collision, "stderr").error, /already exists and will not be replaced/);

  const staleOutput = path.join(tempRoot, "stale.pdf");
  const stale = runManaged([
    filler, "fill", source, staleOutput,
    "--expected-source-sha256", "0".repeat(64),
    "--trust-root", rootCertificate,
    "--field", "ApprovedAmount", "--value", "12500.00",
    "--expected-signature-field", "Certification",
    "--expected-locked-field", "LockedAmount", "--expected-locked-value", "LOCKED-9000",
    "--trusted-input",
  ], { status: 2 });
  assert.match(jsonResult(stale, "stderr").error, /source SHA-256 mismatch/);
  await assert.rejects(fs.access(staleOutput));

  const lockedOutput = path.join(tempRoot, "locked-attempt.pdf");
  const locked = runManaged([
    filler, "fill", source, lockedOutput,
    "--expected-source-sha256", sourceHash,
    "--trust-root", rootCertificate,
    "--field", "LockedAmount", "--value", "12500.00",
    "--expected-signature-field", "Certification",
    "--expected-locked-field", "LockedAmount", "--expected-locked-value", "LOCKED-9000",
    "--trusted-input",
  ], { status: 2 });
  assert.match(jsonResult(locked, "stderr").error, /must be different/);
  await assert.rejects(fs.access(lockedOutput));

  const malformedOutput = path.join(tempRoot, "malformed-value.pdf");
  const malformed = runManaged([
    filler, "fill", source, malformedOutput,
    "--expected-source-sha256", sourceHash,
    "--trust-root", rootCertificate,
    "--field", "ApprovedAmount", "--value", "12500",
    "--expected-signature-field", "Certification",
    "--expected-locked-field", "LockedAmount", "--expected-locked-value", "LOCKED-9000",
    "--trusted-input",
  ], { status: 2 });
  assert.match(jsonResult(malformed, "stderr").error, /canonical non-negative decimal/);
  await assert.rejects(fs.access(malformedOutput));

  const symlinkOutput = path.join(tempRoot, "symlink-output.pdf");
  await fs.symlink(source, symlinkOutput);
  const symlink = runManaged([
    filler, "fill", source, symlinkOutput,
    "--expected-source-sha256", sourceHash,
    "--trust-root", rootCertificate,
    "--field", "ApprovedAmount", "--value", "12500.00",
    "--expected-signature-field", "Certification",
    "--expected-locked-field", "LockedAmount", "--expected-locked-value", "LOCKED-9000",
    "--trusted-input",
  ], { status: 2 });
  assert.match(jsonResult(symlink, "stderr").error, /symbolic link/);
} finally {
  await fs.rm(tempRoot, { recursive: true, force: true });
}

console.log("pyHanko certified form fill smoke ok");
