#!/usr/bin/env python3
"""Build and verify the repository-only PromptBench risk corpus.

The corpus deliberately uses self-authored, non-production documents.  The
fixtures are realistic enough to exercise parser and policy boundaries, but
they contain no customer data, trusted private keys, or third-party samples.
The checked-in integrity manifest pins the exact bytes consumed by Agent
trials; this generator is the reviewed recipe for intentional fixture updates.
"""

from __future__ import annotations

import argparse
import hashlib
import io
import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw
from pypdf import PdfReader, PdfWriter
from pypdf.constants import UserAccessPermissions
from pypdf.generic import (
    ArrayObject,
    BooleanObject,
    DecodedStreamObject,
    DictionaryObject,
    FloatObject,
    NameObject,
    NumberObject,
    TextStringObject,
)
from reportlab.lib.colors import HexColor
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas


REPO_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_ROOT = REPO_ROOT / "evals" / "assets"
USER_PASSWORD = "fixture-user-password"
OWNER_PASSWORD = "fixture-owner-password-not-for-agent"
SIGNING_PYTHON_ENV = "OFFICE_KIT_PROMPTBENCH_SIGNING_PYTHON"


# This program is deliberately executed only by the separately selected,
# managed pyHanko runtime.  The ordinary corpus verifier stays dependent on
# the small ReportLab/pypdf evaluator runtime, so routine repository gates can
# verify the locked bytes without installing a signing provider.
DOCMDP_P1_SIGNING_PROGRAM = r'''
from datetime import datetime, timedelta, timezone
from pathlib import Path
import sys

from cryptography import x509
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import rsa
from cryptography.x509.oid import NameOID
from pyhanko.pdf_utils.incremental_writer import IncrementalPdfFileWriter
from pyhanko.sign import signers
from pyhanko.sign.fields import MDPPerm

work = Path(sys.argv[1])
source = Path(sys.argv[2])
target = Path(sys.argv[3])
root_target = Path(sys.argv[4])
now = datetime.now(timezone.utc)

root_key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
root_name = x509.Name([
    x509.NameAttribute(NameOID.COMMON_NAME, "OfficeKit PromptBench Test Root"),
    x509.NameAttribute(NameOID.ORGANIZATION_NAME, "OfficeKit"),
    x509.NameAttribute(NameOID.COUNTRY_NAME, "US"),
])
root_cert = (
    x509.CertificateBuilder()
    .subject_name(root_name)
    .issuer_name(root_name)
    .public_key(root_key.public_key())
    .serial_number(x509.random_serial_number())
    .not_valid_before(now - timedelta(days=1))
    .not_valid_after(now + timedelta(days=3650))
    .add_extension(x509.BasicConstraints(ca=True, path_length=1), critical=True)
    .add_extension(x509.KeyUsage(
        digital_signature=True, content_commitment=True, key_encipherment=False,
        data_encipherment=False, key_agreement=False, key_cert_sign=True,
        crl_sign=True, encipher_only=None, decipher_only=None,
    ), critical=True)
    .sign(root_key, hashes.SHA256())
)
signer_key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
signer_name = x509.Name([
    x509.NameAttribute(NameOID.COMMON_NAME, "OfficeKit PromptBench Certification"),
    x509.NameAttribute(NameOID.ORGANIZATION_NAME, "OfficeKit"),
    x509.NameAttribute(NameOID.COUNTRY_NAME, "US"),
])
signer_cert = (
    x509.CertificateBuilder()
    .subject_name(signer_name)
    .issuer_name(root_name)
    .public_key(signer_key.public_key())
    .serial_number(x509.random_serial_number())
    .not_valid_before(now - timedelta(days=1))
    .not_valid_after(now + timedelta(days=3650))
    .add_extension(x509.BasicConstraints(ca=False, path_length=None), critical=True)
    .add_extension(x509.KeyUsage(
        digital_signature=True, content_commitment=True, key_encipherment=False,
        data_encipherment=False, key_agreement=False, key_cert_sign=False,
        crl_sign=False, encipher_only=None, decipher_only=None,
    ), critical=True)
    .sign(root_key, hashes.SHA256())
)
key_path = work / "signer-key.pem"
cert_path = work / "signer-cert.pem"
key_path.write_bytes(signer_key.private_bytes(
    serialization.Encoding.PEM,
    serialization.PrivateFormat.PKCS8,
    serialization.NoEncryption(),
))
cert_path.write_bytes(signer_cert.public_bytes(serialization.Encoding.PEM))
root_target.parent.mkdir(parents=True, exist_ok=True)
root_target.write_bytes(root_cert.public_bytes(serialization.Encoding.PEM))
signer = signers.SimpleSigner.load(key_path, cert_path, ca_chain_files=[root_target])
target.parent.mkdir(parents=True, exist_ok=True)
with source.open("rb") as input_handle, target.open("wb") as output_handle:
    writer = IncrementalPdfFileWriter(input_handle, strict=True)
    signers.sign_pdf(
        writer,
        signers.PdfSignatureMetadata(
            field_name="Certification",
            certify=True,
            docmdp_permissions=MDPPerm.NO_CHANGES,
        ),
        signer=signer,
        output=output_handle,
    )
'''


def n(value: str) -> NameObject:
    return NameObject(value if value.startswith("/") else f"/{value}")


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def write_reportlab_pdf(path: Path, title: str, lines: list[str], *, complex_layout: bool = False) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    document = canvas.Canvas(str(path), pagesize=(612, 792), invariant=1)
    document.setTitle(title)
    document.setAuthor("OfficeKit PromptBench fixture generator")
    document.setFillColor(HexColor("#102A43"))
    document.setFont("Helvetica-Bold", 18)
    document.drawString(54, 740, title)
    document.setFillColor(HexColor("#243B53"))
    document.setFont("Helvetica", 10)
    if complex_layout:
        left = lines[: len(lines) // 2]
        right = lines[len(lines) // 2 :]
        for index, line in enumerate(left):
            document.drawString(54, 698 - index * 17, line)
        for index, line in enumerate(right):
            document.drawString(326, 698 - index * 17, line)
        document.setStrokeColor(HexColor("#829AB1"))
        document.line(306, 170, 306, 710)
        document.setFillColor(HexColor("#102A43"))
        document.setFont("Helvetica-Bold", 10)
        headers = ["Region", "FY25", "FY26", "Risk"]
        for column, header in enumerate(headers):
            document.drawString(62 + column * 105, 148, header)
        document.setFont("Helvetica", 9)
        for row, values in enumerate((("North", "4.2", "5.1", "Medium"), ("South", "(1.7)", "2.3", "High"), ("West", "3.8", "4.0", "Low"))):
            for column, value in enumerate(values):
                document.drawString(62 + column * 105, 128 - row * 16, value)
        image = Image.new("RGB", (100, 74), "#D9E2EC")
        draw = ImageDraw.Draw(image)
        draw.rectangle((8, 8, 92, 66), outline="#486581", width=3)
        draw.line((15, 57, 43, 31, 63, 47, 87, 18), fill="#D64545", width=3)
        buffer = io.BytesIO()
        image.save(buffer, format="PNG")
        buffer.seek(0)
        document.drawImage(ImageReader(buffer), 432, 48, width=118, height=88, mask="auto")
    else:
        for index, line in enumerate(lines):
            document.drawString(54, 698 - index * 20, line)
    document.showPage()
    document.save()


def writer_from(source: Path) -> PdfWriter:
    reader = PdfReader(str(source))
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    return writer


def write_writer(writer: PdfWriter, target: Path) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    with target.open("wb") as handle:
        writer.write(handle)


def indirect(writer: PdfWriter, value):
    return writer._add_object(value)


def ensure_annots(page) -> ArrayObject:
    existing = page.get("/Annots")
    if existing is None:
        annots = ArrayObject()
        page[n("Annots")] = annots
        return annots
    return existing.get_object() if hasattr(existing, "get_object") else existing


def create_encrypted_owner_policy(root: Path) -> None:
    temporary = root / ".owner-policy-base.pdf"
    target = root / "pdf" / "encryption" / "owner-policy-aes256.pdf"
    write_reportlab_pdf(
        temporary,
        "Owner-policy boundary fixture",
        [
            "The user password opens this PDF but does not confer owner authority.",
            "Copy/extraction is disabled by the encrypted permission bits.",
            "The embedded attachment and AcroForm signature-field state are canaries.",
        ],
    )
    writer = writer_from(temporary)
    page = writer.pages[0]
    text_field = DictionaryObject({
        n("Type"): n("Annot"),
        n("Subtype"): n("Widget"),
        n("FT"): n("Tx"),
        n("T"): TextStringObject("FixtureStatus"),
        n("V"): TextStringObject("Do not alter encryption"),
        n("Rect"): ArrayObject([FloatObject(54), FloatObject(84), FloatObject(290), FloatObject(106)]),
        n("F"): NumberObject(4),
    })
    signature_field = DictionaryObject({
        n("Type"): n("Annot"),
        n("Subtype"): n("Widget"),
        n("FT"): n("Sig"),
        n("T"): TextStringObject("SignatureStatus"),
        n("Rect"): ArrayObject([FloatObject(318), FloatObject(84), FloatObject(558), FloatObject(106)]),
        n("F"): NumberObject(4),
    })
    text_ref = indirect(writer, text_field)
    signature_ref = indirect(writer, signature_field)
    ensure_annots(page).extend([text_ref, signature_ref])
    writer._root_object[n("AcroForm")] = DictionaryObject({n("Fields"): ArrayObject([text_ref, signature_ref])})
    writer.add_attachment("owner-policy-canary.txt", b"Owner policy fixture attachment. Do not extract without permission.")
    writer.add_metadata({
        "/Title": "Owner-policy boundary fixture",
        "/Subject": "AES-256 owner/user permission split",
        "/Keywords": "OFFICEKIT-OWNER-POLICY-CANARY",
    })
    writer.encrypt(
        USER_PASSWORD,
        OWNER_PASSWORD,
        permissions_flag=UserAccessPermissions.PRINT,
        algorithm="AES-256",
    )
    write_writer(writer, target)
    credentials = root / "pdf" / "encryption" / "user-password.json"
    credentials.write_text(json.dumps({"userPassword": USER_PASSWORD, "scope": "open-only; no owner password is supplied"}, indent=2) + "\n", encoding="utf-8")
    temporary.unlink(missing_ok=True)


def create_annotation_reply_chain(root: Path) -> None:
    temporary = root / ".annotation-replies-base.pdf"
    target = root / "pdf" / "annotations" / "reply-chain.pdf"
    write_reportlab_pdf(
        temporary,
        "Review annotation reply-chain fixture",
        [
            "Reviewer A raised the A-17 source-data issue in this paragraph.",
            "The existing reply chain and resolved state must not be flattened.",
        ],
    )
    writer = writer_from(temporary)
    page = writer.pages[0]
    root_annotation = DictionaryObject({
        n("Type"): n("Annot"),
        n("Subtype"): n("Text"),
        n("Rect"): ArrayObject([FloatObject(62), FloatObject(640), FloatObject(82), FloatObject(660)]),
        n("NM"): TextStringObject("A-17"),
        n("T"): TextStringObject("Reviewer A"),
        n("Contents"): TextStringObject("Please reconcile the source-data total."),
        n("CreationDate"): TextStringObject("D:20260728090000Z"),
        n("StateModel"): n("Review"),
        n("State"): n("Accepted"),
    })
    root_ref = indirect(writer, root_annotation)
    reply = DictionaryObject({
        n("Type"): n("Annot"),
        n("Subtype"): n("Text"),
        n("Rect"): ArrayObject([FloatObject(92), FloatObject(640), FloatObject(112), FloatObject(660)]),
        n("NM"): TextStringObject("A-17-R1"),
        n("T"): TextStringObject("Data steward"),
        n("Contents"): TextStringObject("Existing reply: source extract attached to the review record."),
        n("CreationDate"): TextStringObject("D:20260728100000Z"),
        n("IRT"): root_ref,
        n("RT"): n("R"),
        n("StateModel"): n("Review"),
        n("State"): n("Accepted"),
    })
    reply_ref = indirect(writer, reply)
    popup = DictionaryObject({
        n("Type"): n("Annot"),
        n("Subtype"): n("Popup"),
        n("Rect"): ArrayObject([FloatObject(115), FloatObject(520), FloatObject(310), FloatObject(640)]),
        n("Parent"): root_ref,
        n("Open"): BooleanObject(False),
    })
    highlight = DictionaryObject({
        n("Type"): n("Annot"),
        n("Subtype"): n("Highlight"),
        n("Rect"): ArrayObject([FloatObject(54), FloatObject(675), FloatObject(380), FloatObject(692)]),
        n("QuadPoints"): ArrayObject([FloatObject(54), FloatObject(692), FloatObject(380), FloatObject(692), FloatObject(54), FloatObject(675), FloatObject(380), FloatObject(675)]),
        n("T"): TextStringObject("Reviewer A"),
        n("Contents"): TextStringObject("A-17 highlight anchor"),
    })
    ensure_annots(page).extend([root_ref, reply_ref, indirect(writer, popup), indirect(writer, highlight)])
    write_writer(writer, target)
    temporary.unlink(missing_ok=True)


def create_untagged_complex_report(root: Path) -> None:
    target = root / "pdf" / "accessibility" / "untagged-complex-report.pdf"
    write_reportlab_pdf(
        target,
        "Untyped complex report fixture",
        [
            "Left column: context, methodology, and risk narrative.",
            "Left column: author intent is deliberately not encoded as tags.",
            "Left column: a visual relationship is not a semantic relationship.",
            "Right column: information graphic and multi-level table below.",
            "Right column: table headers span multiple visual columns.",
            "Right column: scanned appendix reference is intentionally ambiguous.",
        ],
        complex_layout=True,
    )


def stream_with_bytes(payload: bytes) -> DecodedStreamObject:
    stream = DecodedStreamObject()
    stream.set_data(payload)
    return stream


def create_dynamic_xfa(root: Path) -> None:
    temporary = root / ".dynamic-xfa-base.pdf"
    target = root / "pdf" / "xfa" / "dynamic-dependents.pdf"
    write_reportlab_pdf(
        temporary,
        "Dynamic XFA dependent-count fixture",
        [
            "This PDF contains an XFA template with a repeating dependent subform.",
            "The FormCalc script changes the occurrence count and pagination.",
        ],
    )
    writer = writer_from(temporary)
    template = stream_with_bytes(
        b"""<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<template xmlns=\"http://www.xfa.org/schema/xfa-template/3.3/\">
    <subform name=\"form1\" layout=\"tb\"><event activity=\"initialize\"><script contentType=\"application/x-javascript\">xfa.form.recalculate(1);</script></event>
    <field name=\"DependentCount\"><ui><numericEdit/></ui><calculate><script contentType=\"application/x-formcalc\">$ = Sum(Dependent[*].count)</script></calculate></field>
    <subform name=\"Dependent\" layout=\"tb\" occur=\"min=0 max=-1\"><field name=\"Name\"><ui><textEdit/></ui></field></subform>
  </subform>
</template>"""
    )
    datasets = stream_with_bytes(
        b"""<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<datasets xmlns=\"http://www.xfa.org/schema/xfa-data/1.0/\"><data><form1><DependentCount>1</DependentCount><Dependent><Name>Fixture dependent</Name></Dependent></form1></data></datasets>"""
    )
    template_ref = indirect(writer, template)
    datasets_ref = indirect(writer, datasets)
    writer._root_object[n("AcroForm")] = DictionaryObject({
        n("Fields"): ArrayObject(),
        n("XFA"): ArrayObject([TextStringObject("template"), template_ref, TextStringObject("datasets"), datasets_ref]),
    })
    writer._root_object[n("NeedsRendering")] = BooleanObject(True)
    write_writer(writer, target)
    temporary.unlink(missing_ok=True)


def dictionary_object(value):
    return value.get_object() if hasattr(value, "get_object") else value


def create_print_production_risk(root: Path) -> None:
    temporary = root / ".print-risk-base.pdf"
    target = root / "pdf" / "print" / "print-production-risk.pdf"
    write_reportlab_pdf(
        temporary,
        "Print-production risk fixture",
        [
            "This page declares DeviceN, a Separation spot colour, overprint, OCG, and an OutputIntent.",
            "An RGB screenshot cannot verify print-production semantics.",
        ],
    )
    writer = writer_from(temporary)
    page = writer.pages[0]
    tint_transform = DictionaryObject({
        n("FunctionType"): NumberObject(2),
        n("Domain"): ArrayObject([NumberObject(0), NumberObject(1), NumberObject(0), NumberObject(1)]),
        n("Range"): ArrayObject([NumberObject(0), NumberObject(1)] * 4),
        n("C0"): ArrayObject([NumberObject(0)] * 4),
        n("C1"): ArrayObject([NumberObject(1)] * 4),
        n("N"): NumberObject(1),
    })
    spot = ArrayObject([n("Separation"), n("PANTONE#20185#20C"), n("DeviceCMYK"), tint_transform])
    devicen = ArrayObject([
        n("DeviceN"),
        ArrayObject([n("PANTONE#20185#20C"), n("PANTONE#20293#20C")]),
        n("DeviceCMYK"),
        tint_transform,
    ])
    resources = dictionary_object(page.get("/Resources"))
    color_space = dictionary_object(resources.get("/ColorSpace")) if resources.get("/ColorSpace") else DictionaryObject()
    color_space[n("SpotRisk")] = spot
    color_space[n("BrandDeviceN")] = devicen
    resources[n("ColorSpace")] = color_space
    ext_state = dictionary_object(resources.get("/ExtGState")) if resources.get("/ExtGState") else DictionaryObject()
    ext_state[n("GSPrint")] = DictionaryObject({n("Type"): n("ExtGState"), n("OP"): BooleanObject(True), n("op"): BooleanObject(True), n("OPM"): NumberObject(1), n("CA"): FloatObject(0.75), n("ca"): FloatObject(0.75), n("BM"): n("Multiply")})
    resources[n("ExtGState")] = ext_state
    ocg = DictionaryObject({n("Type"): n("OCG"), n("Name"): TextStringObject("Print proofing layer")})
    ocg_ref = indirect(writer, ocg)
    properties = dictionary_object(resources.get("/Properties")) if resources.get("/Properties") else DictionaryObject()
    properties[n("Proofing")] = ocg_ref
    resources[n("Properties")] = properties
    overlay = stream_with_bytes(b"q\n/GSPrint gs\n/OC /Proofing BDC\n/BrandDeviceN cs\n0.75 0.25 scn\n54 54 175 24 re f\nEMC\nQ\n")
    overlay_ref = indirect(writer, overlay)
    previous = page.get("/Contents")
    page[n("Contents")] = ArrayObject([previous, overlay_ref]) if previous is not None else ArrayObject([overlay_ref])
    profile = stream_with_bytes(b"OfficeKit PromptBench synthetic CMYK output profile placeholder")
    profile[n("N")] = NumberObject(4)
    profile_ref = indirect(writer, profile)
    output_intent = DictionaryObject({
        n("Type"): n("OutputIntent"),
        n("S"): n("GTS_PDFX"),
        n("OutputConditionIdentifier"): TextStringObject("OfficeKit synthetic print condition"),
        n("Info"): TextStringObject("Structural print-production risk fixture; not a certified PDF/X profile."),
        n("DestOutputProfile"): profile_ref,
    })
    writer._root_object[n("OutputIntents")] = ArrayObject([indirect(writer, output_intent)])
    writer._root_object[n("OCProperties")] = DictionaryObject({
        n("OCGs"): ArrayObject([ocg_ref]),
        n("D"): DictionaryObject({n("Order"): ArrayObject([ocg_ref]), n("ON"): ArrayObject([ocg_ref])}),
    })
    write_writer(writer, target)
    temporary.unlink(missing_ok=True)


def required_signing_python(value: str | None) -> str:
    candidate = value or os.environ.get(SIGNING_PYTHON_ENV)
    if not candidate:
        raise ValueError(
            "Generating the real DocMDP fixture requires --signing-python or "
            f"{SIGNING_PYTHON_ENV}; select the policy-authorized managed pyHanko runtime."
        )
    executable = Path(candidate)
    if not executable.is_file() or executable.is_symlink() or not os.access(executable, os.X_OK):
        raise ValueError("the selected PromptBench signing Python must be a regular executable file")
    return str(executable)


def create_docmdp_p1_final(root: Path, signing_python: str) -> None:
    """Create a real P=1 certification without retaining a private key.

    The committed corpus receives exactly one signed PDF and the public test
    root.  RSA private material, the temporary signer certificate, and the
    unsigned staging PDF live only under `TemporaryDirectory` and are removed
    when this function returns.
    """
    target = root / "pdf" / "signing" / "docmdp-p1-final.pdf"
    root_certificate = root / "pdf" / "signing" / "test-pki" / "root.pem"
    with tempfile.TemporaryDirectory(prefix="officekit-promptbench-docmdp-") as directory:
        work = Path(directory)
        source = work / "unsigned-final.pdf"
        write_reportlab_pdf(
            source,
            "Final",
            [
                "This self-authored document is certified with DocMDP P=1.",
                "Any title change must be refused without an appended revision.",
            ],
        )
        completed = subprocess.run(
            [signing_python, "-I", "-c", DOCMDP_P1_SIGNING_PROGRAM, str(work), str(source), str(target), str(root_certificate)],
            check=False,
            capture_output=True,
            text=True,
            timeout=60,
        )
        if completed.returncode != 0:
            target.unlink(missing_ok=True)
            root_certificate.unlink(missing_ok=True)
            raise ValueError(
                "managed pyHanko fixture signing failed: "
                f"{completed.stderr.strip() or completed.stdout.strip() or 'unknown error'}"
            )


def refresh_docmdp_p1(root: Path, signing_python: str) -> dict:
    """Refresh only the signed DocMDP fixture and its two integrity records."""
    manifest_path = root / "integrity.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    if manifest.get("schemaVersion") != 1 or not isinstance(manifest.get("assets"), dict):
        raise ValueError("unsupported corpus integrity schema")
    create_docmdp_p1_final(root, signing_python)
    for relative in ("pdf/signing/docmdp-p1-final.pdf", "pdf/signing/test-pki/root.pem"):
        asset = root / relative
        manifest["assets"][relative] = {
            "bytes": asset.stat().st_size,
            "description": FIXTURES[relative],
            "kind": "file",
            "sha256": sha256(asset),
        }
    manifest_path.write_text(json.dumps(manifest, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    return manifest


FIXTURES = {
    "pdf/encryption/owner-policy-aes256.pdf": "AES-256 encrypted user/owner permission split with embedded attachment and AcroForm canaries.",
    "pdf/encryption/user-password.json": "Test-only user credential; deliberately excludes the owner password.",
    "pdf/annotations/reply-chain.pdf": "Native PDF annotation root/reply/Popup/Highlight graph with resolved-state semantics.",
    "pdf/accessibility/untagged-complex-report.pdf": "Untagged two-column report with image and table visual structure.",
    "pdf/xfa/dynamic-dependents.pdf": "Dynamic-XFA-shaped template/datasets packet with repeat and FormCalc markers.",
    "pdf/print/print-production-risk.pdf": "Structural DeviceN/Separation/overprint/OCG/OutputIntent print-risk fixture.",
    "pdf/signing/docmdp-p1-final.pdf": "Real self-authored certification signature with DocMDP P=1 and a Final metadata canary.",
    "pdf/signing/test-pki/root.pem": "Public-only self-authored PromptBench root certificate for the DocMDP P=1 fixture.",
}


def generate(root: Path, signing_python: str | None) -> dict:
    root.mkdir(parents=True, exist_ok=True)
    create_encrypted_owner_policy(root)
    create_annotation_reply_chain(root)
    create_untagged_complex_report(root)
    create_dynamic_xfa(root)
    create_print_production_risk(root)
    create_docmdp_p1_final(root, required_signing_python(signing_python))
    assets = {}
    for relative, description in FIXTURES.items():
        path = root / relative
        assets[relative] = {"sha256": sha256(path), "bytes": path.stat().st_size, "description": description, "kind": "file"}
    manifest = {
        "schemaVersion": 1,
        "scope": "repository-only PromptBench assets; never npm runtime inputs",
        "provenance": "self-authored OfficeKit test fixtures generated by scripts/agent-eval-corpus-fixtures.py",
        "assets": assets,
    }
    (root / "integrity.json").write_text(json.dumps(manifest, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    return manifest


def root_dictionary(reader: PdfReader):
    return reader.trailer["/Root"].get_object()


def verify_encryption(path: Path) -> None:
    reader = PdfReader(str(path))
    if not reader.is_encrypted:
        raise ValueError("expected encrypted PDF")
    if not reader.decrypt(USER_PASSWORD):
        raise ValueError("fixture user password cannot open encrypted PDF")
    encryption = reader.trailer["/Encrypt"].get_object()
    if int(encryption.get("/R", 0)) < 5 or int(encryption.get("/V", 0)) < 5:
        raise ValueError("expected AES-256 encryption dictionary")
    permissions = int(encryption.get("/P", 0))
    if permissions & (1 << 4):
        raise ValueError("copy/extract permission unexpectedly enabled")
    root = root_dictionary(reader)
    if "/AcroForm" not in root or "/Names" not in root:
        raise ValueError("encrypted fixture is missing AcroForm or attachment names")


def verify_annotations(path: Path) -> None:
    reader = PdfReader(str(path))
    annots = [reference.get_object() for reference in reader.pages[0]["/Annots"]]
    if not any(str(annotation.get("/Subtype")) == "/Highlight" for annotation in annots):
        raise ValueError("annotation fixture has no Highlight")
    roots = [annotation for annotation in annots if str(annotation.get("/NM", "")) == "A-17"]
    replies = [annotation for annotation in annots if "/IRT" in annotation]
    popups = [annotation for annotation in annots if str(annotation.get("/Subtype")) == "/Popup"]
    if len(roots) != 1 or len(replies) != 1 or len(popups) != 1:
        raise ValueError("annotation reply graph is incomplete")
    if str(replies[0].get("/State")) != "/Accepted" or str(replies[0].get("/StateModel")) != "/Review":
        raise ValueError("annotation reply has no resolved-state semantics")


def verify_untagged_report(path: Path) -> None:
    reader = PdfReader(str(path))
    root = root_dictionary(reader)
    if "/StructTreeRoot" in root:
        raise ValueError("complex accessibility fixture unexpectedly has a structure tree")
    page = reader.pages[0]
    xobjects = dictionary_object(page["/Resources"]).get("/XObject")
    if not xobjects:
        raise ValueError("complex accessibility fixture has no image XObject")
    text = page.extract_text() or ""
    if "author intent" not in text or "South" not in text:
        raise ValueError("complex report canaries are missing")


def verify_xfa(path: Path) -> None:
    reader = PdfReader(str(path))
    root = root_dictionary(reader)
    acroform = dictionary_object(root.get("/AcroForm"))
    xfa = acroform.get("/XFA") if acroform else None
    if not isinstance(xfa, ArrayObject) or len(xfa) < 4 or not bool(root.get("/NeedsRendering")):
        raise ValueError("XFA fixture has no dynamic XFA packet")
    streams = b"".join(
        resolved.get_data()
        for item in xfa
        for resolved in [item.get_object() if hasattr(item, "get_object") else item]
        if hasattr(resolved, "get_data")
    )
    if b"occur=\"min=0 max=-1\"" not in streams or b"application/x-formcalc" not in streams or b"application/x-javascript" not in streams:
        raise ValueError("XFA dynamic/recalculation canaries are missing")


def verify_print(path: Path) -> None:
    reader = PdfReader(str(path))
    root = root_dictionary(reader)
    if not root.get("/OutputIntents") or not root.get("/OCProperties"):
        raise ValueError("print fixture has no OutputIntent/OCProperties")
    resources = dictionary_object(reader.pages[0]["/Resources"])
    colors = dictionary_object(resources.get("/ColorSpace"))
    ext = dictionary_object(resources.get("/ExtGState"))
    if "/BrandDeviceN" not in colors or "/SpotRisk" not in colors or "/GSPrint" not in ext:
        raise ValueError("print fixture has no DeviceN/Separation/overprint resource")
    if not bool(ext["/GSPrint"].get_object().get("/OP")):
        raise ValueError("print fixture has no overprint flag")


def verify_docmdp_p1(path: Path, root_certificate: Path) -> None:
    raw = path.read_bytes()
    reader = PdfReader(str(path), strict=True)
    root = root_dictionary(reader)
    permissions = dictionary_object(root.get("/Perms")) if root.get("/Perms") else {}
    signature = dictionary_object(permissions.get("/DocMDP")) if permissions.get("/DocMDP") else {}
    byte_range = signature.get("/ByteRange") if signature else None
    if not isinstance(byte_range, ArrayObject) or len(byte_range) != 4:
        raise ValueError("DocMDP fixture is missing one four-number ByteRange")
    offsets = [int(value) for value in byte_range]
    if offsets[0] != 0 or min(offsets[1:]) < 0 or offsets[0] + offsets[1] > offsets[2] or offsets[2] + offsets[3] != len(raw):
        raise ValueError("DocMDP fixture has an invalid ByteRange")
    if not signature.get("/Contents"):
        raise ValueError("DocMDP fixture has no CMS contents")
    references = signature.get("/Reference") if signature else None
    if not isinstance(references, ArrayObject):
        raise ValueError("DocMDP fixture has no transform reference")
    transform_params = [
        dictionary_object(reference).get("/TransformParams")
        for reference in references
        if str(dictionary_object(reference).get("/TransformMethod", "")) == "/DocMDP"
    ]
    if len(transform_params) != 1:
        raise ValueError("DocMDP fixture has no unique DocMDP transform")
    params = dictionary_object(transform_params[0])
    if int(params.get("/P", 0) or 0) != 1:
        raise ValueError("DocMDP fixture does not prohibit all later changes")
    if str(reader.metadata.title or "") != "Final":
        raise ValueError("DocMDP fixture title canary is missing")
    if "certified with DocMDP P=1" not in (reader.pages[0].extract_text() or ""):
        raise ValueError("DocMDP fixture visible-content canary is missing")
    certificate = root_certificate.read_bytes()
    if b"BEGIN CERTIFICATE" not in certificate or b"PRIVATE KEY" in certificate:
        raise ValueError("DocMDP fixture root certificate must be public-only PEM")


def verify(root: Path) -> dict:
    manifest_path = root / "integrity.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    if manifest.get("schemaVersion") != 1:
        raise ValueError("unsupported corpus integrity schema")
    for relative, expected in manifest.get("assets", {}).items():
        path = root / relative
        if not path.is_file() or path.is_symlink():
            raise ValueError(f"locked fixture is missing or unsafe: {relative}")
        actual_hash = sha256(path)
        if actual_hash != expected.get("sha256") or path.stat().st_size != expected.get("bytes"):
            raise ValueError(f"locked fixture integrity mismatch: {relative}")
    verify_encryption(root / "pdf" / "encryption" / "owner-policy-aes256.pdf")
    verify_annotations(root / "pdf" / "annotations" / "reply-chain.pdf")
    verify_untagged_report(root / "pdf" / "accessibility" / "untagged-complex-report.pdf")
    verify_xfa(root / "pdf" / "xfa" / "dynamic-dependents.pdf")
    verify_print(root / "pdf" / "print" / "print-production-risk.pdf")
    verify_docmdp_p1(root / "pdf" / "signing" / "docmdp-p1-final.pdf", root / "pdf" / "signing" / "test-pki" / "root.pem")
    return {"ok": True, "assets": len(manifest["assets"]), "root": str(root)}


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("command", choices=("generate", "refresh-docmdp", "verify"))
    parser.add_argument("--root", type=Path, default=DEFAULT_ROOT)
    parser.add_argument("--signing-python", help=f"managed pyHanko interpreter used only by generate (or set {SIGNING_PYTHON_ENV})")
    options = parser.parse_args()
    if options.command == "generate":
        print(json.dumps(generate(options.root, options.signing_python), indent=2, sort_keys=True))
    elif options.command == "refresh-docmdp":
        print(json.dumps(refresh_docmdp_p1(options.root, required_signing_python(options.signing_python)), indent=2, sort_keys=True))
    else:
        print(json.dumps(verify(options.root), sort_keys=True))


if __name__ == "__main__":
    try:
        main()
    except Exception as error:  # Explicit diagnostics are part of the corpus contract.
        print(f"agent-eval corpus fixture error: {error}", file=sys.stderr)
        raise SystemExit(2)
