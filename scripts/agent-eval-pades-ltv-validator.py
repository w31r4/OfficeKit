#!/usr/bin/env python3
"""Evaluator-only cryptographic evidence for the PromptBench PAdES-LTA case.

The Agent never receives this script.  It imports pyHanko directly rather than
calling OfficeKit's provider adapter, so a candidate cannot pass merely by
printing a provider-shaped report.  The profile is deliberately narrow: one
disclosed test root, one CRL, one approval signature and one document timestamp.
It is evidence for the repository's offline test profile, not a PAdES
conformance certificate.
"""

from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
import stat
import sys
from typing import Any


MAX_PDF_BYTES = 512 * 1024 * 1024
MAX_CERTIFICATE_BYTES = 4 * 1024 * 1024
MAX_CRL_BYTES = 4 * 1024 * 1024
MAX_SIGNATURES = 16


class EvaluationError(RuntimeError):
    pass


def regular_file(value: Any, label: str, maximum: int, *, minimum: int = 1) -> Path:
    if not isinstance(value, str) or not value:
        raise EvaluationError(f"{label} must be a non-empty path")
    candidate = Path(value).expanduser()
    candidate = candidate if candidate.is_absolute() else Path.cwd() / candidate
    candidate = Path(os.path.abspath(candidate))
    try:
        metadata = candidate.lstat()
    except FileNotFoundError as exc:
        raise EvaluationError(f"{label} does not exist: {candidate}") from exc
    if stat.S_ISLNK(metadata.st_mode):
        raise EvaluationError(f"{label} is a symbolic link and will not be followed")
    if not stat.S_ISREG(metadata.st_mode):
        raise EvaluationError(f"{label} is not a regular file")
    if not minimum <= metadata.st_size <= maximum:
        raise EvaluationError(f"{label} size {metadata.st_size} is outside {minimum}..{maximum} bytes")
    return candidate


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def dereference(value: Any) -> Any:
    return value.get_object() if hasattr(value, "get_object") else value


def count_dss_entries(dss: Any, key: str) -> int:
    value = dereference(dss.get(key)) if hasattr(dss, "get") and dss.get(key) is not None else []
    try:
        count = len(value)
    except TypeError as exc:
        raise EvaluationError(f"DSS {key} is not an array") from exc
    if count < 0 or count > MAX_SIGNATURES * MAX_SIGNATURES:
        raise EvaluationError(f"DSS {key} count exceeds the evaluator budget")
    return count


def dss_evidence(root: Any) -> dict[str, Any]:
    dss = dereference(root.get("/DSS")) if hasattr(root, "get") and root.get("/DSS") is not None else None
    if not hasattr(dss, "get"):
        return {"present": False, "certificateCount": 0, "crlCount": 0, "ocspCount": 0, "vriCount": 0}
    return {
        "present": True,
        "certificateCount": count_dss_entries(dss, "/Certs"),
        "crlCount": count_dss_entries(dss, "/CRLs"),
        "ocspCount": count_dss_entries(dss, "/OCSPs"),
        "vriCount": count_dss_entries(dss, "/VRI"),
    }


def byte_range(signature: Any) -> dict[str, Any]:
    raw = signature.sig_object.get("/ByteRange") if hasattr(signature, "sig_object") else None
    try:
        offsets = [int(value) for value in raw] if raw is not None and len(raw) == 4 else []
    except (TypeError, ValueError):
        offsets = []
    valid = bool(
        len(offsets) == 4
        and offsets[0] == 0
        and min(offsets[1:]) >= 0
        and offsets[0] + offsets[1] <= offsets[2]
    )
    return {
        "offsets": offsets,
        "validSegments": valid,
        "coveredBytes": offsets[2] + offsets[3] if valid else None,
    }


def timestamp_evidence(status: Any) -> dict[str, Any] | None:
    if status is None:
        return None
    intact = bool(getattr(status, "intact", False))
    valid = bool(getattr(status, "valid", False))
    trusted = bool(getattr(status, "trusted", False))
    return {
        "intact": intact,
        "cryptographicallyValid": valid,
        "trusted": trusted,
        # TimestampSignatureStatus does not expose the PDF signature
        # difference-analysis bottom-line.  Never fabricate that claim.
        "bottomLine": bool(getattr(status, "bottom_line", intact and valid and trusted)),
    }


def validate_signature(signature: Any, index: int, validation_context: Any) -> dict[str, Any]:
    from pyhanko.sign.validation import validate_pdf_signature, validate_pdf_timestamp

    object_type = str(signature.sig_object_type)
    document_timestamp = object_type == "/DocTimeStamp"
    record: dict[str, Any] = {
        "index": index,
        "fieldName": str(signature.field_name),
        "signatureObjectType": object_type,
        "documentTimestamp": document_timestamp,
        "subFilter": str(signature.sig_object.get("/SubFilter", "")),
        "byteRange": byte_range(signature),
    }
    try:
        status = (
            validate_pdf_timestamp(signature, validation_context=validation_context)
            if document_timestamp
            else validate_pdf_signature(
                signature,
                signer_validation_context=validation_context,
                ts_validation_context=validation_context,
            )
        )
        intact = bool(getattr(status, "intact", False))
        valid = bool(getattr(status, "valid", False))
        trusted = bool(getattr(status, "trusted", False))
        record.update({
            "validationCompleted": True,
            "intact": intact,
            "cryptographicallyValid": valid,
            "trusted": trusted,
            "bottomLine": bool(getattr(status, "bottom_line", intact and valid and trusted)),
            "signatureTimestamp": None if document_timestamp else timestamp_evidence(getattr(status, "timestamp_validity", None)),
            "docMDPCompliant": None if document_timestamp else bool(getattr(status, "docmdp_ok", False)),
            "coverage": getattr(getattr(status, "coverage", None), "name", None),
            "modificationLevel": getattr(getattr(status, "modification_level", None), "name", None),
        })
    except Exception as exc:  # malformed CMS is evidence, not an evaluator crash
        record.update({"validationCompleted": False, "validationError": f"{type(exc).__name__}: {exc}"[:2048]})
    return record


def evaluate(payload: dict[str, Any]) -> dict[str, Any]:
    from asn1crypto import crl as asn1_crl
    from pyhanko.keys import load_cert_from_pemder
    from pyhanko.pdf_utils.reader import PdfFileReader
    from pyhanko_certvalidator import ValidationContext

    output = regular_file(payload.get("output"), "output PDF", MAX_PDF_BYTES, minimum=5)
    root = regular_file(payload.get("root"), "test trust root", MAX_CERTIFICATE_BYTES)
    crl = regular_file(payload.get("crl"), "test CRL", MAX_CRL_BYTES)
    expected_output_hash = payload.get("expectedOutputSha256")
    output_hash = sha256(output)
    if expected_output_hash is not None:
        if not isinstance(expected_output_hash, str) or expected_output_hash.lower() != output_hash:
            raise EvaluationError("output PDF SHA-256 does not match the evaluator input")
    with output.open("rb") as stream:
        if stream.read(5) != b"%PDF-":
            raise EvaluationError("output PDF does not begin with a PDF header")

    root_cert = load_cert_from_pemder(str(root))
    crl_value = asn1_crl.CertificateList.load(crl.read_bytes())

    def context() -> Any:
        return ValidationContext(
            trust_roots=[root_cert],
            crls=[crl_value],
            allow_fetching=False,
            revocation_mode="require",
        )

    with output.open("rb") as stream:
        reader = PdfFileReader(stream, strict=True)
        if reader.security_handler is not None:
            raise EvaluationError("encrypted PDFs are unsupported by this bounded evaluator")
        signatures = list(reader.embedded_signatures)
        if len(signatures) > MAX_SIGNATURES:
            raise EvaluationError("PDF has too many embedded signatures for this evaluator")
        records = [validate_signature(signature, index, context()) for index, signature in enumerate(signatures)]
        dss = dss_evidence(reader.root)
        revisions = int(reader.xrefs.total_revisions)

    approval = [record for record in records if not record["documentTimestamp"]]
    document_timestamps = [record for record in records if record["documentTimestamp"]]
    approval_timestamp_ok = bool(approval) and all(
        isinstance(record.get("signatureTimestamp"), dict)
        and bool(record["signatureTimestamp"].get("bottomLine"))
        for record in approval
    )
    ordinary_ok = len(approval) == 1 and all(
        record.get("validationCompleted") and record.get("intact") and record.get("cryptographicallyValid")
        and record.get("trusted") and record.get("bottomLine") and record.get("docMDPCompliant")
        and record.get("subFilter") == "/ETSI.CAdES.detached"
        for record in approval
    )
    document_timestamp_ok = len(document_timestamps) == 1 and all(
        record.get("validationCompleted") and record.get("intact") and record.get("cryptographicallyValid")
        and record.get("trusted") and record.get("bottomLine") and record.get("subFilter") == "/ETSI.RFC3161"
        for record in document_timestamps
    )
    ltv_evidence_ok = dss["present"] and dss["certificateCount"] >= 2 and dss["crlCount"] >= 1 and dss["vriCount"] >= 2
    return {
        "schema": "office-kit.promptbench-pades-ltv-validator.v1",
        "ok": ordinary_ok and approval_timestamp_ok and document_timestamp_ok and ltv_evidence_ok,
        "profile": "bounded-offline-test-pades-lta-evidence",
        "padesProfileConformanceClaimed": False,
        "networkAllowed": False,
        "output": {"sha256": output_hash, "bytes": output.stat().st_size},
        "trust": {"rootSha256": sha256(root), "crlSha256": sha256(crl), "revocationMode": "require"},
        "revisionCount": revisions,
        "signatures": records,
        "summary": {
            "signatureCount": len(records),
            "ordinarySignatureCount": len(approval),
            "documentTimestampCount": len(document_timestamps),
            "signatureTimestampCount": sum(record.get("signatureTimestamp") is not None for record in approval),
            "ordinarySignatureValid": ordinary_ok,
            "signatureTimestampsValid": approval_timestamp_ok,
            "documentTimestampValid": document_timestamp_ok,
            "dssValidationInfoValid": ltv_evidence_ok,
        },
        "dss": dss,
    }


def main() -> int:
    try:
        payload = json.load(sys.stdin)
        if not isinstance(payload, dict):
            raise EvaluationError("evaluator input must be one JSON object")
        print(json.dumps(evaluate(payload), ensure_ascii=False, separators=(",", ":")))
        return 0
    except Exception as exc:
        print(f"{type(exc).__name__}: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
