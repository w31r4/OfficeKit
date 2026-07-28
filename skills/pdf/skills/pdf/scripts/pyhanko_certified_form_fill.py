#!/usr/bin/env python3
"""Finalize one explicitly authorised DocMDP P=2 text field.

This is deliberately not a general signed-PDF editing interface.  It accepts
one flat AcroForm source with one trusted certification signature, DocMDP
``P=2`` and an explicit FieldMDP include list.  It turns exactly one empty,
unlocked text field into a visible, read-only decimal value in one incremental
revision, then proves the resulting policy evidence before publishing a
distinct output without replacement.
"""

from __future__ import annotations

import argparse
import errno
import json
import logging
import os
from pathlib import Path
import re
import shutil
import stat
import subprocess
import sys
import tempfile
import threading
from typing import Any

from pyhanko_provider import ProviderError as ValidationProviderError
from pyhanko_provider import bounded_text, provider_versions, sha256, verify
from pyhanko_sign_provider import ProviderError as SigningProviderError
from pyhanko_sign_provider import (
    destination_path,
    has_exact_prefix,
    publish_without_replace,
    regular_input_path,
    snapshot_file,
)


SCHEMA = "office-kit.pyhanko-certified-form-fill.v1"
DEFAULT_MAX_INPUT_BYTES = 512 * 1024 * 1024
DEFAULT_MAX_OUTPUT_BYTES = 1024 * 1024 * 1024
DEFAULT_MAX_CERTIFICATE_BYTES = 4 * 1024 * 1024
DEFAULT_MAX_PAGES = 10_000
DEFAULT_MAX_FIELDS = 128
DEFAULT_TIMEOUT_SECONDS = 120
DEFAULT_MAX_STDOUT_BYTES = 2 * 1024 * 1024
DEFAULT_MAX_STDERR_BYTES = 512 * 1024
MAX_WORKER_CONFIG_BYTES = 128 * 1024
MAX_FIELD_NAME_CHARS = 128
MAX_FIELD_VALUE_CHARS = 64
DECIMAL_VALUE = re.compile(r"(?:0|[1-9][0-9]{0,8})\.[0-9]{2}\Z")


class ProviderError(RuntimeError):
    pass


def positive_int(value: str) -> int:
    parsed = int(value)
    if parsed <= 0:
        raise argparse.ArgumentTypeError("must be a positive integer")
    return parsed


def expected_hash(value: str) -> str:
    normalized = value.strip().lower()
    if not re.fullmatch(r"[0-9a-f]{64}", normalized):
        raise ProviderError("--expected-source-sha256 must contain exactly 64 hexadecimal characters")
    return normalized


def acroform_name(value: str, option: str) -> str:
    if not (1 <= len(value) <= MAX_FIELD_NAME_CHARS):
        raise ProviderError(f"{option} must contain 1..{MAX_FIELD_NAME_CHARS} characters")
    if not re.fullmatch(r"[A-Za-z0-9_.-]+", value):
        raise ProviderError(f"{option} may contain only ASCII letters, digits, underscore, hyphen, and hierarchy dots")
    if value.startswith(".") or value.endswith(".") or ".." in value:
        raise ProviderError(f"{option} contains an empty hierarchy segment")
    return value


def decimal_value(value: str) -> str:
    if not DECIMAL_VALUE.fullmatch(value):
        raise ProviderError(
            "--value must be a canonical non-negative decimal with exactly two fraction digits "
            "and at most nine integral digits"
        )
    if len(value) > MAX_FIELD_VALUE_CHARS:
        raise ProviderError(f"--value cannot exceed {MAX_FIELD_VALUE_CHARS} characters")
    return value


def limited_text(value: str, option: str, limit: int = 256) -> str:
    if not value or len(value) > limit:
        raise ProviderError(f"{option} must contain 1..{limit} characters")
    if any(ord(character) < 0x20 for character in value):
        raise ProviderError(f"{option} contains unsupported control characters")
    return value


def require_input_trust(args: argparse.Namespace) -> str:
    if bool(args.trusted_input) == bool(args.caller_isolated):
        raise ProviderError("select exactly one of --trusted-input or --caller-isolated")
    return "trusted-input" if args.trusted_input else "caller-isolated"


def safe_destination(value: str, source: Path) -> Path:
    target = destination_path(value, source)
    try:
        parent = target.parent.lstat()
    except FileNotFoundError as exc:
        raise ProviderError(f"output parent directory does not exist: {target.parent}") from exc
    if stat.S_ISLNK(parent.st_mode):
        raise ProviderError(f"output parent directory is a symbolic link and will not be followed: {target.parent}")
    if not stat.S_ISDIR(parent.st_mode):
        raise ProviderError(f"output parent is not a directory: {target.parent}")
    return target


def object_reference(value: Any) -> str | None:
    reference = getattr(value, "reference", None) or getattr(value, "indirect_reference", None)
    if reference is None:
        return None
    try:
        return f"{int(reference.idnum)} {int(reference.generation)} R"
    except (AttributeError, TypeError, ValueError):
        return None


def resolve(value: Any) -> Any:
    getter = getattr(value, "get_object", None)
    if callable(getter):
        return getter()
    return value


def pdf_text(value: Any) -> str | None:
    value = resolve(value)
    if value is None:
        return None
    text = str(value)
    return text if len(text) <= 1_024 else text[:1_024] + "…"


def pdf_name(value: Any) -> str | None:
    text = pdf_text(value)
    if text is None:
        return None
    return text[1:] if text.startswith("/") else text


def integer(value: Any, default: int = 0) -> int:
    value = resolve(value)
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def normal_appearance(field: Any) -> bool:
    appearance = resolve(field.get("/AP")) if isinstance(field, dict) and field.get("/AP") else None
    return isinstance(appearance, dict) and appearance.get("/N") is not None


def flat_form_fields(reader: Any, max_fields: int) -> dict[str, dict[str, Any]]:
    root = resolve(reader.root)
    acroform = resolve(root.get("/AcroForm")) if isinstance(root, dict) and root.get("/AcroForm") else None
    if not isinstance(acroform, dict):
        raise ProviderError("source has no flat AcroForm field tree")
    roots = resolve(acroform.get("/Fields"))
    if not isinstance(roots, (list, tuple)) or not roots:
        raise ProviderError("source has no AcroForm root fields")
    records: dict[str, dict[str, Any]] = {}
    for reference in roots:
        field = resolve(reference)
        if not isinstance(field, dict):
            raise ProviderError("AcroForm contains a non-dictionary root field")
        if field.get("/Kids"):
            raise ProviderError("nested or hierarchical AcroForm fields are outside this controlled transaction")
        name = pdf_text(field.get("/T"))
        field_type = pdf_name(field.get("/FT"))
        if not name or not field_type:
            raise ProviderError("every controlled AcroForm root field must have /T and /FT")
        if name in records:
            raise ProviderError(f"AcroForm repeats field name {name!r}")
        if len(records) >= max_fields:
            raise ProviderError(f"AcroForm contains more than {max_fields} fields")
        flags = integer(field.get("/Ff"))
        records[name] = {
            "name": name,
            "fieldType": field_type,
            "value": pdf_text(field.get("/V")),
            "readOnly": bool(flags & 1),
            "flags": flags,
            "widget": pdf_name(field.get("/Subtype")) == "Widget",
            "hasNormalAppearance": normal_appearance(field),
            "hasDefaultAppearance": field.get("/DA") is not None,
            "reference": object_reference(reference) or object_reference(field),
        }
    return records


def document_snapshot(reader: Any, max_fields: int, max_pages: int) -> dict[str, Any]:
    root = resolve(reader.root)
    pages = resolve(root.get("/Pages")) if isinstance(root, dict) else None
    if not isinstance(pages, dict):
        raise ProviderError("PDF catalog has no readable page tree")
    page_count = integer(pages.get("/Count"))
    if page_count < 1 or page_count > max_pages:
        raise ProviderError(f"PDF page count {page_count} is outside the 1..{max_pages} page budget")
    trailer = reader.trailer
    return {
        "pageCount": page_count,
        "fields": flat_form_fields(reader, max_fields),
        "catalog": {
            "perms": object_reference(root.get("/Perms")) if isinstance(root, dict) else None,
            "metadata": object_reference(root.get("/Metadata")) if isinstance(root, dict) else None,
            "acroForm": object_reference(root.get("/AcroForm")) if isinstance(root, dict) else None,
        },
        "info": object_reference(trailer.get("/Info")) if isinstance(trailer, dict) else None,
    }


def validate_form_surface(snapshot: dict[str, Any], args: argparse.Namespace, *, after: bool = False) -> dict[str, Any]:
    fields = snapshot["fields"]
    target = fields.get(args.field)
    locked = fields.get(args.expected_locked_field)
    if target is None:
        raise ProviderError(f"target field {args.field!r} is absent")
    if locked is None:
        raise ProviderError(f"expected locked field {args.expected_locked_field!r} is absent")
    if target["fieldType"] != "Tx" or not target["widget"]:
        raise ProviderError("target must be one flat visible AcroForm /Tx widget field")
    if locked["fieldType"] != "Tx" or not locked["readOnly"]:
        raise ProviderError("expected locked field must be one read-only flat AcroForm /Tx field")
    if locked["value"] != args.expected_locked_value:
        raise ProviderError("expected locked field value does not match the source-bound precondition")
    if after:
        if target["value"] != args.value:
            raise ProviderError("output target field value does not match the requested decimal")
        if not target["readOnly"]:
            raise ProviderError("output target field is not static/read-only")
        if not target["hasNormalAppearance"]:
            raise ProviderError("output target field has no visible normal appearance")
        if target["hasDefaultAppearance"]:
            raise ProviderError("output target field retained a dynamic default appearance")
    else:
        if target["value"] not in {None, ""}:
            raise ProviderError("target field must be empty before this controlled finalisation")
        if target["readOnly"]:
            raise ProviderError("target field is already read-only")
    return {"target": target, "locked": locked}


def unchanged_non_target_fields(before: dict[str, Any], after: dict[str, Any], target_name: str) -> bool:
    if set(before) != set(after):
        return False
    return all(before[name] == after[name] for name in before if name != target_name)


def validator_args(candidate: Path, digest: str, trust_root: Path, args: argparse.Namespace) -> argparse.Namespace:
    return argparse.Namespace(
        input=str(candidate),
        expected_sha256=digest,
        trust_policy="explicit-roots",
        trust_root=[str(trust_root)],
        other_cert=[],
        moment=None,
        revocation_policy="none",
        require_signature=True,
        require_all_integrity_valid=True,
        require_all_trusted=True,
        require_docmdp_compliant=True,
        require_all_bottom_line=True,
        max_input_bytes=args.max_input_bytes,
        timeout_seconds=args.timeout_seconds,
        max_stdout_bytes=min(args.max_stdout_bytes, DEFAULT_MAX_STDOUT_BYTES),
        max_stderr_bytes=min(args.max_stderr_bytes, DEFAULT_MAX_STDERR_BYTES),
    )


def validate_signature(candidate: Path, digest: str, trust_root: Path, args: argparse.Namespace, phase: str) -> dict[str, Any]:
    report = verify(validator_args(candidate, digest, trust_root, args))
    if not report.get("ok"):
        failures = report.get("policyGates", {}).get("failures", [])
        detail = failures[0] if failures else report.get("conclusion", "unknown validation failure")
        raise ProviderError(f"{phase} explicit-root signature validation failed: {bounded_text(detail)}")
    return report


def certification_record(report: dict[str, Any], args: argparse.Namespace, *, phase: str) -> dict[str, Any]:
    signatures = report.get("signatures", [])
    if report.get("summary", {}).get("signatureCount") != 1 or len(signatures) != 1:
        raise ProviderError(f"{phase} requires exactly one certification signature")
    signature = signatures[0]
    if signature.get("fieldName") != args.expected_signature_field:
        raise ProviderError(f"{phase} certification field does not match --expected-signature-field")
    docmdp = signature.get("docMDP") or {}
    fieldmdp = signature.get("fieldMDP") or {}
    if docmdp.get("permission") != "fill-forms":
        raise ProviderError(f"{phase} requires DocMDP P=2/fill-forms")
    if fieldmdp.get("action") != "include" or fieldmdp.get("fields") != [args.expected_locked_field]:
        raise ProviderError(f"{phase} requires FieldMDP Include for exactly --expected-locked-field")
    if not (signature.get("intact") and signature.get("cryptographicallyValid") and signature.get("trusted") and signature.get("bottomLine")):
        raise ProviderError(f"{phase} certification signature did not satisfy integrity/trust/bottom-line gates")
    if not signature.get("docMDPCompliant"):
        raise ProviderError(f"{phase} certification signature is not DocMDP compliant")
    return signature


def preflight_signature(report: dict[str, Any], args: argparse.Namespace) -> dict[str, Any]:
    signature = certification_record(report, args, phase="preflight")
    if signature.get("coverage") != "entire-file":
        raise ProviderError("preflight certification signature must cover the entire baseline file")
    if signature.get("modificationLevel") != "none" or report.get("summary", {}).get("hasPostSigningChanges"):
        raise ProviderError("preflight source already contains post-certification revisions")
    return signature


def postflight_signature(report: dict[str, Any], args: argparse.Namespace) -> dict[str, Any]:
    signature = certification_record(report, args, phase="postflight")
    if signature.get("coverage") != "entire-revision":
        raise ProviderError("output certification signature does not preserve the signed baseline revision")
    if signature.get("modificationLevel") != "form-filling":
        raise ProviderError("output is not recognised as a DocMDP form-filling revision")
    if signature.get("changedFormFields") != [args.field]:
        raise ProviderError("output change set is not exactly the requested target field")
    return signature


def fill_worker(config: dict[str, Any]) -> dict[str, Any]:
    from pyhanko.pdf_utils.form_tools import populate_static_text_field
    from pyhanko.pdf_utils.incremental_writer import IncrementalPdfFileWriter
    from pyhanko.pdf_utils.reader import PdfFileReader
    from pyhanko.pdf_utils.text import TextBoxStyle

    # pyHanko 0.35.x updates /Info/XMP by default during `write`, which would
    # add an unrelated signed-document mutation.  This small, version-guarded
    # subclass is used only for the controlled field finalisation below.  The
    # public `populate_static_text_field` helper remains responsible for the
    # actual form appearance and field state change.
    class _ControlledIncrementalWriter(IncrementalPdfFileWriter):
        def _update_meta(self) -> None:
            return None

    source = Path(config["input"])
    output = Path(config["output"])
    with source.open("rb") as stream:
        writer = _ControlledIncrementalWriter(stream, strict=True)
        if writer.prev.encrypted:
            raise ProviderError("encrypted PDFs are outside this controlled transaction")
        before = document_snapshot(writer.prev, config["maxFields"], config["maxPages"])
        validate_form_surface(before, argparse.Namespace(**config), after=False)
        populate_static_text_field(
            writer,
            config["field"],
            TextBoxStyle(font_size=10),
            config["value"],
        )
        with output.open("xb") as target:
            writer.write(target)
            target.flush()
            os.fsync(target.fileno())
    with output.open("rb") as stream:
        reader = PdfFileReader(stream, strict=True)
        if reader.security_handler is not None:
            raise ProviderError("output unexpectedly became encrypted")
        after = document_snapshot(reader, config["maxFields"], config["maxPages"])
    validate_form_surface(after, argparse.Namespace(**config), after=True)
    if before["pageCount"] != after["pageCount"]:
        raise ProviderError("controlled form finalisation changed the page count")
    if before["catalog"] != after["catalog"] or before["info"] != after["info"]:
        raise ProviderError("controlled form finalisation changed catalog or metadata references")
    if not unchanged_non_target_fields(before["fields"], after["fields"], config["field"]):
        raise ProviderError("controlled form finalisation changed a non-target form field")
    return {"workerSchema": 1, "before": before, "after": after}


def inspect_worker(config: dict[str, Any]) -> dict[str, Any]:
    from pyhanko.pdf_utils.reader import PdfFileReader

    with Path(config["input"]).open("rb") as stream:
        reader = PdfFileReader(stream, strict=True)
        if reader.security_handler is not None:
            raise ProviderError("encrypted PDFs are outside this controlled transaction")
        return {"workerSchema": 1, "snapshot": document_snapshot(reader, config["maxFields"], config["maxPages"])}


def worker_main() -> int:
    logging.disable(logging.CRITICAL)
    try:
        raw = sys.stdin.buffer.read(MAX_WORKER_CONFIG_BYTES + 1)
        if len(raw) > MAX_WORKER_CONFIG_BYTES:
            raise ProviderError(f"worker configuration exceeds {MAX_WORKER_CONFIG_BYTES} bytes")
        config = json.loads(raw)
        if not isinstance(config, dict):
            raise ProviderError("worker configuration must be a JSON object")
        provider_versions()
        operation = config.get("operation")
        if operation == "inspect":
            result = inspect_worker(config)
        elif operation == "fill":
            result = fill_worker(config)
        else:
            raise ProviderError(f"unsupported worker operation: {operation!r}")
        print(json.dumps(result, sort_keys=True, separators=(",", ":")))
        return 0
    except Exception as exc:
        print(bounded_text(f"{type(exc).__name__}: {exc}"), file=sys.stderr)
        return 2


def run_worker(config: dict[str, Any], args: argparse.Namespace) -> dict[str, Any]:
    process = subprocess.Popen(
        [sys.executable, str(Path(__file__).resolve()), "_worker"],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        shell=False,
        env={**os.environ, "PYTHONDONTWRITEBYTECODE": "1"},
        start_new_session=os.name != "nt",
    )
    stdout = bytearray()
    stderr = bytearray()
    violations: list[str] = []
    lock = threading.Lock()

    def stop(message: str) -> None:
        with lock:
            if not violations:
                violations.append(message)
        try:
            if os.name != "nt":
                os.killpg(process.pid, 9)
            else:
                process.kill()
        except OSError:
            pass

    def pump(stream: Any, target: bytearray, limit: int, label: str) -> None:
        try:
            while True:
                chunk = stream.read(64 * 1024)
                if not chunk:
                    break
                if len(target) + len(chunk) > limit:
                    stop(f"certified-form worker {label} exceeded the {limit} byte budget")
                    break
                target.extend(chunk)
        finally:
            stream.close()

    out_thread = threading.Thread(target=pump, args=(process.stdout, stdout, args.max_stdout_bytes, "stdout"), daemon=True)
    err_thread = threading.Thread(target=pump, args=(process.stderr, stderr, args.max_stderr_bytes, "stderr"), daemon=True)
    out_thread.start()
    err_thread.start()
    encoded = json.dumps(config, separators=(",", ":")).encode("utf-8")
    assert process.stdin is not None
    process.stdin.write(encoded)
    process.stdin.close()
    timed_out = False
    try:
        process.wait(timeout=args.timeout_seconds)
    except subprocess.TimeoutExpired:
        timed_out = True
        stop(f"certified-form worker timed out after {args.timeout_seconds} seconds")
        process.wait()
    out_thread.join()
    err_thread.join()
    if timed_out:
        raise ProviderError(f"certified-form worker timed out after {args.timeout_seconds} seconds")
    if violations:
        raise ProviderError(violations[0])
    stderr_text = stderr.decode("utf-8", "replace").strip()
    if process.returncode != 0:
        raise ProviderError(f"certified-form worker failed (exit {process.returncode}): {bounded_text(stderr_text)}")
    if stderr_text:
        raise ProviderError(f"certified-form worker emitted unexpected diagnostics: {bounded_text(stderr_text)}")
    try:
        result = json.loads(stdout.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ProviderError("certified-form worker did not return one valid UTF-8 JSON document") from exc
    if not isinstance(result, dict) or result.get("workerSchema") != 1:
        raise ProviderError("certified-form worker returned an unsupported result schema")
    return result


def validate_limits(args: argparse.Namespace) -> None:
    for label, value, maximum in (
        ("--max-input-bytes", args.max_input_bytes, DEFAULT_MAX_INPUT_BYTES),
        ("--max-output-bytes", args.max_output_bytes, DEFAULT_MAX_OUTPUT_BYTES),
        ("--max-pages", args.max_pages, DEFAULT_MAX_PAGES),
        ("--max-fields", args.max_fields, DEFAULT_MAX_FIELDS),
        ("--timeout-seconds", args.timeout_seconds, DEFAULT_TIMEOUT_SECONDS),
        ("--max-stdout-bytes", args.max_stdout_bytes, DEFAULT_MAX_STDOUT_BYTES),
        ("--max-stderr-bytes", args.max_stderr_bytes, DEFAULT_MAX_STDERR_BYTES),
    ):
        if value > maximum:
            raise ProviderError(f"{label} cannot exceed the hard maximum {maximum}")


def probe() -> dict[str, Any]:
    versions = provider_versions()
    try:
        from pyhanko.pdf_utils.form_tools import populate_static_text_field  # noqa: F401
        from pyhanko.pdf_utils.incremental_writer import IncrementalPdfFileWriter  # noqa: F401
        from pyhanko.pdf_utils.text import TextBoxStyle  # noqa: F401
    except Exception as exc:
        raise ProviderError(f"pyHanko controlled form APIs could not be imported: {bounded_text(exc)}") from exc
    return {
        "schema": SCHEMA,
        "ok": True,
        "provider": "pyhanko",
        "providerVersion": versions["pyHanko"],
        "certvalidatorVersion": versions["pyhankoCertvalidator"],
        "operation": "fill-certified-docmdp-p2-text-field",
        "savePolicies": ["incremental"],
        "input": "existing-certified-pdf",
        "supported": {
            "docmdpPermission": "fill-forms",
            "fieldmdp": "include-exact-one-locked-field",
            "target": "one-empty-flat-visible-text-field",
            "value": "canonical-non-negative-decimal",
            "output": "visible-static-read-only-text-field",
        },
        "networkAllowed": False,
        "silentFallback": False,
        "padesProfileConformanceClaimed": False,
        "limitations": [
            "this is not a general signed-PDF editor or arbitrary AcroForm filler",
            "it requires exactly one trusted certification signature and rejects existing post-certification revisions",
            "timestamps, LTV/DSS, PKCS#11, remote signing, and dynamic XFA are outside this operation",
        ],
    }


def fill(args: argparse.Namespace) -> dict[str, Any]:
    validate_limits(args)
    input_trust = require_input_trust(args)
    source = regular_input_path(args.input, "input PDF", args.max_input_bytes, pdf=True)
    destination = safe_destination(args.output, source)
    expected = expected_hash(args.expected_source_sha256)
    args.field = acroform_name(args.field, "--field")
    args.expected_signature_field = acroform_name(args.expected_signature_field, "--expected-signature-field")
    args.expected_locked_field = acroform_name(args.expected_locked_field, "--expected-locked-field")
    args.value = decimal_value(args.value)
    args.expected_locked_value = limited_text(args.expected_locked_value, "--expected-locked-value")
    if args.field == args.expected_locked_field:
        raise ProviderError("--field and --expected-locked-field must be different")
    if sha256(source) != expected:
        raise ProviderError(f"source SHA-256 mismatch: expected {expected}, received {sha256(source)}")
    trust_root = regular_input_path(args.trust_root, "trust root", DEFAULT_MAX_CERTIFICATE_BYTES)
    trust_hash = sha256(trust_root)
    versions = provider_versions()

    with tempfile.TemporaryDirectory(prefix="office-kit-certified-form-source-") as source_temporary:
        source_root = Path(source_temporary)
        source_snapshot = source_root / "source.pdf"
        root_snapshot = source_root / "trust-root.pem"
        snapshot_file(source, source_snapshot, expected)
        snapshot_file(trust_root, root_snapshot, trust_hash)
        before_worker = run_worker({
            "operation": "inspect", "input": str(source_snapshot),
            "maxFields": args.max_fields, "maxPages": args.max_pages,
        }, args)
        before = before_worker["snapshot"]
        validate_form_surface(before, args, after=False)
        baseline_validation = validate_signature(source_snapshot, expected, root_snapshot, args, "preflight")
        baseline_signature = preflight_signature(baseline_validation, args)
        if sha256(source_snapshot) != expected or sha256(source) != expected:
            raise ProviderError("source PDF changed during controlled-form preflight")
        if sha256(root_snapshot) != trust_hash or sha256(trust_root) != trust_hash:
            raise ProviderError("trust root changed during controlled-form preflight")

        with tempfile.TemporaryDirectory(prefix=".office-kit-certified-form-", dir=destination.parent) as staging_temporary:
            staging_root = Path(staging_temporary)
            candidate = staging_root / "candidate.pdf"
            worker = run_worker({
                "operation": "fill",
                "input": str(source_snapshot),
                "output": str(candidate),
                "maxFields": args.max_fields,
                "maxPages": args.max_pages,
                "field": args.field,
                "value": args.value,
                "expected_signature_field": args.expected_signature_field,
                "expected_locked_field": args.expected_locked_field,
                "expected_locked_value": args.expected_locked_value,
                "trusted_input": args.trusted_input,
                "caller_isolated": args.caller_isolated,
            }, args)
            if not candidate.is_file() or candidate.is_symlink():
                raise ProviderError("controlled-form worker did not create one regular candidate PDF")
            candidate_bytes = candidate.stat().st_size
            if candidate_bytes <= source_snapshot.stat().st_size or candidate_bytes > args.max_output_bytes:
                raise ProviderError(f"candidate output size {candidate_bytes} is outside the incremental output budget")
            if not has_exact_prefix(source_snapshot, candidate):
                raise ProviderError("candidate does not preserve the complete source byte prefix")
            candidate_hash = sha256(candidate)
            post_validation = validate_signature(candidate, candidate_hash, root_snapshot, args, "postflight")
            post_signature = postflight_signature(post_validation, args)
            after = worker["after"]
            if baseline_validation["revisionCount"] + 1 != post_validation["revisionCount"]:
                raise ProviderError("controlled form finalisation did not add exactly one revision")
            if before != worker["before"]:
                raise ProviderError("worker preflight snapshot did not match the source-bound parent preflight")
            validate_form_surface(after, args, after=True)
            if before["pageCount"] != after["pageCount"] or before["catalog"] != after["catalog"] or before["info"] != after["info"]:
                raise ProviderError("candidate changed page count, catalog references, or metadata references")
            if not unchanged_non_target_fields(before["fields"], after["fields"], args.field):
                raise ProviderError("candidate changed a non-target field")
            if sha256(source_snapshot) != expected or sha256(source) != expected:
                raise ProviderError("source PDF changed before output publication")
            if sha256(root_snapshot) != trust_hash or sha256(trust_root) != trust_hash:
                raise ProviderError("trust root changed before output publication")
            publish_without_replace(candidate, destination)

    output_hash = sha256(destination)
    if output_hash != candidate_hash:
        raise ProviderError("published output hash differs from the validated candidate")
    return {
        "schema": SCHEMA,
        "ok": True,
        "operationCompleted": True,
        "provider": {
            "name": "pyhanko",
            "version": versions["pyHanko"],
            "certvalidatorVersion": versions["pyhankoCertvalidator"],
            "python": {"executable": sys.executable, "version": sys.version.split()[0]},
        },
        "operation": "fill-certified-docmdp-p2-text-field",
        "inputTrust": input_trust,
        "networkAllowed": False,
        "silentFallback": False,
        "padesProfileConformanceClaimed": False,
        "source": {"path": str(source), "bytes": source.stat().st_size, "sha256": expected},
        "trustRoot": {"path": str(trust_root), "sha256": trust_hash},
        "output": {"path": str(destination), "bytes": destination.stat().st_size, "sha256": output_hash},
        "savePolicy": {
            "strategy": "incremental",
            "sourcePrefixPreserved": True,
            "revisionsBefore": baseline_validation["revisionCount"],
            "revisionsAfter": post_validation["revisionCount"],
            "metadataUpdateSuppressed": True,
        },
        "field": {
            "target": args.field,
            "value": args.value,
            "locked": {"name": args.expected_locked_field, "value": args.expected_locked_value},
            "before": before["fields"][args.field],
            "after": after["fields"][args.field],
            "nonTargetFieldsUnchanged": True,
        },
        "signature": {
            "field": args.expected_signature_field,
            "preflight": baseline_signature,
            "postflight": post_signature,
            "preflightValidation": baseline_validation,
            "postflightValidation": post_validation,
        },
        "transaction": {
            "distinctOutput": True,
            "noReplace": True,
            "privateSourceSnapshot": True,
            "privateTrustRootSnapshot": True,
            "sourceReprovedBeforePublish": True,
            "trustRootReprovedBeforePublish": True,
            "outputPublishedAtomically": True,
        },
        "limitations": [
            "the original certification signature remains valid for its signed revision; it does not approve arbitrary later revisions",
            "this operation finalises a target as static/read-only and is not an interactive form editor",
            "the report validates only the recorded explicit root with revocation policy none and does not claim PAdES/LTV conformance",
            "run independent structural inspection and native rendering before delivery",
        ],
    }


def add_common_limits(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--max-input-bytes", type=positive_int, default=DEFAULT_MAX_INPUT_BYTES)
    parser.add_argument("--max-output-bytes", type=positive_int, default=DEFAULT_MAX_OUTPUT_BYTES)
    parser.add_argument("--max-pages", type=positive_int, default=DEFAULT_MAX_PAGES)
    parser.add_argument("--max-fields", type=positive_int, default=DEFAULT_MAX_FIELDS)
    parser.add_argument("--timeout-seconds", type=positive_int, default=DEFAULT_TIMEOUT_SECONDS)
    parser.add_argument("--max-stdout-bytes", type=positive_int, default=DEFAULT_MAX_STDOUT_BYTES)
    parser.add_argument("--max-stderr-bytes", type=positive_int, default=DEFAULT_MAX_STDERR_BYTES)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)
    subparsers.add_parser("probe", help="require the bounded DocMDP P=2 form-finalisation surface")
    fill_parser = subparsers.add_parser("fill", help="finalise one explicitly allowed certified text field")
    fill_parser.add_argument("input")
    fill_parser.add_argument("output")
    fill_parser.add_argument("--expected-source-sha256", required=True)
    fill_parser.add_argument("--trust-root", required=True)
    fill_parser.add_argument("--field", required=True)
    fill_parser.add_argument("--value", required=True)
    fill_parser.add_argument("--expected-signature-field", required=True)
    fill_parser.add_argument("--expected-locked-field", required=True)
    fill_parser.add_argument("--expected-locked-value", required=True)
    trust = fill_parser.add_mutually_exclusive_group(required=True)
    trust.add_argument("--trusted-input", action="store_true")
    trust.add_argument("--caller-isolated", action="store_true")
    add_common_limits(fill_parser)
    return parser


def main() -> int:
    from python_runtime import reexec_configured_provider_python

    reexec_configured_provider_python()
    if len(sys.argv) > 1 and sys.argv[1] == "_worker":
        return worker_main()
    args = build_parser().parse_args()
    try:
        result = probe() if args.command == "probe" else fill(args)
        print(json.dumps(result, indent=2, sort_keys=True))
        return 0
    except (ProviderError, ValidationProviderError, SigningProviderError, OSError, ValueError) as exc:
        print(json.dumps({
            "ok": False,
            "provider": "pyhanko",
            "operation": "fill-certified-docmdp-p2-text-field" if getattr(args, "command", None) == "fill" else getattr(args, "command", None),
            "error": bounded_text(exc),
            "silentFallback": False,
        }, sort_keys=True), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
