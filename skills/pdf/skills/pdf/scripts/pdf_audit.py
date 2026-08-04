#!/usr/bin/env python3
"""Validate the canonical PDF operation audit envelope against delivered bytes."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import secrets
import sys


SCHEMA = "office-kit.pdf-audit.v1"
SHA256 = re.compile(r"^[0-9a-f]{64}$")
SAVE_POLICIES = {"read-only", "rewrite", "incremental", "sanitize"}


class AuditError(RuntimeError):
    pass


def digest(path: Path) -> dict[str, str | int]:
    payload = path.read_bytes()
    return {"bytes": len(payload), "sha256": hashlib.sha256(payload).hexdigest()}


def require_object(value, label: str) -> dict:
    if not isinstance(value, dict):
        raise AuditError(f"{label} must be an object")
    return value


def validate_file_evidence(evidence, actual_path: Path, label: str) -> None:
    record = require_object(evidence, label)
    if not isinstance(record.get("path"), str) or not record["path"].strip():
        raise AuditError(f"{label}.path must be a non-empty string")
    recorded_path = Path(record["path"]).expanduser()
    if not recorded_path.is_absolute() or recorded_path.resolve() != actual_path.expanduser().resolve():
        raise AuditError(f"{label}.path must be the exact absolute path {actual_path.expanduser().resolve()}")
    if not isinstance(record.get("bytes"), int) or record["bytes"] < 1:
        raise AuditError(f"{label}.bytes must be a positive integer")
    if not isinstance(record.get("sha256"), str) or not SHA256.fullmatch(record["sha256"]):
        raise AuditError(f"{label}.sha256 must be a lowercase SHA-256 digest")
    if not actual_path.is_file():
        raise AuditError(f"{label} file does not exist: {actual_path}")
    actual = digest(actual_path)
    if record["bytes"] != actual["bytes"] or record["sha256"] != actual["sha256"]:
        raise AuditError(f"{label} bytes/hash do not match {actual_path}")


def validate_input_evidence(evidence, actual_paths: list[Path]) -> int:
    if evidence is None and not actual_paths:
        return 0
    if not isinstance(evidence, list) or not evidence:
        raise AuditError("audit.inputs must be a non-empty array when --input is used")
    if not actual_paths:
        raise AuditError("repeat --input for every audit.inputs record so the validator can recompute bytes")
    if len(evidence) != len(actual_paths):
        raise AuditError(f"audit.inputs has {len(evidence)} records but {len(actual_paths)} --input paths were provided")
    remaining = list(evidence)
    for actual_path in actual_paths:
        resolved = actual_path.expanduser().resolve()
        match = next((item for item in remaining if isinstance(item, dict) and Path(str(item.get("path", ""))).expanduser().resolve() == resolved), None)
        if match is None:
            raise AuditError(f"audit.inputs has no record for {resolved}")
        validate_file_evidence(match, resolved, f"inputs[{evidence.index(match)}]")
        remaining.remove(match)
    return len(evidence)


def require_regular_file(path: Path, label: str) -> Path:
    candidate = path.expanduser()
    if candidate.is_symlink() or not candidate.is_file():
        raise AuditError(f"{label} must be a regular file")
    return candidate.resolve()


def require_docmdp_no_changes_verification(verification_path: Path, source: Path, trust_root: Path) -> dict:
    """Load a pyHanko report and retain only verified P=1 refusal evidence.

    This is intentionally narrower than a general signature editor: it accepts
    a complete, explicit-root pyHanko report for a source that currently has a
    fully valid DocMDP P=1 certification, then returns the exact facts that a
    no-mutation audit may publish. The caller still owns the requested change;
    this helper only records why no change may be made.
    """
    report_path = require_regular_file(verification_path, "signature verification report")
    root_path = require_regular_file(trust_root, "DocMDP trust root")
    try:
        report = json.loads(report_path.read_text(encoding="utf8"))
    except json.JSONDecodeError as exc:
        raise AuditError(f"signature verification report is not valid JSON: {exc}") from exc
    report = require_object(report, "signature verification report")
    if report.get("schema") != "office-kit.pyhanko-verify.v1":
        raise AuditError("signature verification report must use office-kit.pyhanko-verify.v1")
    validate_file_evidence(report.get("source"), source, "signature verification report.source")
    if report.get("sourceProtected") is not True or report.get("silentFallback") is not False:
        raise AuditError("signature verification report must prove a protected source and no fallback")
    if report.get("savePolicy") != "read-only":
        raise AuditError("signature verification report must use read-only save policy")

    provider = require_object(report.get("provider"), "signature verification report.provider")
    if provider.get("name") != "pyhanko" or not isinstance(provider.get("version"), str) or not provider["version"].strip():
        raise AuditError("signature verification report must name a versioned pyhanko provider")
    policy_gates = require_object(report.get("policyGates"), "signature verification report.policyGates")
    requested = require_object(policy_gates.get("requested"), "signature verification report.policyGates.requested")
    if policy_gates.get("passed") is not True or any(requested.get(field) is not True for field in (
        "requireSignature",
        "requireAllIntegrityValid",
        "requireAllTrusted",
        "requireDocMDPCompliant",
        "requireAllBottomLine",
    )):
        raise AuditError("signature verification report must pass all required integrity, trust, DocMDP, and bottom-line gates")

    validation_policy = require_object(report.get("validationPolicy"), "signature verification report.validationPolicy")
    if validation_policy.get("trustPolicy") != "explicit-roots":
        raise AuditError("signature verification report must use explicit-roots trust")
    root_evidence = absolute_file_evidence(root_path)
    roots = validation_policy.get("trustRoots")
    def matching_root(entry) -> bool:
        if not isinstance(entry, dict) or not isinstance(entry.get("path"), str):
            return False
        entry_path = Path(entry["path"]).expanduser()
        return (
            entry_path.is_absolute()
            and entry_path.resolve() == root_path
            and entry.get("bytes") == root_evidence["bytes"]
            and entry.get("sha256") == root_evidence["sha256"]
        )
    if not isinstance(roots, list) or not any(
        matching_root(entry)
        for entry in roots
    ):
        raise AuditError("signature verification report does not bind the selected explicit trust root")

    summary = require_object(report.get("summary"), "signature verification report.summary")
    if any(summary.get(field) is not True for field in (
        "allBottomLine",
        "allDocMDPCompliant",
        "allIntegrityValid",
        "allTrusted",
        "allValidationCompleted",
    )):
        raise AuditError("signature verification summary is incomplete")
    signatures = report.get("signatures")
    if not isinstance(signatures, list) or not signatures:
        raise AuditError("signature verification report must contain at least one signature")
    matching = []
    for signature in signatures:
        if not isinstance(signature, dict):
            continue
        docmdp = signature.get("docMDP")
        if not isinstance(docmdp, dict) or docmdp.get("present") is not True or docmdp.get("permissionCode") != 1:
            continue
        if all(signature.get(field) is True for field in (
            "intact",
            "trusted",
            "bottomLine",
            "docMDPCompliant",
            "validationCompleted",
        )) and signature.get("coverage") == "entire-file" and signature.get("modificationLevel") == "none":
            matching.append(signature)
    if len(matching) != 1:
        raise AuditError("signature verification report must prove exactly one intact, trusted DocMDP P=1 certification")
    signature = matching[0]
    field_name = signature.get("fieldName")
    if not isinstance(field_name, str) or not field_name.strip():
        raise AuditError("DocMDP certification field name is missing")
    return {
        "signaturePolicy": {
            "certificationField": field_name,
            "certificationSignaturePresent": True,
            "docMDP": {
                "permission": "no-changes",
                "permissionCode": 1,
                "requestedChangeAllowed": False,
            },
            "policyDecision": "refuse_without_mutation",
        },
        "signatureVerification": {
            "conclusion": "valid-under-selected-policy",
            "cryptographicallyValid": True,
            "intact": True,
            "trusted": True,
            "trustPolicy": "explicit-roots",
            "trustRoot": root_evidence,
            "bottomLine": True,
            "coverage": "entire-file",
            "docMDPCompliantBeforeRequestedChange": True,
            "revisionCount": report.get("revisionCount"),
            "signatureCount": len(signatures),
            "signedRevision": signature.get("signedRevision"),
        },
    }


def validate_docmdp_no_changes_record(record: dict, trust_root: Path) -> None:
    policy = require_object(record.get("signaturePolicy"), "signaturePolicy")
    docmdp = require_object(policy.get("docMDP"), "signaturePolicy.docMDP")
    if policy.get("certificationSignaturePresent") is not True or policy.get("policyDecision") != "refuse_without_mutation":
        raise AuditError("DocMDP refusal audit must record a certification no-mutation decision")
    if docmdp.get("permission") != "no-changes" or docmdp.get("permissionCode") != 1 or docmdp.get("requestedChangeAllowed") is not False:
        raise AuditError("DocMDP refusal audit must record P=1 no-changes policy")
    verification = require_object(require_object(record.get("validation"), "validation").get("signatureVerification"), "validation.signatureVerification")
    if any(verification.get(field) is not True for field in (
        "cryptographicallyValid",
        "intact",
        "trusted",
        "bottomLine",
        "docMDPCompliantBeforeRequestedChange",
    )) or verification.get("conclusion") != "valid-under-selected-policy" or verification.get("trustPolicy") != "explicit-roots" or verification.get("coverage") != "entire-file":
        raise AuditError("DocMDP refusal audit has incomplete signature validation evidence")
    validate_file_evidence(verification.get("trustRoot"), require_regular_file(trust_root, "DocMDP trust root"), "validation.signatureVerification.trustRoot")


def validate_record(
    record: dict,
    source: Path,
    inputs: list[Path],
    artifact: Path | None,
    required_operation: str | None,
    artifacts: dict[str, Path] | None = None,
) -> dict:
    record = require_object(record, "audit")
    if record.get("schema") != SCHEMA:
        raise AuditError(f"schema must be {SCHEMA!r}")
    status = record.get("status")
    if status not in {"succeeded", "failed_closed"}:
        raise AuditError("status must be 'succeeded' or 'failed_closed'")
    validate_file_evidence(record.get("source"), source, "source")
    input_count = validate_input_evidence(record.get("inputs"), inputs)

    provider = require_object(record.get("provider"), "provider")
    for field in ("actual", "version"):
        if not isinstance(provider.get(field), str) or not provider[field].strip():
            raise AuditError(f"provider.{field} must be a non-empty string")
    if provider.get("silentFallback") is not False:
        raise AuditError("provider.silentFallback must be false")

    policy = require_object(record.get("savePolicy"), "savePolicy")
    if policy.get("strategy") not in SAVE_POLICIES:
        raise AuditError("savePolicy.strategy must be read-only, rewrite, incremental, or sanitize")
    preflight = require_object(record.get("preflight"), "preflight")
    if not isinstance(preflight.get("probeCompleted"), bool) or not isinstance(preflight.get("planCompleted"), bool):
        raise AuditError("preflight.probeCompleted and preflight.planCompleted must be booleans")
    operation = require_object(record.get("operation"), "operation")
    if not isinstance(operation.get("type"), str) or not operation["type"].strip():
        raise AuditError("operation.type must be a non-empty string")
    if required_operation and operation["type"] != required_operation:
        raise AuditError(f"operation.type must be {required_operation!r}")
    validation = require_object(record.get("validation"), "validation")
    if operation["type"] == "extract-attachments":
        if policy.get("strategy") != "read-only":
            raise AuditError("extract-attachments audits must use savePolicy.strategy read-only")
        if validation.get("attachmentsOpenedOrExecuted") is not False:
            raise AuditError(
                "extract-attachments audits must record validation.attachmentsOpenedOrExecuted as false"
            )

    output_records = record.get("outputs")
    if status == "succeeded":
        if preflight["probeCompleted"] is not True or preflight["planCompleted"] is not True:
            raise AuditError("a succeeded audit requires completed provider probe and route plan")
        if output_records is not None:
            if not isinstance(output_records, dict) or not output_records:
                raise AuditError("audit.outputs must be a non-empty object")
            if not artifacts:
                raise AuditError("--artifact-json/--artifact-csv are required when audit.outputs is used")
            if set(output_records) != set(artifacts):
                raise AuditError(
                    f"audit.outputs keys {sorted(output_records)} do not match artifact keys {sorted(artifacts)}; provide matching --artifact-json/--artifact-csv flags"
                )
            for name, output_path in artifacts.items():
                validate_file_evidence(output_records.get(name), output_path, f"outputs.{name}")
            if artifact is not None:
                validate_file_evidence(record.get("output"), artifact, "output")
            elif "json" in artifacts:
                validate_file_evidence(record.get("output"), artifacts["json"], "output")
        else:
            if artifact is None:
                raise AuditError("--artifact is required for a succeeded audit")
            validate_file_evidence(record.get("output"), artifact, "output")
    else:
        if record.get("output") is not None or output_records is not None:
            raise AuditError("failed_closed audit outputs must be null/absent")
        if artifact is not None and artifact.exists():
            raise AuditError("failed_closed audit must not have a partial artifact")
        if not isinstance(record.get("reason"), str) or not record["reason"].strip():
            raise AuditError("failed_closed audit requires a non-empty reason")

    return {
        "ok": True,
        "schema": SCHEMA,
        "status": status,
        "provider": provider["actual"],
        "providerVersion": provider["version"],
        "savePolicy": policy["strategy"],
        "operation": operation["type"],
        "inputs": input_count,
        "outputs": len(output_records) if isinstance(output_records, dict) else 1,
        "silentFallback": False,
    }


def absolute_file_evidence(path: Path) -> dict[str, str | int]:
    resolved = path.expanduser().resolve()
    return {"path": str(resolved), **digest(resolved)}


def write_new_json(path: Path, record: dict) -> Path:
    """Atomically publish a new audit without overwriting an earlier record."""
    target = path.expanduser().resolve()
    if target.exists() or target.is_symlink():
        raise AuditError(f"refuses to overwrite existing audit: {target}")
    target.parent.mkdir(parents=True, exist_ok=True)
    temporary = target.with_name(f".{target.name}.tmp-{os.getpid()}-{secrets.token_hex(8)}")
    try:
        with temporary.open("x", encoding="utf8") as handle:
            json.dump(record, handle, indent=2, sort_keys=True)
            handle.write("\n")
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary, target)
    except Exception:
        temporary.unlink(missing_ok=True)
        raise
    return target


def create_failed_closed_record(args) -> dict:
    source = args.source.expanduser().resolve()
    audit = args.audit.expanduser().resolve()
    if not source.is_file():
        raise AuditError(f"source file does not exist: {source}")
    if source == audit:
        raise AuditError("audit path must be different from source")
    if not args.provider.strip() or not args.provider_version.strip():
        raise AuditError("provider and provider version must be non-empty")
    if not args.operation.strip() or not args.reason.strip():
        raise AuditError("operation and reason must be non-empty")
    if audit.exists() or audit.is_symlink():
        raise AuditError(f"refuses to overwrite existing audit: {audit}")
    docmdp_evidence = None
    if args.require_docmdp_no_changes:
        if args.signature_verification is None or args.trust_root is None:
            raise AuditError("--require-docmdp-no-changes requires --signature-verification and --trust-root")
        if args.provider != "pyhanko" or args.strategy != "read-only":
            raise AuditError("DocMDP P=1 refusal requires provider pyhanko and read-only strategy")
        if not args.probe_completed or not args.plan_completed or not args.source_inspected:
            raise AuditError("DocMDP P=1 refusal requires completed probe, plan, and source inspection")
        docmdp_evidence = require_docmdp_no_changes_verification(args.signature_verification, source, args.trust_root)
    elif args.signature_verification is not None or args.trust_root is not None:
        raise AuditError("--signature-verification and --trust-root require --require-docmdp-no-changes")

    capabilities = None
    if args.capabilities_json is not None:
        capabilities_path = args.capabilities_json.expanduser().resolve()
        if not capabilities_path.is_file():
            raise AuditError(f"capabilities JSON does not exist: {capabilities_path}")
        try:
            capabilities = json.loads(capabilities_path.read_text("utf8"))
        except (OSError, UnicodeError, json.JSONDecodeError) as error:
            raise AuditError(f"capabilities JSON is not valid UTF-8 JSON: {capabilities_path}: {error}") from error
        if not isinstance(capabilities, dict) or not capabilities:
            raise AuditError("capabilities JSON must be a non-empty object")

    # A failed-closed task publishes only its audit in the delivery directory.
    # This makes the typed no-artifact assertion an observed local fact instead
    # of a prose promise. Diagnostics belong under tmp/ or another caller-owned
    # directory, not alongside the delivered audit.
    try:
        existing_entries = sorted(entry.name for entry in audit.parent.iterdir())
    except FileNotFoundError:
        existing_entries = []
    if existing_entries:
        raise AuditError(
            f"failed-closed output directory must be empty before publishing audit: {audit.parent} contains {existing_entries}"
        )

    provider_record = {
        "actual": args.provider,
        "version": args.provider_version,
        "silentFallback": False,
    }
    if capabilities is not None:
        provider_record["capabilities"] = capabilities

    record = {
        "schema": SCHEMA,
        "status": "failed_closed",
        "delivered_modified_pdf": False,
        "source": absolute_file_evidence(source),
        "output": None,
        "provider": provider_record,
        "savePolicy": {
            "strategy": args.strategy,
            "sourceOverwrite": False,
            "artifactWritten": False,
            "publication": "audit_only",
        },
        "preflight": {
            "probeCompleted": bool(args.probe_completed),
            "planCompleted": bool(args.plan_completed),
            "sourceInspectionCompleted": bool(args.source_inspected),
        },
        "operation": {
            "type": args.operation,
            "mutationAttempted": False,
            "performed": False,
            "result": "not_attempted",
        },
        "validation": {
            "sourceIdentity": {"sourcePreserved": True},
            "artifactChecks": {
                "modifiedPdfPresent": False,
                "partialArtifactPresent": False,
            },
            "outputDirectory": {
                "path": str(audit.parent),
                "entriesBeforeAudit": [],
                "auditOnly": True,
            },
        },
        "reason": args.reason,
    }
    if docmdp_evidence is not None:
        record["signaturePolicy"] = docmdp_evidence["signaturePolicy"]
        record["validation"]["signatureVerification"] = docmdp_evidence["signatureVerification"]
        record["warnings"] = [
            "DocMDP P=1 permits no post-certification changes.",
            "An intact signed ByteRange does not authorize a later revision.",
        ]
    validate_record(record, source, [], None, args.operation)
    if docmdp_evidence is not None:
        validate_docmdp_no_changes_record(record, args.trust_root)
    return record


def parser() -> argparse.ArgumentParser:
    root = argparse.ArgumentParser(description=__doc__)
    subparsers = root.add_subparsers(dest="command", required=True)
    validate = subparsers.add_parser("validate", help="validate a canonical audit and recompute byte evidence")
    validate.add_argument("audit", type=Path)
    validate.add_argument("--source", type=Path, required=True)
    validate.add_argument("--input", type=Path, action="append", default=[])
    validate.add_argument("--artifact", type=Path)
    validate.add_argument("--artifact-json", type=Path, help="validate the json member of a multi-output audit")
    validate.add_argument("--artifact-csv", type=Path, help="validate the csv member of a multi-output audit")
    validate.add_argument("--require-operation")
    validate.add_argument("--require-docmdp-no-changes", action="store_true")
    validate.add_argument("--trust-root", type=Path)
    failed_closed = subparsers.add_parser(
        "failed-closed",
        help="atomically write a canonical audit-only no-mutation refusal",
    )
    failed_closed.add_argument("audit", type=Path)
    failed_closed.add_argument("--source", type=Path, required=True)
    failed_closed.add_argument("--provider", required=True)
    failed_closed.add_argument("--provider-version", required=True)
    failed_closed.add_argument("--operation", required=True)
    failed_closed.add_argument("--reason", required=True)
    failed_closed.add_argument("--strategy", choices=sorted(SAVE_POLICIES), default="read-only")
    failed_closed.add_argument("--probe-completed", action="store_true")
    failed_closed.add_argument("--plan-completed", action="store_true")
    failed_closed.add_argument("--source-inspected", action="store_true")
    failed_closed.add_argument("--signature-verification", type=Path)
    failed_closed.add_argument("--require-docmdp-no-changes", action="store_true")
    failed_closed.add_argument("--trust-root", type=Path)
    failed_closed.add_argument(
        "--capabilities-json",
        type=Path,
        help="read a provider capability object and bind it under provider.capabilities",
    )
    return root


def main() -> int:
    from python_runtime import reexec_configured_provider_python
    reexec_configured_provider_python()
    args = parser().parse_args()
    try:
        if args.command == "failed-closed":
            record = create_failed_closed_record(args)
            target = write_new_json(args.audit, record)
            print(json.dumps({
                "ok": True,
                "schema": SCHEMA,
                "status": "failed_closed",
                "audit": str(target),
                "source": record["source"],
                "provider": record["provider"]["actual"],
                "providerVersion": record["provider"]["version"],
                "operation": record["operation"]["type"],
                "silentFallback": False,
            }, indent=2, sort_keys=True))
            return 0
        record = json.loads(args.audit.read_text("utf8"))
        artifacts = {
            name: value
            for name, value in (("json", args.artifact_json), ("csv", args.artifact_csv))
            if value is not None
        }
        result = validate_record(record, args.source, args.input, args.artifact, args.require_operation, artifacts or None)
        if args.require_docmdp_no_changes:
            if args.trust_root is None:
                raise AuditError("--require-docmdp-no-changes requires --trust-root")
            validate_docmdp_no_changes_record(record, args.trust_root)
        elif args.trust_root is not None:
            raise AuditError("--trust-root requires --require-docmdp-no-changes")
        print(json.dumps(result, indent=2, sort_keys=True))
        return 0
    except (AuditError, OSError, json.JSONDecodeError) as exc:
        print(json.dumps({"ok": False, "error": str(exc)}), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
