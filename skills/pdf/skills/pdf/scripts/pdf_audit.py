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


def validate_record(record: dict, source: Path, inputs: list[Path], artifact: Path | None, required_operation: str | None) -> dict:
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
    require_object(record.get("validation"), "validation")

    if status == "succeeded":
        if preflight["probeCompleted"] is not True or preflight["planCompleted"] is not True:
            raise AuditError("a succeeded audit requires completed provider probe and route plan")
        if artifact is None:
            raise AuditError("--artifact is required for a succeeded audit")
        validate_file_evidence(record.get("output"), artifact, "output")
    else:
        if record.get("output") is not None:
            raise AuditError("failed_closed audit output must be null")
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

    record = {
        "schema": SCHEMA,
        "status": "failed_closed",
        "delivered_modified_pdf": False,
        "source": absolute_file_evidence(source),
        "output": None,
        "provider": {
            "actual": args.provider,
            "version": args.provider_version,
            "silentFallback": False,
        },
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
    validate_record(record, source, [], None, args.operation)
    return record


def parser() -> argparse.ArgumentParser:
    root = argparse.ArgumentParser(description=__doc__)
    subparsers = root.add_subparsers(dest="command", required=True)
    validate = subparsers.add_parser("validate", help="validate a canonical audit and recompute byte evidence")
    validate.add_argument("audit", type=Path)
    validate.add_argument("--source", type=Path, required=True)
    validate.add_argument("--input", type=Path, action="append", default=[])
    validate.add_argument("--artifact", type=Path)
    validate.add_argument("--require-operation")
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
        print(json.dumps(validate_record(record, args.source, args.input, args.artifact, args.require_operation), indent=2, sort_keys=True))
        return 0
    except (AuditError, OSError, json.JSONDecodeError) as exc:
        print(json.dumps({"ok": False, "error": str(exc)}), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
