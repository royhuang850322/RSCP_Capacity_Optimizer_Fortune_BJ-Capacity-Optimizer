"""Shared helpers for signing Fortune BJ Optimizer license files."""
from __future__ import annotations

import base64
import json
import os
import re
import shutil
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey


PRODUCT_NAME = "Fortune BJ Optimizer"
DEFAULT_TOOL_REPOSITORY_NAME = "fortune_bj_optimizer"
DEFAULT_FEATURES = {
    "global_scheduling": True,
    "excel_report": True,
}


def default_license_admin_root() -> Path:
    configured = os.environ.get("RSCP_LICENSE_ADMIN_ROOT", "").strip()
    if configured:
        return Path(configured)
    default_root = Path(r"D:\RSCP_License_Admin")
    if Path(r"D:\\").exists():
        return default_root
    return Path.cwd() / "RSCP_License_Admin"


def canonical_license_bytes(payload: dict[str, Any]) -> bytes:
    signable = {key: value for key, value in payload.items() if key != "signature"}
    return json.dumps(
        signable,
        sort_keys=True,
        separators=(",", ":"),
        ensure_ascii=False,
    ).encode("utf-8")


def load_private_key(path: str) -> Ed25519PrivateKey:
    with open(path, "rb") as handle:
        key = serialization.load_pem_private_key(handle.read(), password=None)
    if not isinstance(key, Ed25519PrivateKey):
        raise TypeError("Private key is not an Ed25519 key.")
    return key


def parse_iso_date(raw_value: str, field_name: str) -> date:
    try:
        return date.fromisoformat(raw_value.strip())
    except ValueError as exc:
        raise ValueError(f"{field_name} must use YYYY-MM-DD format.") from exc


def generate_default_license_id(prefix: str = "LIC") -> str:
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    return f"{prefix}-{timestamp}"


def sanitize_path_component(value: str, fallback: str = "UNKNOWN") -> str:
    sanitized = re.sub(r'[<>:"/\\|?*\x00-\x1F]+', "_", str(value or "").strip())
    sanitized = sanitized.strip(" .")
    return sanitized or fallback


def ensure_customer_tool_dirs(
    customer_name: str,
    *,
    admin_root: str | Path | None = None,
    tool_name: str = DEFAULT_TOOL_REPOSITORY_NAME,
) -> dict[str, Path]:
    root = Path(admin_root or default_license_admin_root())
    customer_dir = sanitize_path_component(customer_name, "UNKNOWN_CUSTOMER")
    tool_dir = sanitize_path_component(tool_name, DEFAULT_TOOL_REPOSITORY_NAME)
    base = root / customer_dir / tool_dir
    paths = {
        "base": base,
        "requests": base / "requests",
        "issued": base / "issued",
        "active": base / "active",
        "archive": base / "archive",
        "notes": base / "notes",
    }
    for path in paths.values():
        path.mkdir(parents=True, exist_ok=True)
    return paths


def build_issued_license_path(
    customer_name: str,
    license_id: str,
    *,
    admin_root: str | Path | None = None,
    tool_name: str = DEFAULT_TOOL_REPOSITORY_NAME,
) -> Path:
    dirs = ensure_customer_tool_dirs(customer_name, admin_root=admin_root, tool_name=tool_name)
    file_name = f"{sanitize_path_component(license_id, 'LICENSE')}.json"
    return dirs["issued"] / file_name


def build_active_license_path(
    customer_name: str,
    *,
    admin_root: str | Path | None = None,
    tool_name: str = DEFAULT_TOOL_REPOSITORY_NAME,
) -> Path:
    dirs = ensure_customer_tool_dirs(customer_name, admin_root=admin_root, tool_name=tool_name)
    return dirs["active"] / "license.json"


def archive_existing_active_license(
    customer_name: str,
    *,
    admin_root: str | Path | None = None,
    tool_name: str = DEFAULT_TOOL_REPOSITORY_NAME,
) -> Path | None:
    dirs = ensure_customer_tool_dirs(customer_name, admin_root=admin_root, tool_name=tool_name)
    active_path = dirs["active"] / "license.json"
    if not active_path.exists():
        return None
    try:
        payload = json.loads(active_path.read_text(encoding="utf-8"))
        license_id = sanitize_path_component(str(payload.get("license_id") or "active_license"))
    except Exception:
        license_id = "active_license"
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    archive_path = dirs["archive"] / f"{license_id}_{timestamp}.json"
    shutil.move(str(active_path), str(archive_path))
    return archive_path


def activate_issued_license(
    issued_path: str | Path,
    customer_name: str,
    *,
    admin_root: str | Path | None = None,
    tool_name: str = DEFAULT_TOOL_REPOSITORY_NAME,
) -> Path:
    issued = Path(issued_path)
    if not issued.exists():
        raise FileNotFoundError(f"Issued license not found: {issued}")
    active_path = build_active_license_path(customer_name, admin_root=admin_root, tool_name=tool_name)
    archive_existing_active_license(customer_name, admin_root=admin_root, tool_name=tool_name)
    shutil.copy2(str(issued), str(active_path))
    return active_path


def copy_machine_request_to_admin(
    source_path: str | Path,
    customer_name: str,
    *,
    admin_root: str | Path | None = None,
    tool_name: str = DEFAULT_TOOL_REPOSITORY_NAME,
    machine_label: str = "",
) -> Path:
    source = Path(source_path)
    if not source.exists():
        raise FileNotFoundError(f"Machine fingerprint file not found: {source}")
    dirs = ensure_customer_tool_dirs(customer_name, admin_root=admin_root, tool_name=tool_name)
    label = sanitize_path_component(machine_label, "UNKNOWN_MACHINE")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = dirs["requests"] / f"machine_fingerprint_{label}_{timestamp}.json"
    shutil.copy2(str(source), str(dest))
    return dest


def build_license_payload(
    *,
    license_id: str,
    license_type: str,
    customer_name: str,
    customer_id: str,
    issue_date: str,
    expiry_date: str,
    binding_mode: str,
    machine_fingerprint: str = "",
    machine_label: str = "",
    note: str = "",
    features: dict[str, Any] | None = None,
) -> dict[str, Any]:
    license_id = license_id.strip()
    license_type = license_type.strip()
    customer_name = customer_name.strip()
    customer_id = customer_id.strip()
    binding_mode = binding_mode.strip().lower()
    machine_fingerprint = machine_fingerprint.strip()
    machine_label = machine_label.strip()
    note = note.strip()

    if not license_id:
        raise ValueError("license_id is required.")
    if not license_type:
        raise ValueError("license_type is required.")
    if not customer_name:
        raise ValueError("customer_name is required.")
    if not customer_id:
        raise ValueError("customer_id is required.")
    issue_date_value = parse_iso_date(issue_date, "issue_date")
    expiry_date_value = parse_iso_date(expiry_date, "expiry_date")
    if issue_date_value > expiry_date_value:
        raise ValueError("issue_date cannot be later than expiry_date.")
    if binding_mode not in {"unbound", "machine_locked"}:
        raise ValueError("binding_mode must be 'unbound' or 'machine_locked'.")
    if binding_mode == "machine_locked" and not machine_fingerprint:
        raise ValueError("machine_fingerprint is required for machine_locked licenses.")

    return {
        "license_version": 1,
        "product_name": PRODUCT_NAME,
        "license_id": license_id,
        "license_type": license_type,
        "customer_name": customer_name,
        "customer_id": customer_id,
        "issue_date": issue_date_value.isoformat(),
        "expiry_date": expiry_date_value.isoformat(),
        "binding_mode": binding_mode,
        "machine_fingerprint": machine_fingerprint,
        "machine_label": machine_label,
        "features": dict(features or DEFAULT_FEATURES),
        "note": note,
    }


def build_trial_license_payload(
    *,
    license_id: str,
    customer_name: str,
    customer_id: str,
    days_valid: int,
    issue_date: str = "",
    note: str = "Trial license",
    features: dict[str, Any] | None = None,
) -> dict[str, Any]:
    if days_valid <= 0:
        raise ValueError("days_valid must be greater than zero.")
    issue_date_value = parse_iso_date(issue_date, "issue_date") if issue_date.strip() else date.today()
    expiry_date_value = issue_date_value + timedelta(days=days_valid - 1)
    return build_license_payload(
        license_id=license_id,
        license_type="trial",
        customer_name=customer_name,
        customer_id=customer_id,
        issue_date=issue_date_value.isoformat(),
        expiry_date=expiry_date_value.isoformat(),
        binding_mode="unbound",
        machine_fingerprint="",
        machine_label="",
        note=note,
        features=features,
    )


def sign_license_payload(payload: dict[str, Any], private_key_path: str) -> dict[str, Any]:
    signed_payload = dict(payload)
    private_key = load_private_key(private_key_path)
    signature = private_key.sign(canonical_license_bytes(signed_payload))
    signed_payload["signature"] = base64.b64encode(signature).decode("ascii")
    return signed_payload


def write_license_file(payload: dict[str, Any], out_path: str) -> str:
    output_path = Path(out_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)
    return str(output_path)


def create_signed_license(
    *,
    private_key_path: str,
    out_path: str,
    license_id: str,
    license_type: str,
    customer_name: str,
    customer_id: str,
    issue_date: str,
    expiry_date: str,
    binding_mode: str,
    machine_fingerprint: str = "",
    machine_label: str = "",
    note: str = "",
    features: dict[str, Any] | None = None,
) -> dict[str, Any]:
    payload = build_license_payload(
        license_id=license_id,
        license_type=license_type,
        customer_name=customer_name,
        customer_id=customer_id,
        issue_date=issue_date,
        expiry_date=expiry_date,
        binding_mode=binding_mode,
        machine_fingerprint=machine_fingerprint,
        machine_label=machine_label,
        note=note,
        features=features,
    )
    signed_payload = sign_license_payload(payload, private_key_path)
    write_license_file(signed_payload, out_path)
    return signed_payload


def create_signed_trial_license(
    *,
    private_key_path: str,
    out_path: str,
    license_id: str,
    customer_name: str,
    customer_id: str,
    days_valid: int,
    issue_date: str = "",
    note: str = "Trial license",
    features: dict[str, Any] | None = None,
) -> dict[str, Any]:
    payload = build_trial_license_payload(
        license_id=license_id,
        customer_name=customer_name,
        customer_id=customer_id,
        days_valid=days_valid,
        issue_date=issue_date,
        note=note,
        features=features,
    )
    signed_payload = sign_license_payload(payload, private_key_path)
    write_license_file(signed_payload, out_path)
    return signed_payload


def load_machine_identity_json(path: str) -> dict[str, str]:
    with open(path, "r", encoding="utf-8") as handle:
        payload = json.load(handle)
    if not isinstance(payload, dict):
        raise ValueError("machine_fingerprint.json format is invalid.")
    fingerprint = str(payload.get("machine_fingerprint") or "").strip()
    machine_label = str(payload.get("machine_label") or "").strip()
    if not fingerprint:
        raise ValueError("machine_fingerprint.json does not contain machine_fingerprint.")
    return {
        "machine_fingerprint": fingerprint,
        "machine_label": machine_label,
    }
