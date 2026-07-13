"""
Internal helper to generate a signed license.json file.

This script requires an Ed25519 private key file that is kept outside the
customer-facing delivery package.
"""
from __future__ import annotations

import argparse
from pathlib import Path
import sys

BOOTSTRAP_ROOT = Path(__file__).resolve().parents[2]
if str(BOOTSTRAP_ROOT) not in sys.path:
    sys.path.insert(0, str(BOOTSTRAP_ROOT))

PROJECT_ROOT = BOOTSTRAP_ROOT
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from license_admin.license_tools.common import (
    DEFAULT_TOOL_REPOSITORY_NAME,
    activate_issued_license,
    build_issued_license_path,
    create_signed_license,
    default_license_admin_root,
)


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate a signed Fortune BJ Optimizer license file.")
    parser.add_argument("--private-key", required=True, help="Path to the Ed25519 private key PEM.")
    parser.add_argument("--out", default="", help="Optional explicit output path. If omitted, use the managed issued-folder path.")
    parser.add_argument("--license-id", required=True, help="Unique license ID.")
    parser.add_argument("--license-type", default="commercial", help="License type, e.g. trial or commercial.")
    parser.add_argument("--customer-name", required=True, help="Licensed customer name.")
    parser.add_argument("--customer-id", required=True, help="Licensed customer ID.")
    parser.add_argument("--issue-date", required=True, help="Issue date in YYYY-MM-DD format.")
    parser.add_argument("--expiry-date", required=True, help="Expiry date in YYYY-MM-DD format.")
    parser.add_argument(
        "--binding-mode",
        choices=["unbound", "machine_locked"],
        default="machine_locked",
        help="License binding mode.",
    )
    parser.add_argument("--machine-fingerprint", default="", help="Required for machine_locked licenses.")
    parser.add_argument("--machine-label", default="", help="Human-readable machine label.")
    parser.add_argument("--note", default="", help="Optional note stored in the license file.")
    parser.add_argument("--admin-root", default=str(default_license_admin_root()), help="Managed license repository root.")
    parser.add_argument("--tool-name", default=DEFAULT_TOOL_REPOSITORY_NAME, help="Tool folder name inside the managed license repository.")
    parser.add_argument("--no-activate", action="store_true", help="Do not copy the generated license into the active folder.")
    args = parser.parse_args()

    if args.binding_mode == "machine_locked" and not args.machine_fingerprint.strip():
        raise SystemExit("--machine-fingerprint is required when --binding-mode=machine_locked")

    out_path = args.out.strip() or str(
        build_issued_license_path(
            args.customer_name,
            args.license_id,
            admin_root=args.admin_root,
            tool_name=args.tool_name,
        )
    )

    payload = create_signed_license(
        private_key_path=args.private_key,
        out_path=out_path,
        license_id=args.license_id,
        license_type=args.license_type,
        customer_name=args.customer_name,
        customer_id=args.customer_id,
        issue_date=args.issue_date,
        expiry_date=args.expiry_date,
        binding_mode=args.binding_mode,
        machine_fingerprint=args.machine_fingerprint,
        machine_label=args.machine_label,
        note=args.note,
    )
    print(f"License file written to: {Path(out_path)}")
    if not args.no_activate:
        active_path = activate_issued_license(
            out_path,
            args.customer_name,
            admin_root=args.admin_root,
            tool_name=args.tool_name,
        )
        print(f"Active license updated: {active_path}")
    print(f"Valid from {payload['issue_date']} to {payload['expiry_date']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
