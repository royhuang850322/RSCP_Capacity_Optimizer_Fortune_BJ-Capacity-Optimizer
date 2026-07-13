"""
Internal helper to generate a short-term unbound trial license.

This script is intended for RSCP internal use only. It creates a signed
`license.json` that does not bind to a specific machine, which makes it
convenient for time-boxed evaluations and demos.
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
    create_signed_trial_license,
    default_license_admin_root,
)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Generate a signed unbound trial license for the Fortune BJ Optimizer."
    )
    parser.add_argument("--private-key", required=True, help="Path to the Ed25519 private key PEM.")
    parser.add_argument("--out", default="", help="Optional explicit output path. If omitted, use the managed issued-folder path.")
    parser.add_argument("--license-id", required=True, help="Unique license ID.")
    parser.add_argument("--customer-name", required=True, help="Licensed customer name.")
    parser.add_argument("--customer-id", required=True, help="Licensed customer ID.")
    parser.add_argument("--days-valid", type=int, default=30, help="Number of valid calendar days.")
    parser.add_argument("--issue-date", default="", help="Optional issue date in YYYY-MM-DD format.")
    parser.add_argument("--note", default="Trial license", help="Optional note stored in the license file.")
    parser.add_argument("--admin-root", default=str(default_license_admin_root()), help="Managed license repository root.")
    parser.add_argument("--tool-name", default=DEFAULT_TOOL_REPOSITORY_NAME, help="Tool folder name inside the managed license repository.")
    parser.add_argument("--no-activate", action="store_true", help="Do not copy the generated license into the active folder.")
    args = parser.parse_args()

    if args.days_valid <= 0:
        raise SystemExit("--days-valid must be greater than zero")

    out_path = args.out.strip() or str(
        build_issued_license_path(
            args.customer_name,
            args.license_id,
            admin_root=args.admin_root,
            tool_name=args.tool_name,
        )
    )

    payload = create_signed_trial_license(
        private_key_path=args.private_key,
        out_path=out_path,
        license_id=args.license_id,
        customer_name=args.customer_name,
        customer_id=args.customer_id,
        days_valid=args.days_valid,
        issue_date=args.issue_date,
        note=args.note,
    )
    print(f"Trial license file written to: {Path(out_path)}")
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
