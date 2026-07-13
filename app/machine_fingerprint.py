"""
Machine identity helpers for offline license binding.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import socket
import sys
from datetime import datetime
from typing import Any

if sys.platform == "win32":
    import winreg
else:  # pragma: no cover - tool is intended for Windows deployment
    winreg = None


MACHINE_GUID_REG_PATH = r"SOFTWARE\Microsoft\Cryptography"
MACHINE_GUID_REG_VALUE = "MachineGuid"


def get_machine_guid() -> str:
    if sys.platform != "win32" or winreg is None:
        raise RuntimeError("Machine fingerprint is only supported on Windows.")

    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, MACHINE_GUID_REG_PATH) as key:
            value, _value_type = winreg.QueryValueEx(key, MACHINE_GUID_REG_VALUE)
    except OSError as exc:  # pragma: no cover - depends on local registry state
        raise RuntimeError(
            "Could not read Windows MachineGuid. Run the tool with permission to read "
            r"HKLM\SOFTWARE\Microsoft\Cryptography\MachineGuid."
        ) from exc

    machine_guid = str(value or "").strip().lower()
    if not machine_guid:
        raise RuntimeError("Windows MachineGuid is empty.")
    return machine_guid


def get_machine_fingerprint() -> str:
    machine_guid = get_machine_guid()
    digest = hashlib.sha256(machine_guid.encode("utf-8")).hexdigest()
    return f"sha256:{digest}"


def get_machine_label() -> str:
    return os.environ.get("COMPUTERNAME") or socket.gethostname() or "UNKNOWN_MACHINE"


def sanitize_machine_label(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", str(value or "").strip())
    return cleaned or "UNKNOWN_MACHINE"


def build_machine_identity_payload() -> dict[str, Any]:
    return {
        "product_name": "Fortune BJ Optimizer",
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "machine_label": get_machine_label(),
        "machine_fingerprint": get_machine_fingerprint(),
    }


def _main() -> int:
    parser = argparse.ArgumentParser(description="Print or export the local machine fingerprint.")
    parser.add_argument(
        "--out",
        help="Optional JSON output path. If omitted, prints the machine fingerprint to the console.",
    )
    parser.add_argument(
        "--out-dir",
        help="Optional output directory. Writes a timestamped machine_fingerprint_*.json file there.",
    )
    args = parser.parse_args()

    payload = build_machine_identity_payload()
    if args.out and args.out_dir:
        raise SystemExit("Use either --out or --out-dir, not both.")
    if args.out or args.out_dir:
        output_path = args.out
        if args.out_dir:
            os.makedirs(args.out_dir, exist_ok=True)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"machine_fingerprint_{sanitize_machine_label(payload['machine_label'])}_{timestamp}.json"
            output_path = os.path.join(args.out_dir, filename)
        with open(output_path, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, ensure_ascii=False, indent=2)
        print(f"Machine fingerprint written to: {output_path}")
    else:
        print("Fortune BJ Optimizer - Machine Fingerprint")
        print("------------------------------------------------")
        print(f"Machine label      : {payload['machine_label']}")
        print(f"Machine fingerprint: {payload['machine_fingerprint']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(_main())
