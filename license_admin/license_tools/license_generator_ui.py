"""
Internal Tkinter GUI for generating signed license files without typing
PowerShell commands.
"""
from __future__ import annotations

import os
import sys
import tkinter as tk
from datetime import date, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

BOOTSTRAP_ROOT = Path(__file__).resolve().parents[2]
if str(BOOTSTRAP_ROOT) not in sys.path:
    sys.path.insert(0, str(BOOTSTRAP_ROOT))

PROJECT_ROOT = BOOTSTRAP_ROOT
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from license_admin.license_tools.common import (
    DEFAULT_TOOL_REPOSITORY_NAME,
    activate_issued_license,
    build_active_license_path,
    build_issued_license_path,
    copy_machine_request_to_admin,
    create_signed_license,
    create_signed_trial_license,
    default_license_admin_root,
    ensure_customer_tool_dirs,
    generate_default_license_id,
    load_machine_identity_json,
    parse_iso_date,
    sanitize_path_component,
)


def default_private_key_path() -> Path:
    local_key = PROJECT_ROOT / "license_admin" / "private_keys" / "license_signing_ed25519_private.pem"
    if local_key.exists():
        return local_key
    shared_key = Path(r"C:\Users\super\capacity_optimizer\license_admin\private_keys\license_signing_ed25519_private.pem")
    if shared_key.exists():
        return shared_key
    return local_key


class LicenseGeneratorApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Fortune BJ Optimizer - License Generator")
        self.root.geometry("860x680")
        self.root.minsize(780, 560)

        default_private_key = default_private_key_path()

        self.profile_var = tk.StringVar(value="trial")
        self.admin_root_var = tk.StringVar(value=str(default_license_admin_root()))
        self.private_key_var = tk.StringVar(value=str(default_private_key))
        self.customer_name_var = tk.StringVar()
        self.customer_id_var = tk.StringVar()
        self.tool_name_var = tk.StringVar(value=DEFAULT_TOOL_REPOSITORY_NAME)
        self.license_id_var = tk.StringVar(value=generate_default_license_id("LIC-TRIAL"))
        self.license_type_var = tk.StringVar(value="trial")
        self.issue_date_var = tk.StringVar(value=date.today().isoformat())
        self.expiry_date_var = tk.StringVar(value=date.today().isoformat())
        self.days_valid_var = tk.StringVar(value="14")
        self.binding_mode_var = tk.StringVar(value="unbound")
        self.machine_fingerprint_var = tk.StringVar()
        self.machine_label_var = tk.StringVar()
        self.note_var = tk.StringVar(value="Trial license")

        self.customer_root_var = tk.StringVar()
        self.issued_path_var = tk.StringVar()
        self.active_path_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Fill in the form and click Generate License.")

        self._build_ui()
        self._bind_events()
        self._apply_profile_defaults(reset_license_id=False)
        self._refresh_repository_paths()

    def _build_ui(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        outer = ttk.Frame(self.root, padding=8)
        outer.grid(row=0, column=0, sticky="nsew")
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(0, weight=1)

        self.canvas = tk.Canvas(outer, highlightthickness=0)
        self.canvas.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=self.canvas.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        container = ttk.Frame(self.canvas, padding=12)
        container.columnconfigure(0, weight=1)
        self.canvas_window = self.canvas.create_window((0, 0), window=container, anchor="nw")

        container.bind("<Configure>", self._on_container_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self._bind_mousewheel()

        ttk.Label(
            container,
            text="Fortune BJ Optimizer - Internal License Generator",
            font=("Segoe UI", 15, "bold"),
        ).grid(row=0, column=0, sticky="w")

        ttk.Label(
            container,
            text=(
                "This tool stores license files under RSCP's internal repository using the pattern "
                "AdminRoot / Customer / Tool / requests|issued|active|archive."
            ),
            wraplength=760,
        ).grid(row=1, column=0, sticky="w", pady=(4, 10))

        profile_frame = ttk.LabelFrame(container, text="License Profile", padding=10)
        profile_frame.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        profile_frame.columnconfigure(0, weight=1)
        profile_frame.columnconfigure(1, weight=1)

        ttk.Radiobutton(profile_frame, text="Trial / Unbound", variable=self.profile_var, value="trial").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Radiobutton(profile_frame, text="Manual / Custom", variable=self.profile_var, value="custom").grid(
            row=0, column=1, sticky="w"
        )

        repo_frame = ttk.LabelFrame(container, text="Repository", padding=10)
        repo_frame.grid(row=3, column=0, sticky="ew", pady=(0, 10))
        repo_frame.columnconfigure(1, weight=1)

        ttk.Label(repo_frame, text="Admin Root").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(repo_frame, textvariable=self.admin_root_var).grid(row=0, column=1, sticky="ew", pady=4)
        ttk.Button(repo_frame, text="Browse...", command=self._browse_admin_root).grid(row=0, column=2, padx=(8, 0), pady=4)

        ttk.Label(repo_frame, text="Private Key").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(repo_frame, textvariable=self.private_key_var).grid(row=1, column=1, sticky="ew", pady=4)
        ttk.Button(repo_frame, text="Browse...", command=self._browse_private_key).grid(row=1, column=2, padx=(8, 0), pady=4)

        ttk.Label(repo_frame, text="Customer / Tool Root").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(repo_frame, textvariable=self.customer_root_var, state="readonly").grid(row=2, column=1, columnspan=2, sticky="ew", pady=4)

        ttk.Label(repo_frame, text="Issued File").grid(row=3, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(repo_frame, textvariable=self.issued_path_var, state="readonly").grid(row=3, column=1, columnspan=2, sticky="ew", pady=4)

        ttk.Label(repo_frame, text="Active File").grid(row=4, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(repo_frame, textvariable=self.active_path_var, state="readonly").grid(row=4, column=1, columnspan=2, sticky="ew", pady=4)

        details_frame = ttk.LabelFrame(container, text="License Details", padding=10)
        details_frame.grid(row=4, column=0, sticky="ew", pady=(0, 10))
        for idx in range(4):
            details_frame.columnconfigure(idx, weight=1)

        self._add_labeled_entry(details_frame, "Customer Name", self.customer_name_var, 0, 0)
        self._add_labeled_entry(details_frame, "Customer ID", self.customer_id_var, 0, 2)
        self._add_labeled_entry(details_frame, "Tool Name", self.tool_name_var, 1, 0)
        self._add_labeled_entry(details_frame, "License ID", self.license_id_var, 1, 2)
        self.license_type_entry = self._add_labeled_entry(details_frame, "License Type", self.license_type_var, 2, 0)
        self.issue_date_entry = self._add_labeled_entry(details_frame, "Issue Date", self.issue_date_var, 2, 2)
        self.expiry_date_entry = self._add_labeled_entry(details_frame, "Expiry Date", self.expiry_date_var, 3, 0)
        self.days_valid_entry = self._add_labeled_entry(details_frame, "Days Valid", self.days_valid_var, 3, 2)

        ttk.Label(details_frame, text="Binding Mode").grid(row=8, column=0, sticky="w", pady=(10, 2))
        self.binding_mode_combo = ttk.Combobox(
            details_frame,
            textvariable=self.binding_mode_var,
            values=["unbound", "machine_locked"],
            state="readonly",
        )
        self.binding_mode_combo.grid(row=9, column=0, columnspan=2, sticky="ew", padx=(0, 8))

        ttk.Label(details_frame, text="Note").grid(row=8, column=2, sticky="w", pady=(10, 2))
        ttk.Entry(details_frame, textvariable=self.note_var).grid(row=9, column=2, columnspan=2, sticky="ew")

        machine_frame = ttk.LabelFrame(container, text="Machine Binding", padding=10)
        machine_frame.grid(row=5, column=0, sticky="ew", pady=(0, 10))
        machine_frame.columnconfigure(1, weight=1)

        ttk.Label(machine_frame, text="Machine Fingerprint").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        self.machine_fingerprint_entry = ttk.Entry(machine_frame, textvariable=self.machine_fingerprint_var)
        self.machine_fingerprint_entry.grid(row=0, column=1, sticky="ew", pady=4)

        ttk.Label(machine_frame, text="Machine Label").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        self.machine_label_entry = ttk.Entry(machine_frame, textvariable=self.machine_label_var)
        self.machine_label_entry.grid(row=1, column=1, sticky="ew", pady=4)

        self.load_machine_button = ttk.Button(
            machine_frame,
            text="Load machine_fingerprint.json",
            command=self._load_machine_json,
        )
        self.load_machine_button.grid(row=0, column=2, rowspan=2, padx=(8, 0), sticky="ns")

        actions_frame = ttk.Frame(container)
        actions_frame.grid(row=6, column=0, sticky="ew")

        ttk.Button(actions_frame, text="Generate License", command=self._generate_license).grid(row=0, column=0, sticky="w")
        ttk.Button(actions_frame, text="Open Customer Folder", command=self._open_customer_folder).grid(
            row=0, column=1, padx=(10, 0), sticky="w"
        )
        ttk.Button(actions_frame, text="Clear Form", command=self._clear_form).grid(
            row=0, column=2, padx=(10, 0), sticky="w"
        )

        status_frame = ttk.LabelFrame(container, text="Status", padding=10)
        status_frame.grid(row=7, column=0, sticky="ew", pady=(12, 0))
        status_frame.columnconfigure(0, weight=1)
        ttk.Label(status_frame, textvariable=self.status_var, wraplength=740, foreground="#1f4e79").grid(
            row=0, column=0, sticky="w"
        )

    def _on_container_configure(self, _event=None) -> None:
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event) -> None:
        self.canvas.itemconfigure(self.canvas_window, width=event.width)

    def _bind_mousewheel(self) -> None:
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Button-4>", self._on_mousewheel)
        self.canvas.bind_all("<Button-5>", self._on_mousewheel)

    def _on_mousewheel(self, event) -> None:
        if getattr(event, "delta", 0):
            step = -1 * int(event.delta / 120) if event.delta else 0
            if step:
                self.canvas.yview_scroll(step, "units")
            return
        if getattr(event, "num", None) == 4:
            self.canvas.yview_scroll(-1, "units")
        elif getattr(event, "num", None) == 5:
            self.canvas.yview_scroll(1, "units")

    def _bind_events(self) -> None:
        self.profile_var.trace_add("write", lambda *_args: self._apply_profile_defaults(reset_license_id=True))
        self.binding_mode_var.trace_add("write", lambda *_args: self._apply_binding_state())
        self.issue_date_var.trace_add("write", lambda *_args: self._recalculate_trial_expiry())
        self.days_valid_var.trace_add("write", lambda *_args: self._recalculate_trial_expiry())
        self.admin_root_var.trace_add("write", lambda *_args: self._refresh_repository_paths())
        self.customer_name_var.trace_add("write", lambda *_args: self._refresh_repository_paths())
        self.tool_name_var.trace_add("write", lambda *_args: self._refresh_repository_paths())
        self.license_id_var.trace_add("write", lambda *_args: self._refresh_repository_paths())

    def _add_labeled_entry(
        self,
        parent: ttk.LabelFrame,
        label: str,
        variable: tk.StringVar,
        row: int,
        col: int,
    ) -> ttk.Entry:
        ttk.Label(parent, text=label).grid(row=row * 2, column=col, sticky="w", pady=(6, 2))
        entry = ttk.Entry(parent, textvariable=variable)
        entry.grid(row=row * 2 + 1, column=col, columnspan=2, sticky="ew", padx=(0, 8))
        return entry

    def _browse_admin_root(self) -> None:
        selected = filedialog.askdirectory(
            title="Select RSCP license admin root",
            initialdir=self.admin_root_var.get().strip() or str(PROJECT_ROOT),
        )
        if selected:
            self.admin_root_var.set(selected)

    def _browse_private_key(self) -> None:
        selected = filedialog.askopenfilename(
            title="Select Ed25519 private key",
            filetypes=[("PEM files", "*.pem"), ("All files", "*.*")],
            initialdir=str(PROJECT_ROOT),
        )
        if selected:
            self.private_key_var.set(selected)

    def _refresh_repository_paths(self) -> None:
        customer_name = self.customer_name_var.get().strip()
        tool_name = self.tool_name_var.get().strip() or DEFAULT_TOOL_REPOSITORY_NAME
        admin_root = self.admin_root_var.get().strip() or str(default_license_admin_root())
        license_id = self.license_id_var.get().strip() or "PENDING_LICENSE_ID"
        preview_customer = sanitize_path_component(customer_name, "CUSTOMER_NAME")
        preview_tool = sanitize_path_component(tool_name, DEFAULT_TOOL_REPOSITORY_NAME)
        base = Path(admin_root) / preview_customer / preview_tool
        self.customer_root_var.set(str(base))
        self.issued_path_var.set(str(base / "issued" / f"{sanitize_path_component(license_id, 'LICENSE')}.json"))
        self.active_path_var.set(str(base / "active" / "license.json"))

    def _load_machine_json(self) -> None:
        selected = filedialog.askopenfilename(
            title="Select machine_fingerprint.json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            initialdir=str(PROJECT_ROOT),
        )
        if not selected:
            return
        try:
            payload = load_machine_identity_json(selected)
        except Exception as exc:
            messagebox.showerror("Load failed", str(exc), parent=self.root)
            return

        self.machine_fingerprint_var.set(payload["machine_fingerprint"])
        self.machine_label_var.set(payload["machine_label"])

        customer_name = self.customer_name_var.get().strip()
        if customer_name:
            archived_path = copy_machine_request_to_admin(
                selected,
                customer_name,
                admin_root=self.admin_root_var.get().strip(),
                tool_name=self.tool_name_var.get().strip() or DEFAULT_TOOL_REPOSITORY_NAME,
                machine_label=payload["machine_label"],
            )
            self.status_var.set(
                f"Loaded machine fingerprint and copied request file to {archived_path}"
            )
        else:
            self.status_var.set(
                "Loaded machine fingerprint. Fill Customer Name to start archiving request files."
            )

    def _apply_profile_defaults(self, reset_license_id: bool) -> None:
        is_trial = self.profile_var.get() == "trial"
        if is_trial:
            self.license_type_var.set("trial")
            self.binding_mode_var.set("unbound")
            if not self.note_var.get().strip() or self.note_var.get().strip().lower() == "trial license":
                self.note_var.set("Trial license")
            if reset_license_id or not self.license_id_var.get().strip():
                self.license_id_var.set(generate_default_license_id("LIC-TRIAL"))
        else:
            if self.license_type_var.get().strip().lower() == "trial":
                self.license_type_var.set("commercial")
            if self.binding_mode_var.get() == "unbound":
                self.binding_mode_var.set("machine_locked")
            if reset_license_id or not self.license_id_var.get().strip():
                self.license_id_var.set(generate_default_license_id("LIC-COMM"))
        self._apply_profile_state()
        self._apply_binding_state()
        self._recalculate_trial_expiry()
        self._refresh_repository_paths()

    def _apply_profile_state(self) -> None:
        is_trial = self.profile_var.get() == "trial"
        self.license_type_entry.configure(state="disabled" if is_trial else "normal")
        self.binding_mode_combo.configure(state="disabled" if is_trial else "readonly")
        self.days_valid_entry.configure(state="normal" if is_trial else "disabled")
        self.expiry_date_entry.configure(state="disabled" if is_trial else "normal")

    def _apply_binding_state(self) -> None:
        machine_locked = self.profile_var.get() != "trial" and self.binding_mode_var.get() == "machine_locked"
        state = "normal" if machine_locked else "disabled"
        self.machine_fingerprint_entry.configure(state=state)
        self.machine_label_entry.configure(state=state)
        self.load_machine_button.configure(state=state)
        if not machine_locked:
            self.machine_fingerprint_var.set("")
            self.machine_label_var.set("")

    def _recalculate_trial_expiry(self) -> None:
        if self.profile_var.get() != "trial":
            return
        try:
            issue_value = parse_iso_date(self.issue_date_var.get(), "Issue Date")
            days_valid = int(self.days_valid_var.get().strip())
            if days_valid <= 0:
                raise ValueError
        except Exception:
            self.expiry_date_var.set("")
            return
        self.expiry_date_var.set((issue_value + timedelta(days=days_valid - 1)).isoformat())

    def _clear_form(self) -> None:
        self.customer_name_var.set("")
        self.customer_id_var.set("")
        self.tool_name_var.set(DEFAULT_TOOL_REPOSITORY_NAME)
        self.machine_fingerprint_var.set("")
        self.machine_label_var.set("")
        self.note_var.set("Trial license" if self.profile_var.get() == "trial" else "")
        self.issue_date_var.set(date.today().isoformat())
        self.days_valid_var.set("14")
        self._apply_profile_defaults(reset_license_id=True)
        self.status_var.set("Form cleared.")

    def _open_customer_folder(self) -> None:
        customer_root = Path(self.customer_root_var.get().strip())
        customer_root.mkdir(parents=True, exist_ok=True)
        os.startfile(str(customer_root))

    def _generate_license(self) -> None:
        customer_name = self.customer_name_var.get().strip()
        if not customer_name:
            messagebox.showerror("Missing customer", "Customer Name is required.", parent=self.root)
            return

        private_key_path = self.private_key_var.get().strip()
        if not private_key_path:
            messagebox.showerror("Missing private key", "Private key path is required.", parent=self.root)
            return

        admin_root = self.admin_root_var.get().strip() or str(default_license_admin_root())
        tool_name = self.tool_name_var.get().strip() or DEFAULT_TOOL_REPOSITORY_NAME
        license_id = self.license_id_var.get().strip() or generate_default_license_id("LIC")
        issued_path = build_issued_license_path(customer_name, license_id, admin_root=admin_root, tool_name=tool_name)

        try:
            if self.profile_var.get() == "trial":
                payload = create_signed_trial_license(
                    private_key_path=private_key_path,
                    out_path=str(issued_path),
                    license_id=license_id,
                    customer_name=customer_name,
                    customer_id=self.customer_id_var.get(),
                    days_valid=int(self.days_valid_var.get().strip()),
                    issue_date=self.issue_date_var.get(),
                    note=self.note_var.get(),
                )
            else:
                payload = create_signed_license(
                    private_key_path=private_key_path,
                    out_path=str(issued_path),
                    license_id=license_id,
                    license_type=self.license_type_var.get(),
                    customer_name=customer_name,
                    customer_id=self.customer_id_var.get(),
                    issue_date=self.issue_date_var.get(),
                    expiry_date=self.expiry_date_var.get(),
                    binding_mode=self.binding_mode_var.get(),
                    machine_fingerprint=self.machine_fingerprint_var.get(),
                    machine_label=self.machine_label_var.get(),
                    note=self.note_var.get(),
                )
            active_path = activate_issued_license(
                issued_path,
                customer_name,
                admin_root=admin_root,
                tool_name=tool_name,
            )
        except Exception as exc:
            messagebox.showerror("Generation failed", str(exc), parent=self.root)
            self.status_var.set(f"Generation failed: {exc}")
            return

        self._refresh_repository_paths()
        success_message = (
            f"License generated successfully.\n\n"
            f"Customer folder: {self.customer_root_var.get()}\n"
            f"Issued file: {issued_path}\n"
            f"Active file: {active_path}\n"
            f"License ID: {payload['license_id']}\n"
            f"Type: {payload['license_type']}\n"
            f"Binding: {payload['binding_mode']}\n"
            f"Valid from {payload['issue_date']} to {payload['expiry_date']}"
        )
        self.status_var.set(success_message.replace("\n", " "))
        messagebox.showinfo("License generated", success_message, parent=self.root)


def main() -> int:
    root = tk.Tk()
    try:
        style = ttk.Style(root)
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass
    LicenseGeneratorApp(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
