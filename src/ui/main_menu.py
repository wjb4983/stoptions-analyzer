from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk

from config import API_KEY_PATH, DEFAULT_REMOTE_EXECUTION_SETTINGS, validate_remote_execution_settings
from execution.remote_ssh_backend import build_remote_backend_from_settings
from ui.helpers import load_remote_secrets, save_api_key, save_remote_secrets


class MainMenu(ttk.Frame):
    def __init__(self, parent: ttk.Frame, controller: StoptionsApp) -> None:
        super().__init__(parent)
        self.controller = controller

        title = ttk.Label(self, text="Stoptions Analyzer", font=("Arial", 24, "bold"))
        title.pack(pady=20)

        description = ttk.Label(
            self,
            text="Manage tickers, select a stock, and explore option strategy analysis.",
            wraplength=600,
            justify="center",
        )
        description.pack(pady=10)

        api_frame = ttk.LabelFrame(self, text="Massive API Key")
        api_frame.pack(pady=15, padx=40, fill="x")
        api_frame.columnconfigure(1, weight=1)

        ttk.Label(api_frame, text="API Key").grid(row=0, column=0, padx=10, pady=8, sticky="w")
        self.api_key_var = tk.StringVar(value=self.controller.api_key)
        self.api_key_entry = ttk.Entry(api_frame, textvariable=self.api_key_var, show="*")
        self.api_key_entry.grid(row=0, column=1, padx=10, pady=8, sticky="ew")
        ttk.Button(api_frame, text="Save Key", command=self.save_api_key).grid(
            row=0, column=2, padx=10, pady=8
        )

        self._build_remote_backend_section()

        button_frame = ttk.Frame(self)
        button_frame.pack(pady=40)

        ttk.Button(
            button_frame,
            text="Enter Stock Tickers",
            command=lambda: controller.show_frame("TickerEntryPage"),
            width=30,
        ).grid(row=0, column=0, pady=10)

        ttk.Button(
            button_frame,
            text="Select Stock",
            command=lambda: controller.show_frame("TickerSelectPage"),
            width=30,
        ).grid(row=1, column=0, pady=10)

        ttk.Button(
            button_frame,
            text="Analysis",
            command=lambda: controller.show_frame("AnalysisPage"),
            width=30,
        ).grid(row=2, column=0, pady=10)

        ttk.Button(
            button_frame,
            text="General Analysis",
            command=lambda: controller.show_frame("GeneralAnalysisPage"),
            width=30,
        ).grid(row=3, column=0, pady=10)


        ttk.Button(
            button_frame,
            text="Intraday Replay",
            command=lambda: controller.show_frame("IntradayReplayPage"),
            width=30,
        ).grid(row=4, column=0, pady=10)

        ttk.Button(
            button_frame,
            text="Backtesting",
            command=lambda: controller.show_frame("BacktestingPage"),
            width=30,
        ).grid(row=5, column=0, pady=10)

        ttk.Button(
            button_frame,
            text="Research Lab",
            command=lambda: controller.show_frame("ResearchLabPage"),
            width=30,
        ).grid(row=6, column=0, pady=10)

        ttk.Button(
            button_frame,
            text="Create Regime",
            command=self.open_create_regime_workspace,
            width=30,
        ).grid(row=7, column=0, pady=10)

    def open_create_regime_workspace(self) -> None:
        self.controller.show_frame("CreateRegimePage")

    def refresh(self) -> None:
        self.api_key_var.set(self.controller.api_key)
        self._load_remote_settings_into_form()

    def save_api_key(self) -> None:
        key = self.api_key_var.get().strip()
        if not key:
            messagebox.showinfo("Missing key", "Enter a Massive API key first.")
            return
        save_api_key(key)
        self.controller.api_key = key
        messagebox.showinfo(
            "Saved", f"API key saved to {API_KEY_PATH} (not tracked in git)."
        )

    def _build_remote_backend_section(self) -> None:
        remote_frame = ttk.LabelFrame(self, text="Remote Backend")
        remote_frame.pack(pady=12, padx=40, fill="x")
        remote_frame.columnconfigure(1, weight=1)
        remote_frame.columnconfigure(3, weight=1)

        settings = dict(DEFAULT_REMOTE_EXECUTION_SETTINGS)
        settings.update(getattr(self.controller.state, "remote_execution_settings", {}))
        secrets = load_remote_secrets()

        self.remote_mode_var = tk.StringVar(value=str(settings.get("mode", "local")))
        self.remote_host_var = tk.StringVar(value=str(settings.get("ssh_host", "")))
        self.remote_port_var = tk.StringVar(value=str(settings.get("ssh_port", "22")))
        self.remote_user_var = tk.StringVar(value=str(settings.get("ssh_user", "")))
        self.remote_project_path_var = tk.StringVar(value=str(settings.get("remote_project_path", "~/stoptions_jobs")))
        self.remote_python_command_var = tk.StringVar(value=str(settings.get("remote_python_command", "python")))
        self.remote_venv_path_var = tk.StringVar(value=str(settings.get("remote_venv_path", "")))
        self.remote_scheduler_enabled_var = tk.BooleanVar(value=bool(settings.get("scheduler_enabled", False)))
        self.remote_scheduler_name_var = tk.StringVar(value=str(settings.get("scheduler_name", "")))
        self.remote_scheduler_queue_var = tk.StringVar(value=str(settings.get("scheduler_queue", "")))
        self.remote_scheduler_max_jobs_var = tk.StringVar(value=str(settings.get("scheduler_max_concurrent_jobs", "1")))
        self.remote_scheduler_poll_var = tk.StringVar(value=str(settings.get("scheduler_poll_seconds", "1.5")))
        self.remote_api_policy_var = tk.StringVar(value=str(settings.get("api_policy", "server_managed")))
        self.remote_server_api_key_file_var = tk.StringVar(value=str(settings.get("server_api_key_file", "")))
        self.remote_ssh_options_var = tk.StringVar(value=str(secrets.get("ssh_options", "")))
        self.remote_ssh_identity_var = tk.StringVar(value=str(secrets.get("ssh_identity_file", "")))

        row = 0
        ttk.Label(remote_frame, text="Mode").grid(row=row, column=0, sticky="w", padx=8, pady=4)
        ttk.Combobox(remote_frame, textvariable=self.remote_mode_var, values=["local", "remote"], state="readonly", width=16).grid(row=row, column=1, sticky="w", padx=8, pady=4)
        ttk.Label(remote_frame, text="API policy").grid(row=row, column=2, sticky="w", padx=8, pady=4)
        ttk.Combobox(
            remote_frame,
            textvariable=self.remote_api_policy_var,
            values=["server_managed", "forward_from_client"],
            state="readonly",
            width=20,
        ).grid(row=row, column=3, sticky="w", padx=8, pady=4)

        row += 1
        ttk.Label(remote_frame, text="Security").grid(row=row, column=0, sticky="nw", padx=8, pady=4)
        ttk.Label(
            remote_frame,
            text=(
                "server_managed (recommended): remote worker reads MASSIVE_API_KEY from server env "
                "or from a secure file path outside this repo.\n"
                "forward_from_client: forwards your local key only at process launch (ephemeral); "
                "the key is never written to job files or logs."
            ),
            wraplength=760,
            justify="left",
        ).grid(row=row, column=1, columnspan=3, sticky="w", padx=8, pady=4)

        row += 1
        ttk.Label(remote_frame, text="SSH host").grid(row=row, column=0, sticky="w", padx=8, pady=4)
        ttk.Entry(remote_frame, textvariable=self.remote_host_var).grid(row=row, column=1, sticky="ew", padx=8, pady=4)
        ttk.Label(remote_frame, text="SSH port").grid(row=row, column=2, sticky="w", padx=8, pady=4)
        ttk.Entry(remote_frame, textvariable=self.remote_port_var).grid(row=row, column=3, sticky="ew", padx=8, pady=4)

        row += 1
        ttk.Label(remote_frame, text="SSH user").grid(row=row, column=0, sticky="w", padx=8, pady=4)
        ttk.Entry(remote_frame, textvariable=self.remote_user_var).grid(row=row, column=1, sticky="ew", padx=8, pady=4)
        ttk.Label(remote_frame, text="Remote root").grid(row=row, column=2, sticky="w", padx=8, pady=4)
        ttk.Entry(remote_frame, textvariable=self.remote_project_path_var).grid(row=row, column=3, sticky="ew", padx=8, pady=4)

        row += 1
        ttk.Label(remote_frame, text="Python command").grid(row=row, column=0, sticky="w", padx=8, pady=4)
        ttk.Entry(remote_frame, textvariable=self.remote_python_command_var).grid(row=row, column=1, sticky="ew", padx=8, pady=4)
        ttk.Label(remote_frame, text="Virtualenv path").grid(row=row, column=2, sticky="w", padx=8, pady=4)
        ttk.Entry(remote_frame, textvariable=self.remote_venv_path_var).grid(row=row, column=3, sticky="ew", padx=8, pady=4)

        row += 1
        ttk.Label(remote_frame, text="Server key file (optional)").grid(row=row, column=0, sticky="w", padx=8, pady=4)
        ttk.Entry(remote_frame, textvariable=self.remote_server_api_key_file_var).grid(row=row, column=1, sticky="ew", padx=8, pady=4)
        ttk.Label(remote_frame, text="").grid(row=row, column=2, sticky="w", padx=8, pady=4)
        ttk.Label(remote_frame, text="").grid(row=row, column=3, sticky="w", padx=8, pady=4)

        row += 1
        ttk.Checkbutton(remote_frame, text="Use scheduler", variable=self.remote_scheduler_enabled_var).grid(row=row, column=0, sticky="w", padx=8, pady=4)
        ttk.Label(remote_frame, text="Scheduler name").grid(row=row, column=2, sticky="w", padx=8, pady=4)
        ttk.Entry(remote_frame, textvariable=self.remote_scheduler_name_var).grid(row=row, column=3, sticky="ew", padx=8, pady=4)

        row += 1
        ttk.Label(remote_frame, text="Scheduler queue").grid(row=row, column=0, sticky="w", padx=8, pady=4)
        ttk.Entry(remote_frame, textvariable=self.remote_scheduler_queue_var).grid(row=row, column=1, sticky="ew", padx=8, pady=4)
        ttk.Label(remote_frame, text="Max concurrent jobs").grid(row=row, column=2, sticky="w", padx=8, pady=4)
        ttk.Entry(remote_frame, textvariable=self.remote_scheduler_max_jobs_var).grid(row=row, column=3, sticky="ew", padx=8, pady=4)

        row += 1
        ttk.Label(remote_frame, text="Poll seconds").grid(row=row, column=0, sticky="w", padx=8, pady=4)
        ttk.Entry(remote_frame, textvariable=self.remote_scheduler_poll_var).grid(row=row, column=1, sticky="ew", padx=8, pady=4)
        ttk.Label(remote_frame, text="SSH options (secret)").grid(row=row, column=2, sticky="w", padx=8, pady=4)
        ttk.Entry(remote_frame, textvariable=self.remote_ssh_options_var).grid(row=row, column=3, sticky="ew", padx=8, pady=4)

        row += 1
        ttk.Label(remote_frame, text="SSH identity file (secret)").grid(row=row, column=0, sticky="w", padx=8, pady=4)
        ttk.Entry(remote_frame, textvariable=self.remote_ssh_identity_var).grid(row=row, column=1, sticky="ew", padx=8, pady=4)
        action_row = ttk.Frame(remote_frame)
        action_row.grid(row=row, column=3, sticky="e", padx=8, pady=4)
        ttk.Button(action_row, text="Validate connection", command=self.validate_remote_connection).pack(side="left", padx=(0, 8))
        ttk.Button(action_row, text="Save remote settings", command=self.save_remote_settings).pack(side="left")

    def _collect_remote_settings(self) -> dict[str, object]:
        return {
            "mode": self.remote_mode_var.get().strip().lower() or "local",
            "ssh_host": self.remote_host_var.get().strip(),
            "ssh_port": self.remote_port_var.get().strip() or "22",
            "ssh_user": self.remote_user_var.get().strip(),
            "remote_project_path": self.remote_project_path_var.get().strip(),
            "remote_python_command": self.remote_python_command_var.get().strip(),
            "remote_venv_path": self.remote_venv_path_var.get().strip(),
            "scheduler_enabled": bool(self.remote_scheduler_enabled_var.get()),
            "scheduler_name": self.remote_scheduler_name_var.get().strip(),
            "scheduler_queue": self.remote_scheduler_queue_var.get().strip(),
            "scheduler_max_concurrent_jobs": self.remote_scheduler_max_jobs_var.get().strip() or "1",
            "scheduler_poll_seconds": self.remote_scheduler_poll_var.get().strip() or "1.5",
            "api_policy": self.remote_api_policy_var.get().strip().lower() or "server_managed",
            "server_api_key_file": self.remote_server_api_key_file_var.get().strip(),
        }

    def _load_remote_settings_into_form(self) -> None:
        settings = dict(DEFAULT_REMOTE_EXECUTION_SETTINGS)
        settings.update(getattr(self.controller.state, "remote_execution_settings", {}))
        self.remote_mode_var.set(str(settings.get("mode", "local")))
        self.remote_host_var.set(str(settings.get("ssh_host", "")))
        self.remote_port_var.set(str(settings.get("ssh_port", "22")))
        self.remote_user_var.set(str(settings.get("ssh_user", "")))
        self.remote_project_path_var.set(str(settings.get("remote_project_path", "~/stoptions_jobs")))
        self.remote_python_command_var.set(str(settings.get("remote_python_command", "python")))
        self.remote_venv_path_var.set(str(settings.get("remote_venv_path", "")))
        self.remote_scheduler_enabled_var.set(bool(settings.get("scheduler_enabled", False)))
        self.remote_scheduler_name_var.set(str(settings.get("scheduler_name", "")))
        self.remote_scheduler_queue_var.set(str(settings.get("scheduler_queue", "")))
        self.remote_scheduler_max_jobs_var.set(str(settings.get("scheduler_max_concurrent_jobs", "1")))
        self.remote_scheduler_poll_var.set(str(settings.get("scheduler_poll_seconds", "1.5")))
        self.remote_api_policy_var.set(str(settings.get("api_policy", "server_managed")))
        self.remote_server_api_key_file_var.set(str(settings.get("server_api_key_file", "")))

    def save_remote_settings(self) -> None:
        settings = self._collect_remote_settings()
        errors = validate_remote_execution_settings(settings)
        if errors:
            messagebox.showerror("Remote settings invalid", "\n• " + "\n• ".join(errors))
            return
        self.controller.state.remote_execution_settings = settings
        save_remote_secrets(
            {
                "ssh_options": self.remote_ssh_options_var.get().strip(),
                "ssh_identity_file": self.remote_ssh_identity_var.get().strip(),
            }
        )
        self.controller.configure_execution_backend()
        self.controller.persist_state()
        messagebox.showinfo("Saved", "Remote backend settings saved.")

    def validate_remote_connection(self) -> None:
        settings = self._collect_remote_settings()
        errors = validate_remote_execution_settings(settings)
        if errors:
            messagebox.showerror("Remote settings invalid", "\n• " + "\n• ".join(errors))
            return
        if str(settings.get("mode", "local")).strip().lower() != "remote":
            messagebox.showinfo("Validation", "Mode is local; no remote connection check needed.")
            return
        try:
            backend = build_remote_backend_from_settings(settings)
            ok, detail = backend.validate_connection()
        except Exception as exc:
            messagebox.showerror("Connection failed", str(exc))
            return
        if ok:
            messagebox.showinfo("Connection successful", detail)
        else:
            messagebox.showerror("Connection failed", detail)
