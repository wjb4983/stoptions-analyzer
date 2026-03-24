from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk


TERMINAL_STATES = {"completed", "failed", "canceled", "succeeded"}


class RemoteJobsPage(ttk.Frame):
    def __init__(self, parent: ttk.Frame, controller: StoptionsApp) -> None:
        super().__init__(parent)
        self.controller = controller
        self._rows: dict[str, dict[str, object]] = {}

        ttk.Label(self, text="Remote Jobs", font=("Arial", 18, "bold")).pack(pady=(14, 8))
        ttk.Label(
            self,
            text="Review active and historical jobs. Reattach existing jobs and refresh local summaries.",
            justify="center",
            wraplength=940,
        ).pack(pady=(0, 10))

        controls = ttk.Frame(self)
        controls.pack(fill="x", padx=30, pady=(0, 6))
        ttk.Button(controls, text="Refresh statuses", command=self.refresh).pack(side="left")
        ttk.Button(controls, text="Reattach selected", command=self.reattach_selected).pack(side="left", padx=(8, 0))
        ttk.Button(controls, text="Refresh summaries", command=self.refresh_summaries_selected).pack(side="left", padx=(8, 0))
        ttk.Button(
            controls,
            text="Back to Main Menu",
            command=lambda: self.controller.show_frame("MainMenu"),
        ).pack(side="right")

        table_frame = ttk.Frame(self)
        table_frame.pack(fill="both", expand=True, padx=30, pady=(0, 16))
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        columns = ("job_type", "last_known_state", "submitted_at", "server_host", "summary_cache_path")
        self.jobs_tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="browse")
        self.jobs_tree.grid(row=0, column=0, sticky="nsew")
        headers = {
            "job_type": "Job Type",
            "last_known_state": "State",
            "submitted_at": "Submitted",
            "server_host": "Server Host",
            "summary_cache_path": "Summary Cache Path",
        }
        widths = {
            "job_type": 220,
            "last_known_state": 120,
            "submitted_at": 220,
            "server_host": 220,
            "summary_cache_path": 430,
        }
        for col in columns:
            self.jobs_tree.heading(col, text=headers[col])
            self.jobs_tree.column(col, width=widths[col], anchor="w")

        scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.jobs_tree.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.jobs_tree.configure(yscrollcommand=scroll.set)

        self.status_var = tk.StringVar(value="No remote jobs tracked.")
        ttk.Label(self, textvariable=self.status_var, justify="left").pack(fill="x", padx=30, pady=(0, 12))

    def refresh(self) -> None:
        for row in self.controller.job_manager.list_remote_jobs():
            self.controller.job_manager.refresh_job_status(str(row.get("job_id", "")))
        rows = self.controller.job_manager.list_remote_jobs()
        self._rows = {str(row.get("job_id", "")): row for row in rows}
        self.jobs_tree.delete(*self.jobs_tree.get_children())
        for row in rows:
            job_id = str(row.get("job_id", "")).strip()
            if not job_id:
                continue
            self.jobs_tree.insert(
                "",
                "end",
                iid=job_id,
                values=(
                    row.get("job_type", "unknown"),
                    row.get("last_known_state", "queued"),
                    row.get("submitted_at", "-") or "-",
                    row.get("server_host", "unknown"),
                    row.get("summary_cache_path", "-") or "-",
                ),
            )
        active = sum(
            1
            for row in rows
            if str(row.get("last_known_state", "")).strip().lower() not in TERMINAL_STATES
        )
        self.status_var.set(f"Tracked jobs: {len(rows)} (active: {active}, historical: {max(0, len(rows) - active)}).")

    def _selected_job_id(self) -> str | None:
        selected = self.jobs_tree.selection()
        if not selected:
            return None
        return str(selected[0]).strip() or None

    def reattach_selected(self) -> None:
        job_id = self._selected_job_id()
        if not job_id:
            messagebox.showinfo("Remote Jobs", "Select a job first.")
            return
        if self.controller.job_manager.reattach_job(job_id):
            self.refresh()
            messagebox.showinfo("Remote Jobs", f"Reattached {job_id}.")
            return
        messagebox.showwarning("Remote Jobs", f"Unable to reattach {job_id}.")

    def refresh_summaries_selected(self) -> None:
        job_id = self._selected_job_id()
        if not job_id:
            messagebox.showinfo("Remote Jobs", "Select a job first.")
            return
        summary_path = self.controller.job_manager.refresh_job_summary(job_id)
        if summary_path:
            self.refresh()
            messagebox.showinfo("Remote Jobs", f"Summary refreshed for {job_id}.\nCached at: {summary_path}")
            return
        messagebox.showwarning(
            "Remote Jobs",
            "Summary refresh is available only for completed/canceled/failed jobs with remote access.",
        )
