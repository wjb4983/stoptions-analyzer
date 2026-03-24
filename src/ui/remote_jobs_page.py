from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk


class RemoteJobsPage(ttk.Frame):
    def __init__(self, parent: ttk.Frame, controller: StoptionsApp) -> None:
        super().__init__(parent)
        self.controller = controller

        ttk.Label(self, text="Remote Jobs", font=("Arial", 18, "bold")).pack(pady=(12, 8))
        ttk.Label(
            self,
            text="Track active and historical jobs. Reattach to existing jobs or refresh completed summaries.",
            justify="center",
        ).pack(pady=(0, 8))

        controls = ttk.Frame(self)
        controls.pack(fill="x", padx=20, pady=(0, 8))
        ttk.Button(controls, text="Refresh statuses", command=self._refresh_all_statuses).pack(side="left")
        ttk.Button(controls, text="Reattach", command=self._reattach_selected).pack(side="left", padx=(8, 0))
        ttk.Button(controls, text="Refresh summaries", command=self._refresh_selected_summaries).pack(side="left", padx=(8, 0))
        ttk.Button(controls, text="Back to Main Menu", command=lambda: self.controller.show_frame("MainMenu")).pack(side="right")

        table_frame = ttk.Frame(self)
        table_frame.pack(fill="both", expand=True, padx=20, pady=(0, 16))
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        columns = ("job_id", "job_type", "state", "submitted_at", "server_host", "summary_cache_path")
        self.table = ttk.Treeview(table_frame, columns=columns, show="headings", height=18)
        headings = {
            "job_id": "Job ID",
            "job_type": "Type",
            "state": "Last known state",
            "submitted_at": "Submitted at",
            "server_host": "Server host",
            "summary_cache_path": "Summary cache path",
        }
        for key in columns:
            self.table.heading(key, text=headings[key])
            self.table.column(key, width=160 if key != "summary_cache_path" else 320, anchor="w")

        y_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.table.yview)
        self.table.configure(yscrollcommand=y_scroll.set)
        self.table.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")

    def refresh(self) -> None:
        self._populate_table()

    def _populate_table(self) -> None:
        for item in self.table.get_children():
            self.table.delete(item)
        jobs = getattr(self.controller.state, "remote_jobs", {})
        if not isinstance(jobs, dict):
            return
        sorted_jobs = sorted(
            jobs.values(),
            key=lambda item: str(item.get("submitted_at") or ""),
            reverse=True,
        )
        for payload in sorted_jobs:
            if not isinstance(payload, dict):
                continue
            job_id = str(payload.get("job_id", "")).strip()
            if not job_id:
                continue
            self.table.insert(
                "",
                "end",
                iid=job_id,
                values=(
                    job_id,
                    str(payload.get("job_type", "unknown")),
                    str(payload.get("last_known_state", "unknown")),
                    str(payload.get("submitted_at", "") or ""),
                    str(payload.get("server_host", "unknown")),
                    str(payload.get("summary_cache_path", "") or ""),
                ),
            )

    def _selected_job_id(self) -> str | None:
        selected = self.table.selection()
        if not selected:
            return None
        return str(selected[0])

    def _refresh_all_statuses(self) -> None:
        self.controller.job_manager.rehydrate_remote_jobs()
        self._populate_table()

    def _reattach_selected(self) -> None:
        job_id = self._selected_job_id()
        if not job_id:
            messagebox.showinfo("Remote jobs", "Select a job first.")
            return
        if not self.controller.job_manager.reattach_job(job_id):
            messagebox.showerror("Remote jobs", f"Unable to reattach to {job_id}.")
            return
        self._populate_table()
        messagebox.showinfo("Remote jobs", f"Reattached to {job_id} and refreshed status.")

    def _refresh_selected_summaries(self) -> None:
        job_id = self._selected_job_id()
        if not job_id:
            messagebox.showinfo("Remote jobs", "Select a job first.")
            return
        if not self.controller.job_manager.refresh_job_summary(job_id):
            messagebox.showerror(
                "Remote jobs",
                "Summary refresh failed. Ensure the job is completed and reachable from the backend.",
            )
            return
        self._populate_table()
        messagebox.showinfo("Remote jobs", f"Refreshed summaries for {job_id}.")
