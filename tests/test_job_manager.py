from __future__ import annotations

from dataclasses import dataclass

from execution.contracts import SubmitJobRequest
from state import AppState
from ui.job_manager import JobManager


class FlakyBackend:
    def __init__(self) -> None:
        self.submit_calls = 0
        self.status_calls = 0

    def submit_job(self, job_type: str, payload: dict[str, object]) -> str:
        self.submit_calls += 1
        return "job-1"

    def get_status(self, job_id: str) -> str:
        self.status_calls += 1
        if self.status_calls == 1:
            raise RuntimeError("ssh transport timed out")
        if self.status_calls == 2:
            return "running"
        return "succeeded"

    def stream_logs(self, job_id: str) -> list[str]:
        return ["ok"]

    def get_result(self, job_id: str) -> object:
        return {"ok": True}


class RecoveryBackend:
    def __init__(self) -> None:
        self.registered: list[str] = []

    def register_existing_job(self, *, job_id: str, job_type: str = "unknown") -> None:
        self.registered.append(job_id)

    def get_status(self, job_id: str) -> str:
        return "running"


class StatusOverrideBackend:
    def register_existing_job(self, *, job_id: str, job_type: str = "unknown") -> None:
        return None

    def get_status(self, job_id: str) -> str:
        return "running"

    def get_result(self, job_id: str) -> object:
        return {"server_run_dir": "/remote/runs/job-1", "summary_files": ["/remote/runs/job-1/summary.json"]}


class CompletedBackend(StatusOverrideBackend):
    def get_status(self, job_id: str) -> str:
        return "completed"


@dataclass
class FakeController:
    state: AppState
    execution_backend: object

    def __post_init__(self) -> None:
        self.persist_count = 0

    def persist_state(self) -> None:
        self.persist_count += 1


def test_job_manager_retries_transient_status_errors_and_persists_metadata() -> None:
    controller = FakeController(state=AppState(), execution_backend=FlakyBackend())
    manager = JobManager(controller=controller, poll_interval_seconds=0.01, max_retries=2)

    result = manager.run_job_and_wait(
        request=SubmitJobRequest(job_type="backtesting.multi_signal", payload={}),
        source_page="test",
    )

    assert result.status == "completed"
    assert result.result == {"ok": True}
    assert "job-1" in controller.state.active_jobs
    assert controller.state.active_jobs["job-1"]["transport_retries"] == 1


def test_recover_active_jobs_repolls_remote_status() -> None:
    state = AppState(
        active_jobs={
            "job-9": {
                "job_id": "job-9",
                "job_type": "backtesting.multi_signal",
                "status": "queued",
            }
        }
    )
    backend = RecoveryBackend()
    controller = FakeController(state=state, execution_backend=backend)

    JobManager(controller=controller)

    assert backend.registered == ["job-9"]
    assert controller.state.active_jobs["job-9"]["status"] == "running"


def test_active_job_state_roundtrip_includes_new_fields(tmp_path, monkeypatch) -> None:
    import state as state_module

    state_path = tmp_path / "app_state.json"
    monkeypatch.setattr(state_module, "STATE_PATH", state_path)

    app_state = AppState(active_jobs={"abc": {"job_type": "backtest", "status": "running", "server_hostname": "h"}})
    app_state.save()
    loaded = AppState.load()

    assert "abc" in loaded.active_jobs
    assert loaded.active_jobs["abc"]["job_type"] == "backtest"
    assert loaded.active_jobs["abc"]["status"] == "running"


def test_rehydrate_remote_jobs_uses_server_canonical_status() -> None:
    controller = FakeController(
        state=AppState(
            remote_jobs={
                "job-1": {
                    "job_id": "job-1",
                    "job_type": "backtesting.multi_signal",
                    "submitted_at": "2026-01-01T00:00:00Z",
                    "last_known_state": "queued",
                    "server_host": "host-1",
                    "summary_cache_path": None,
                }
            }
        ),
        execution_backend=StatusOverrideBackend(),
    )
    manager = JobManager(controller=controller)
    manager.rehydrate_remote_jobs()

    assert controller.state.remote_jobs["job-1"]["last_known_state"] == "running"
    assert controller.state.active_jobs["job-1"]["status"] == "running"


def test_refresh_summary_updates_cache_path_for_completed_job() -> None:
    controller = FakeController(
        state=AppState(
            active_jobs={
                "job-1": {
                    "job_id": "job-1",
                    "job_type": "backtesting.multi_signal",
                    "status": "completed",
                    "submitted_at": "2026-01-01T00:00:00Z",
                    "server_hostname": "host-1",
                }
            }
        ),
        execution_backend=CompletedBackend(),
    )
    manager = JobManager(controller=controller)
    ok = manager.refresh_job_summary("job-1")

    assert ok
    assert controller.state.active_jobs["job-1"]["summary_paths"]["0"].endswith("summary.json")
    assert controller.state.remote_jobs["job-1"]["summary_cache_path"].endswith("summary.json")
