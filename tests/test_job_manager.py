from __future__ import annotations

from dataclasses import dataclass

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


@dataclass
class FakeController:
    state: AppState
    execution_backend: FlakyBackend

    def __post_init__(self) -> None:
        self.persist_count = 0

    def persist_state(self) -> None:
        self.persist_count += 1


def test_job_manager_retries_transient_status_errors_and_persists_metadata() -> None:
    controller = FakeController(state=AppState(), execution_backend=FlakyBackend())
    manager = JobManager(controller=controller, poll_interval_seconds=0.01, max_retries=2)

    result = manager.run_job_and_wait(job_type="backtesting.multi_signal", payload={}, source_page="test")

    assert result.status == "succeeded"
    assert result.result == {"ok": True}
    assert "job-1" in controller.state.active_jobs
    assert controller.state.active_jobs["job-1"]["transport_retries"] == 1


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
