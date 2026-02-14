from __future__ import annotations

import threading
import time
from types import SimpleNamespace

from src.ui.research_lab_page import ResearchLabPage, ResearchTask


class _Var:
    def __init__(self, value: str) -> None:
        self._value = value

    def get(self) -> str:
        return self._value

    def set(self, value: str) -> None:
        self._value = value


def _build_page(max_concurrent_jobs: int = 1) -> ResearchLabPage:
    page = object.__new__(ResearchLabPage)
    page._task_queue = []
    page._active_task_ids = set()
    page._task_enqueue_counter = 0
    page._max_concurrent_jobs = max_concurrent_jobs
    page._append_output = lambda *_args, **_kwargs: None
    page._refresh_task_queue_ui = lambda: None
    page._refresh_selected_task_logs = lambda: None
    page._wizard_refresh_nav_state = lambda: None
    page._refresh_governance_dashboard_from_output = lambda *_args, **_kwargs: None
    page._append_explainability_cards = lambda *_args, **_kwargs: None
    page._emit_research_pack = lambda *_args, **_kwargs: None
    page.after = lambda _delay, callback: callback()
    return page


def test_scheduler_respects_priority_with_single_worker() -> None:
    page = _build_page(max_concurrent_jobs=1)
    release_first = threading.Event()
    started: list[str] = []

    def runner(context, _config, _token):
        label = str(context["label"])
        started.append(label)
        if label == "low":
            release_first.wait(timeout=2)
        return f"done:{label}"

    cfg = SimpleNamespace()

    page._task_priority_var = _Var("low")
    page._enqueue_task(label="low", target=runner, context={"label": "low", "cache_root": "/tmp"}, config=cfg)

    page._task_priority_var.set("high")
    page._enqueue_task(label="high", target=runner, context={"label": "high", "cache_root": "/tmp"}, config=cfg)

    page._task_priority_var.set("normal")
    page._enqueue_task(label="normal", target=runner, context={"label": "normal", "cache_root": "/tmp"}, config=cfg)

    time.sleep(0.05)
    assert started == ["low"]
    assert [task.state for task in page._task_queue] == ["running", "queued", "queued"]

    release_first.set()
    deadline = time.time() + 2
    while time.time() < deadline and len(started) < 3:
        time.sleep(0.01)

    assert started == ["low", "high", "normal"]

    deadline = time.time() + 2
    while time.time() < deadline and any(task.state != "succeeded" for task in page._task_queue):
        time.sleep(0.01)
    assert all(task.state == "succeeded" for task in page._task_queue)


def test_cancel_running_task_transitions_canceling_to_canceled() -> None:
    page = _build_page(max_concurrent_jobs=2)
    task = ResearchTask(
        task_id="t1",
        label="run",
        target=lambda _c, _cfg, _token: "ok",
        context={},
        config=SimpleNamespace(),
        state="running",
    )
    page._task_queue = [task]
    page._selected_task = lambda: task

    page._cancel_selected_task()

    assert task.state == "canceling"
    assert task.cancel_requested is True
    assert task.cancellation_reason == "Canceled from Research Lab UI"

    page._finish_task(task.task_id, "", canceled=True)
    assert task.state == "canceled"
    assert task.cancellation_confirmed is True


def test_retry_transition_sets_retrying_and_requeues() -> None:
    page = _build_page(max_concurrent_jobs=1)
    calls = {"schedule": 0}
    page._schedule_tasks = lambda: calls.__setitem__("schedule", calls["schedule"] + 1)

    original_token_task = ResearchTask(
        task_id="t2",
        label="retry",
        target=lambda _c, _cfg, _token: "ok",
        context={},
        config=SimpleNamespace(),
        state="failed",
    )
    old_token = original_token_task.cancellation_token
    page._task_queue = [original_token_task]
    page._selected_task = lambda: original_token_task

    page._retry_selected_task()

    assert original_token_task.state == "queued"
    assert any("Retry requested." in line for line in original_token_task.logs)
    assert any("Task re-queued for retry." in line for line in original_token_task.logs)
    assert original_token_task.cancellation_token is not old_token
    assert calls["schedule"] == 1
