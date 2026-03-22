from datetime import datetime

from ui.intraday_replay_page import (
    EASTERN_TZ,
    IntradayReplayPage,
    _group_bars_by_day,
    _minute_of_day,
)


class Var:
    def __init__(self, value: str = ""):
        self.value = value

    def get(self) -> str:
        return self.value

    def set(self, value: str) -> None:
        self.value = value


class DummyStatus:
    def __init__(self):
        self.value = ""

    def set(self, value: str) -> None:
        self.value = value


def _ts(year: int, month: int, day: int, hour: int, minute: int) -> int:
    return int(datetime(year, month, day, hour, minute, tzinfo=EASTERN_TZ).timestamp() * 1000)


def test_group_bars_by_day_orders_days_and_timestamps():
    bars = [
        {"t": _ts(2025, 1, 3, 9, 32), "c": 101.0},
        {"t": _ts(2025, 1, 2, 9, 31), "c": 99.0},
        {"t": _ts(2025, 1, 2, 9, 30), "c": 98.0},
    ]

    grouped = _group_bars_by_day(bars)

    assert [day for day, _ in grouped] == ["2025-01-02", "2025-01-03"]
    assert [row["c"] for row in grouped[0][1]] == [98.0, 99.0]


def test_minute_of_day_uses_eastern_clock():
    assert _minute_of_day(_ts(2025, 1, 2, 9, 30)) == 570


def test_speed_multiplier_is_clamped_and_defaults_when_invalid():
    page = IntradayReplayPage.__new__(IntradayReplayPage)
    page.speed_var = Var("200")
    assert page._read_speed_multiplier() == 10.0

    page.speed_var = Var("0")
    assert page._read_speed_multiplier() == 0.1

    page.speed_var = Var("bad")
    assert page._read_speed_multiplier() == 1.0


def test_ticker_and_days_back_parsing_guardrails():
    page = IntradayReplayPage.__new__(IntradayReplayPage)
    page.ticker_var = Var(" nvda ")
    page.days_back_var = Var("0")
    assert page._read_ticker() == "NVDA"
    assert page._read_days_back() == 1

    page.days_back_var = Var("500")
    assert page._read_days_back() == 60

    page.days_back_var = Var("bad")
    assert page._read_days_back() == 14


def test_shift_day_guardrails_and_autoplay_behavior():
    page = IntradayReplayPage.__new__(IntradayReplayPage)
    page.day_series = [("2025-01-02", []), ("2025-01-03", [])]
    page.day_index = 0
    page.stage_index = 0
    page.status_var = DummyStatus()

    calls = {"stop": 0, "draw": 0, "schedule": 0}
    page.stop_replay = lambda: calls.__setitem__("stop", calls["stop"] + 1)
    page._draw_current_day = lambda *_, **__: calls.__setitem__("draw", calls["draw"] + 1)
    page._schedule_next_stage = lambda: calls.__setitem__("schedule", calls["schedule"] + 1)

    page.shift_day(-1, autoplay=True)
    assert page.day_index == 0
    assert "oldest" in page.status_var.value
    assert calls == {"stop": 0, "draw": 0, "schedule": 0}

    page.shift_day(1, autoplay=True)
    assert page.day_index == 1
    assert calls["stop"] == 1
    assert calls["schedule"] == 1
    assert calls["draw"] == 0

    page.shift_day(1, autoplay=True)
    assert page.day_index == 1
    assert "newest" in page.status_var.value
