import os
from pathlib import Path

STATE_PATH = Path(__file__).resolve().parent / "app_state.txt"
CONFIG_DIR = Path.home() / ".stoptions_analyzer"
API_KEY_PATH = CONFIG_DIR / "api_key.txt"
DATA_DIR = Path(__file__).resolve().parent / "data"
API_BASE_URL = os.getenv("MASSIVE_BASE_URL", "https://api.polygon.io")
HORIZON_CONFIGS = [
    ("Day", 1, 10, "10m"),
    ("3 Day", 3, 30, "30m"),
    ("Week", 7, 60, "1h"),
    ("Month", 30, 120, "2h"),
    ("3M", 90, 360, "6h"),
    ("6M", 180, 720, "12h"),
    ("12M", 365, 1440, "1d"),
    ("3Y", 1095, 4320, "3d"),
    ("5Y", 1825, 7200, "5d"),
    ("10Y", 3650, 10080, "7d"),
]
ANALYSIS_OUTPUT_DIR = DATA_DIR / "analysis_outputs"
BACKTEST_CACHE_DIR = DATA_DIR / "backtest_cache"
BACKTEST_OUTPUT_DIR = DATA_DIR / "backtest_outputs"
DEFAULT_GENERAL_ANALYSIS_SETTINGS = {
    "analysis_type": "Cross-Sectional",
    "cross_sectional_strategy": "Momentum",
    "time_series_strategy": "Momentum",
    "lookback_days": 90,
    "skip_days": 5,
    "top_quantile": 0.2,
    "bottom_quantile": 0.2,
    "momentum_use_volatility_scaling": False,
    "momentum_use_residual": False,
    "momentum_use_multi_horizon": False,
    "time_series_use_volatility_scaling": False,
    "time_series_use_residual": False,
    "time_series_use_multi_horizon": False,
    "time_series_use_zscore": False,
    "time_series_winsorize_sigma": None,
    "output_dir": str(ANALYSIS_OUTPUT_DIR),
}

DEFAULT_BACKTEST_SETTINGS = {
    "strategy_name": "Time-Series Momentum",
    "run_mode": "single",
    "lookback_days": "90",
    "skip_days": "5",
    "entry_signal": "ts_momentum",
    "entry_signal_params": "{}",
    "exit_signal": "none",
    "exit_signal_params": "{}",
    "costs_bps": "5",
    "sweep_entry_grid": '{"ts_momentum": [{"lookback_days": 90, "skip_days": 5}]}',
    "sweep_exit_grid": '{"none": [{}]}',
    "sweep_core_grid": '{"lookback_days": [90], "skip_days": [5], "costs_bps": [5.0]}',
    "sweep_seed": "42",
    "sweep_max_workers": "",
    "sweep_top_n": "10",
    "sweep_fail_fast": False,
    "sweep_continue_on_error": True,
    "start_date": "",
    "end_date": "",
    "notes": "",
    "backtest_data_root": str(BACKTEST_CACHE_DIR),
}
