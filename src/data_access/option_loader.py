from data_access.api_client import MassiveApiClient
from data_access.cache import load_cached_market_data, save_cached_market_data
from utils.parsing import effective_market_date, normalize_option_records


def load_option_records(api_client: MassiveApiClient, ticker: str) -> list[dict]:
    cache_payload = load_cached_market_data(ticker) or {}
    cache_date = cache_payload.get("last_updated")
    today_label = effective_market_date().isoformat()
    cached_options = cache_payload.get("options")
    if cached_options is not None and cache_date == today_label:
        return normalize_option_records(cached_options or [])
    option_data = api_client.fetch_option_snapshots(ticker)
    option_records = normalize_option_records(option_data)
    cache_payload.update(
        {
            "last_updated": today_label,
            "options": option_records,
        }
    )
    save_cached_market_data(ticker, cache_payload)
    return option_records
