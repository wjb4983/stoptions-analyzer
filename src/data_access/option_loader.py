from data_access.api_client import MassiveApiClient
from data_access.cache import load_cached_market_data, save_cached_market_data
from utils.parsing import effective_market_date, normalize_option_records


def _option_identity(record: dict) -> tuple:
    return (
        record.get("ticker"),
        record.get("expiration_date"),
        record.get("contract_type"),
        record.get("strike_price"),
    )


def _validate_option_records(records: list[dict]) -> list[dict]:
    ordered = sorted(records, key=_option_identity)
    duplicates: set[tuple] = set()
    seen: set[tuple] = set()
    for record in ordered:
        identity = _option_identity(record)
        if identity in seen:
            duplicates.add(identity)
        seen.add(identity)

    if duplicates:
        rendered = ", ".join(str(dup) for dup in sorted(duplicates))
        raise ValueError(f"Duplicate option contracts detected: {rendered}")
    return ordered


def load_option_records(api_client: MassiveApiClient, ticker: str) -> list[dict]:
    cache_payload = load_cached_market_data(ticker) or {}
    cache_date = cache_payload.get("last_updated")
    today_label = effective_market_date().isoformat()
    cached_options = cache_payload.get("options")
    if cached_options is not None and cache_date == today_label:
        return _validate_option_records(normalize_option_records(cached_options or []))
    option_data = api_client.fetch_option_snapshots(ticker)
    option_records = _validate_option_records(normalize_option_records(option_data))
    cache_payload.update(
        {
            "last_updated": today_label,
            "options": option_records,
        }
    )
    save_cached_market_data(ticker, cache_payload)
    return option_records
