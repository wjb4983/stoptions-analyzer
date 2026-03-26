from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict, Iterable, Optional


class SnapshotCacheConnector:
    """Loads per-ticker snapshot cache JSON files.

    Default search roots include:
    - src/data
    - ../stoptions_analyzer/src/data
    - ../stoptions-analyzer/src/data
    """

    def __init__(self, roots: Optional[Iterable[Path]] = None) -> None:
        self.roots = list(roots) if roots else [
            Path("src/data"),
            Path("../stoptions_analyzer/src/data"),
            Path("../stoptions-analyzer/src/data"),
        ]

    def load(self, safe_ticker: str) -> Dict[str, Any]:
        filename = f"{safe_ticker}.json"
        for root in self.roots:
            path = root / filename
            if path.exists():
                with path.open("r", encoding="utf-8") as f:
                    return json.load(f)
        searched = ", ".join(str(r / filename) for r in self.roots)
        raise FileNotFoundError(f"Snapshot JSON not found. Searched: {searched}")
