"""Provider boundary. Implement adapters here without coupling them to scoring/export."""
from abc import ABC, abstractmethod
from pathlib import Path
from typing import Iterable, Mapping, Any

class RacingDataProvider(ABC):
    @abstractmethod
    def fetch_racecards(self, date: str) -> Iterable[Mapping[str, Any]]:
        """Return normalized racecard rows for YYYY-MM-DD."""
        raise NotImplementedError

class CSVProvider(RacingDataProvider):
    def __init__(self, path: str | Path):
        self.path = Path(path)

    def fetch_racecards(self, date: str):
        import csv
        with self.path.open("r", encoding="utf-8-sig", newline="") as f:
            for row in csv.DictReader(f):
                if row.get("date") == date:
                    yield row
