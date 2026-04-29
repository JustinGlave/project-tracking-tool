"""
app_backend.py — Data and storage logic for <AppName>

Replace the example Item dataclass and ItemStore with your actual data model.
Data is stored as JSON in %APPDATA%\\ATS Inc\\<AppName>\\.
"""

from __future__ import annotations

import json
import os
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Optional


# ── Path helper ───────────────────────────────────────────────────────────────

def _app_data_path(filename: str) -> str:
    """Path to user data in %APPDATA%\\ATS Inc\\<AppName>\\."""
    base = Path(os.environ.get("APPDATA", Path.home())) / "ATS Inc" / "YOUR APP NAME"
    base.mkdir(parents=True, exist_ok=True)
    return str(base / filename)


DATA_FILE = _app_data_path("data.json")


# ── Data model ────────────────────────────────────────────────────────────────

@dataclass
class Item:
    """Replace this with your actual data model."""
    id: str
    name: str
    description: str = ""
    # Add more fields here


# ── Storage ───────────────────────────────────────────────────────────────────

class ItemStore:
    """Loads and saves items to a JSON file."""

    def __init__(self) -> None:
        self._items: list[Item] = []
        self._load()

    def _load(self) -> None:
        if not os.path.exists(DATA_FILE):
            return
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as fh:
                raw = json.load(fh)
            self._items = [Item(**entry) for entry in raw.get("items", [])]
        except (OSError, json.JSONDecodeError, TypeError):
            self._items = []

    def _save(self) -> None:
        with open(DATA_FILE, "w", encoding="utf-8") as fh:
            json.dump({"items": [asdict(i) for i in self._items]}, fh, indent=2)

    # ── Public API ────────────────────────────────────────────────────────────

    def all(self) -> list[Item]:
        return list(self._items)

    def get(self, item_id: str) -> Optional[Item]:
        return next((i for i in self._items if i.id == item_id), None)

    def add(self, item: Item) -> None:
        self._items.append(item)
        self._save()

    def update(self, item: Item) -> None:
        for idx, existing in enumerate(self._items):
            if existing.id == item.id:
                self._items[idx] = item
                self._save()
                return

    def delete(self, item_id: str) -> None:
        self._items = [i for i in self._items if i.id != item_id]
        self._save()
