from __future__ import annotations

import hashlib
import json
import logging
import os
import secrets
import tempfile
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)

_PBKDF2_ITERATIONS = 260_000
_PBKDF2_ALGO = "sha256"


def _hash_password(password: str, salt: str) -> str:
    dk = hashlib.pbkdf2_hmac(
        _PBKDF2_ALGO,
        password.encode("utf-8"),
        bytes.fromhex(salt),
        _PBKDF2_ITERATIONS,
    )
    return dk.hex()


ROLES = ("admin", "user", "view_only")
ROLE_LABELS = {"admin": "Admin", "user": "User", "view_only": "View Only"}


@dataclass
class UserRecord:
    username: str
    password_hash: str
    salt: str
    must_change_password: bool = False
    created_at: str = ""
    role: str = "user"  # "admin", "user", or "view_only"


class UserManager:
    MIN_PASSWORD_LENGTH = 8

    def __init__(self, users_path: Path) -> None:
        self._path = users_path
        self._path.parent.mkdir(parents=True, exist_ok=True)
        self._auto_migrate_admin()

    # ------------------------------------------------------------------ #
    # Public API                                                           #
    # ------------------------------------------------------------------ #

    def list_users(self) -> list[str]:
        data = self._load()
        return list(data.keys())

    def create_user(
        self,
        username: str,
        password: str,
        must_change_password: bool = True,
        role: str = "user",
    ) -> None:
        username = username.strip()
        if not username:
            raise ValueError("Username cannot be empty.")
        if len(password) < self.MIN_PASSWORD_LENGTH:
            raise ValueError(
                f"Password must be at least {self.MIN_PASSWORD_LENGTH} characters."
            )
        if role not in ROLES:
            raise ValueError(f"Invalid role '{role}'. Must be one of: {', '.join(ROLES)}")

        data = self._load()
        if username.casefold() in {k.casefold() for k in data}:
            raise ValueError(f"User '{username}' already exists.")

        salt = secrets.token_hex(32)
        pw_hash = _hash_password(password, salt)
        record = UserRecord(
            username=username,
            password_hash=pw_hash,
            salt=salt,
            must_change_password=must_change_password,
            created_at=datetime.now().replace(microsecond=0).isoformat(sep=" "),
            role=role,
        )
        data[username] = asdict(record)
        self._save(data)

    def authenticate(self, username: str, password: str) -> Optional[UserRecord]:
        """Return UserRecord if credentials are valid, else None."""
        data = self._load()
        # Case-insensitive lookup
        match = next((v for k, v in data.items() if k.casefold() == username.casefold()), None)
        if match is None:
            return None
        expected = _hash_password(password, match["salt"])
        if not secrets.compare_digest(expected, match["password_hash"]):
            return None
        return UserRecord(**match)

    def change_password(self, username: str, new_password: str) -> None:
        if len(new_password) < self.MIN_PASSWORD_LENGTH:
            raise ValueError(
                f"Password must be at least {self.MIN_PASSWORD_LENGTH} characters."
            )
        data = self._load()
        key = next((k for k in data if k.casefold() == username.casefold()), None)
        if key is None:
            raise ValueError(f"User '{username}' not found.")
        salt = secrets.token_hex(32)
        data[key]["salt"] = salt
        data[key]["password_hash"] = _hash_password(new_password, salt)
        data[key]["must_change_password"] = False
        self._save(data)

    def delete_user(self, username: str) -> None:
        data = self._load()
        key = next((k for k in data if k.casefold() == username.casefold()), None)
        if key is None:
            return
        if data[key].get("role") == "admin":
            admin_count = sum(1 for v in data.values() if v.get("role") == "admin")
            if admin_count <= 1:
                raise ValueError("Cannot delete the only administrator account.")
        del data[key]
        self._save(data)

    def get_user(self, username: str) -> Optional[UserRecord]:
        data = self._load()
        match = next((v for k, v in data.items() if k.casefold() == username.casefold()), None)
        if match is None:
            return None
        return UserRecord(**match)

    def has_any_users(self) -> bool:
        return bool(self._load())

    def reset_password(self, username: str, temp_password: str) -> None:
        """Set a temporary password and flag the account for a forced change on next login."""
        self.change_password(username, temp_password)
        data = self._load()
        key = next((k for k in data if k.casefold() == username.casefold()), None)
        if key:
            data[key]["must_change_password"] = True
            self._save(data)

    def set_role(self, username: str, role: str) -> None:
        """Change the role for a user."""
        if role not in ROLES:
            raise ValueError(f"Invalid role '{role}'. Must be one of: {', '.join(ROLES)}")
        data = self._load()
        key = next((k for k in data if k.casefold() == username.casefold()), None)
        if key is None:
            raise ValueError(f"User '{username}' not found.")
        data[key]["role"] = role
        self._save(data)

    # ------------------------------------------------------------------ #
    # Internal helpers                                                     #
    # ------------------------------------------------------------------ #

    def _auto_migrate_admin(self) -> None:
        """Migrate legacy is_admin field → role, and ensure at least one admin exists."""
        if not self._path.exists():
            return
        try:
            raw = json.loads(self._path.read_text(encoding="utf-8"))
        except Exception:
            return
        if not raw:
            return

        changed = False
        for val in raw.values():
            if "role" not in val:
                # Migrate from old is_admin bool
                val["role"] = "admin" if val.pop("is_admin", False) else "user"
                changed = True
            else:
                # Remove legacy field if it still lingers
                if val.pop("is_admin", None) is not None:
                    changed = True

        # Ensure at least one admin
        if not any(v.get("role") == "admin" for v in raw.values()):
            first_key = next(iter(raw))
            raw[first_key]["role"] = "admin"
            changed = True
            logger.info("Auto-promoted '%s' to admin (migration).", first_key)

        if changed:
            self._save(raw)

    def _load(self) -> dict:
        if not self._path.exists():
            return {}
        try:
            data = json.loads(self._path.read_text(encoding="utf-8"))
            # Normalize legacy is_admin field so UserRecord(**match) never fails
            for val in data.values():
                if "role" not in val:
                    val["role"] = "admin" if val.pop("is_admin", False) else "user"
                else:
                    val.pop("is_admin", None)
            return data
        except Exception:
            logger.exception("Failed to load users file: %s", self._path)
            return {}

    def _save(self, data: dict) -> None:
        tmp_fd, tmp_str = tempfile.mkstemp(dir=self._path.parent, suffix=".tmp")
        tmp_path = Path(tmp_str)
        try:
            with open(tmp_fd, "w", encoding="utf-8") as fh:
                json.dump(data, fh, indent=2)
            tmp_path.replace(self._path)
        except Exception:
            tmp_path.unlink(missing_ok=True)
            raise
