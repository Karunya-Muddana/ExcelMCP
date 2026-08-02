"""Shared helpers for the ~/.excelmcp config directory.

Every file this package writes there (OAuth token cache, structure graph,
embedding vectors) is written through here so that two properties hold
everywhere:

* **Restrictive permissions.** The token cache holds a refresh token; a
  world-readable copy is a credential leak. On POSIX the directory is 0700 and
  files are 0600. On Windows ``os.chmod`` only toggles the read-only bit, so we
  rely on the default per-user ACL on ``%USERPROFILE%`` instead — the chmod
  calls are best-effort and never fatal.

* **Atomic replacement.** A crash or a full disk partway through a write used to
  leave a truncated ``graph.json`` that made every subsequent tool call fail.
  Writes land in a temp file in the same directory, are flushed to disk, then
  moved into place with ``os.replace``, which is atomic on both POSIX and NTFS.
"""

import json
import os
import tempfile
from pathlib import Path
from typing import Any

CONFIG_DIR_NAME = ".excelmcp"

_DIR_MODE = 0o700
_FILE_MODE = 0o600


def _chmod_quietly(path: Path, mode: int) -> None:
    """Best-effort permission tightening; a no-op where the platform ignores it."""
    try:
        os.chmod(path, mode)
    except (OSError, NotImplementedError):
        pass


def get_config_dir() -> Path:
    """Returns ~/.excelmcp, creating it with restrictive permissions if needed."""
    config_dir = Path.home() / CONFIG_DIR_NAME
    config_dir.mkdir(parents=True, exist_ok=True)
    _chmod_quietly(config_dir, _DIR_MODE)
    return config_dir


def atomic_write_text(path: Path, text: str, *, sensitive: bool = False) -> None:
    """Writes text to path atomically, replacing any existing file.

    Set sensitive=True for anything credential-bearing so the file is created
    0600 before content is written to it.
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp_name = tempfile.mkstemp(
        dir=str(path.parent), prefix=f".{path.name}.", suffix=".tmp"
    )
    tmp_path = Path(tmp_name)
    try:
        if sensitive:
            _chmod_quietly(tmp_path, _FILE_MODE)
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(text)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp_path, path)
    except BaseException:
        tmp_path.unlink(missing_ok=True)
        raise
    if sensitive:
        _chmod_quietly(path, _FILE_MODE)


def atomic_write_bytes(path: Path, data: bytes) -> None:
    """Writes bytes to path atomically, replacing any existing file."""
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp_name = tempfile.mkstemp(
        dir=str(path.parent), prefix=f".{path.name}.", suffix=".tmp"
    )
    tmp_path = Path(tmp_name)
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(data)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp_path, path)
    except BaseException:
        tmp_path.unlink(missing_ok=True)
        raise


def read_json(path: Path, default: Any) -> Any:
    """Reads JSON, returning default if the file is absent, corrupt, or unreadable.

    A hand-edited or half-written config should degrade to "empty" rather than
    crashing every tool call in the server.
    """
    if not path.exists():
        return default
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError, UnicodeDecodeError) as exc:
        print(f"[Storage] Ignoring unreadable {path.name} ({exc}); treating as empty.")
        return default
