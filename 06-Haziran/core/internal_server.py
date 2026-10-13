"""Restricted file discovery for the company-internal deployment."""

from __future__ import annotations

import os
from pathlib import Path


EXCEL_SUFFIXES = {".xls", ".xlsx", ".xlsm"}


def configured_directory(setting_name: str) -> tuple[Path | None, str | None]:
    """Resolve a configured local/UNC directory without returning its path in errors."""
    raw_value = os.environ.get(setting_name, "").strip()
    if not raw_value:
        return None, "missing"

    expanded = os.path.expandvars(os.path.expanduser(raw_value))
    try:
        directory = Path(expanded).resolve(strict=True)
    except (OSError, RuntimeError, ValueError):
        return None, "unavailable"

    if not directory.is_dir():
        return None, "unavailable"
    return directory, None


def _is_direct_child(candidate: Path, root: Path) -> bool:
    try:
        resolved = candidate.resolve(strict=True)
        resolved_root = root.resolve(strict=True)
        return resolved.parent == resolved_root
    except (OSError, RuntimeError, ValueError):
        return False


def safe_directories(root: Path) -> list[Path]:
    """Return direct child directories only; never follow links outside the root."""
    candidates: list[Path] = []
    try:
        entries = list(os.scandir(root))
    except OSError:
        return candidates

    for entry in entries:
        if entry.name.startswith((".", "~", "$")) or entry.is_symlink():
            continue
        try:
            if entry.is_dir(follow_symlinks=False):
                candidate = Path(entry.path)
                if _is_direct_child(candidate, root):
                    candidates.append(candidate.resolve(strict=True))
        except (OSError, RuntimeError, ValueError):
            continue
    return sorted(candidates, key=lambda path: path.name.casefold())


def safe_excel_files(root: Path) -> list[Path]:
    """Return supported Excel files directly inside a configured folder."""
    candidates: list[Path] = []
    try:
        entries = list(os.scandir(root))
    except OSError:
        return candidates

    for entry in entries:
        if entry.name.startswith(("~$", ".")) or entry.is_symlink():
            continue
        if Path(entry.name).suffix.casefold() not in EXCEL_SUFFIXES:
            continue
        try:
            if entry.is_file(follow_symlinks=False):
                candidate = Path(entry.path)
                if _is_direct_child(candidate, root):
                    candidates.append(candidate.resolve(strict=True))
        except (OSError, RuntimeError, ValueError):
            continue
    return sorted(candidates, key=lambda path: path.name.casefold())
