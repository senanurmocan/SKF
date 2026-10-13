# -*- coding: utf-8 -*-
"""Known monthly source layouts and unexpected format change warnings."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
from typing import Any, Iterable

from core.number_cleaner import normalize_header


_LAYOUT_CONFIG_PATH = Path(__file__).resolve().parents[1] / 'config' / 'source_header_layouts.json'
try:
    with _LAYOUT_CONFIG_PATH.open(encoding='utf-8') as layout_file:
        _LAYOUT_CONFIG = json.load(layout_file)
except (OSError, json.JSONDecodeError):
    _LAYOUT_CONFIG = {'baseline_month': '2026-06', 'header_signatures': {}}

_BASELINE_MONTH = _LAYOUT_CONFIG.get('baseline_month', '2026-06')
_KNOWN_LAYOUTS = {
    normalize_header(region_name): frozenset(
        [signature] if isinstance(signature, str) else signature
    )
    for region_name, signature in _LAYOUT_CONFIG.get('header_signatures', {}).items()
}


def is_known_source_header_layout(region_name: str | None) -> bool:
    """Return whether a region has an approved header layout baseline."""
    return normalize_header(region_name) in _KNOWN_LAYOUTS


def find_effective_header_row(rows: Iterable[Iterable[Any]], scan_limit: int = 5):
    """Find the first likely header row, skipping a short report title/banner."""
    scanned_rows = list(rows)[:scan_limit]
    if not scanned_rows:
        return 0, ()

    for row_index, row in enumerate(scanned_rows):
        if sum(1 for value in row if value is not None and str(value).strip()) > 2:
            return row_index, row
    return 0, scanned_rows[0]


def _header_signature(headers: Iterable[Any]) -> str:
    normalized_headers = [
        (index, normalize_header(value))
        for index, value in enumerate(headers)
        if value is not None and str(value).strip()
    ]
    serialized = json.dumps(normalized_headers, ensure_ascii=False, separators=(',', ':'))
    return hashlib.sha256(serialized.encode('utf-8')).hexdigest()


def known_layout_change_warning(
    headers: Iterable[Any],
    region_name: str,
    source_name: str,
    header_row_number: int | None = None,
) -> dict[str, str] | None:
    """Describe a source format change from the approved monthly layout."""
    headers = tuple(headers)
    expected_signatures = _KNOWN_LAYOUTS.get(normalize_header(region_name))
    if not expected_signatures:
        return None

    actual_signature = _header_signature(headers)
    if actual_signature in expected_signatures:
        return None

    populated_header_count = sum(
        1 for value in headers if value is not None and str(value).strip()
    )
    row_description = (
        f"Başlık satırı {header_row_number} ({populated_header_count} dolu başlık)"
        if header_row_number is not None
        else f"{populated_header_count} dolu başlık"
    )
    return {
        'Dağıtım Bölgesi': region_name,
        'Kaynak': source_name,
        'Alan': 'Kaynak biçimi',
        'Eşlenen sütun': 'Başlık satırı',
        'Beklenen başlık': f'{_BASELINE_MONTH} için kayıtlı kaynak düzenlerinden biri',
        'Sütundaki başlık': 'Kaynak başlık düzeni değişmiş',
        'Başlığın bulunduğu sütun': row_description,
        'Düzen imzası': actual_signature[:16],
    }
