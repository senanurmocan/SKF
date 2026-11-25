"""Read delimited text files exported by Excel in common Turkish encodings."""

from __future__ import annotations

import codecs
import csv
from collections.abc import Iterator
from pathlib import Path


def _is_decodable(path: str | Path, encoding: str) -> bool:
    decoder = codecs.getincrementaldecoder(encoding)(errors="strict")
    try:
        with open(path, "rb") as source:
            while chunk := source.read(1024 * 1024):
                decoder.decode(chunk)
        decoder.decode(b"", final=True)
    except UnicodeDecodeError:
        return False
    return True


def _detect_encoding(path: str | Path) -> str:
    with open(path, "rb") as source:
        prefix = source.read(4)
    if prefix.startswith((codecs.BOM_UTF16_LE, codecs.BOM_UTF16_BE)):
        return "utf-16"
    if prefix.startswith(codecs.BOM_UTF8):
        return "utf-8-sig"
    if _is_decodable(path, "utf-8"):
        return "utf-8-sig"
    if _is_decodable(path, "cp1254"):
        return "cp1254"
    raise ValueError("CSV dosyas\u0131n\u0131n karakter kodlamas\u0131 okunamad\u0131.")


def iter_csv_rows(path: str | Path) -> Iterator[list[str]]:
    """Yield CSV rows, detecting comma, semicolon, tab, or pipe separators."""
    encoding = _detect_encoding(path)
    with open(path, "r", encoding=encoding, newline="") as source:
        sample = source.read(65536)
        source.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
        except csv.Error:
            dialect = csv.excel
        yield from csv.reader(source, dialect)
