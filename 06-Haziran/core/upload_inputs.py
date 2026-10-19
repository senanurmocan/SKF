"""Stage browser uploads for the existing filesystem-based invoice readers.

Each Streamlit session should keep its own ``UploadWorkspace`` in session_state.
The workspace contains only uploaded copies and can safely be discarded after use.
"""

from __future__ import annotations

import hashlib
import io
import re
import shutil
import stat
import tempfile
import unicodedata
import zipfile
import zlib
from collections.abc import Iterable
from dataclasses import dataclass
from pathlib import Path, PureWindowsPath
from typing import Any, Callable


MIB = 1024 * 1024
DISTRIBUTION_EXTENSIONS = {".xlsx", ".xls", ".xml", ".html"}
EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}


class UploadInputError(ValueError):
    """An uploaded file or directory cannot be used safely by the readers."""


@dataclass(frozen=True)
class _InputFile:
    parts: tuple[str, ...]
    data: bytes | memoryview


@dataclass(frozen=True)
class _StagedInput:
    signature: str
    directory: Path
    result: Path


def _relative_parts(name: Any) -> tuple[str, ...]:
    """Interpret browser/ZIP paths consistently on both Windows and Linux."""
    if not isinstance(name, str) or not name:
        raise UploadInputError("Yüklenen dosyanın adı bulunamadı.")
    normalized = unicodedata.normalize("NFC", name).replace("\\", "/")
    if normalized.startswith("/") or PureWindowsPath(normalized).drive:
        raise UploadInputError("Yüklemede mutlak dosya yolu kullanılamaz.")
    parts = tuple(normalized.split("/"))
    invalid_characters = '<>:"|?*'
    for part in parts:
        if (
            not part
            or part in {".", ".."}
            or part != part.rstrip(" .")
            or any(ord(character) < 32 or character in invalid_characters for character in part)
        ):
            raise UploadInputError("Yüklemede geçersiz veya klasör dışına çıkan dosya yolu var.")
        stem = part.split(".", 1)[0].upper()
        if stem in {"CON", "PRN", "AUX", "NUL"} or re.fullmatch(r"(?:COM|LPT)[1-9]", stem):
            raise UploadInputError("Yüklemede işletim sistemine ayrılmış bir dosya adı var.")
    return parts


def _ignored(parts: tuple[str, ...]) -> bool:
    return parts[-1].startswith("~$") or any(
        part.casefold() == "__macosx" or part.startswith(".") for part in parts
    )


def _region_key(value: str) -> str:
    translated = value.casefold().translate(str.maketrans("ıöüçşğ", "ioucsg"))
    return "".join(character for character in translated if character.isalnum())


def _check_path_collisions(paths: Iterable[tuple[str, ...]]) -> None:
    files = {tuple(part.casefold() for part in parts) for parts in paths}
    if any(parts[:index] in files for parts in files for index in range(1, len(parts))):
        raise UploadInputError("Yüklemede aynı yol hem dosya hem klasör olarak kullanılmış.")


def _default_region_names() -> set[str]:
    from config.mappings import (
        REFERENCE_TO_EXTRACTION_MAPPING,
        REGION_NAME_MAPPING,
        REGION_NORMALIZATION_MAPPING,
    )

    names: set[str] = set()
    for mapping in (
        REFERENCE_TO_EXTRACTION_MAPPING,
        REGION_NAME_MAPPING,
        REGION_NORMALIZATION_MAPPING,
    ):
        names.update(mapping)
        names.update(mapping.values())
    return names


class UploadWorkspace:
    """Own an isolated temporary filesystem tree for one browser session."""

    def __init__(
        self,
        *,
        max_total_bytes: int = 400 * MIB,
        max_files: int = 2000,
        max_distribution_file_bytes: int = 50 * MIB,
        region_names: Iterable[str] | None = None,
    ) -> None:
        if min(max_total_bytes, max_files, max_distribution_file_bytes) <= 0:
            raise ValueError("Yükleme sınırları pozitif olmalıdır.")
        self.max_total_bytes = max_total_bytes
        self.max_files = max_files
        self.max_distribution_file_bytes = max_distribution_file_bytes
        names = _default_region_names()
        if region_names is not None:
            names.update(region_names)
        self._region_keys = {_region_key(name) for name in names}
        self._temporary = tempfile.TemporaryDirectory(prefix="skf-upload-")
        self.root_path = Path(self._temporary.name).resolve()
        self._staged: dict[str, _StagedInput] = {}
        self._closed = False

    def __enter__(self) -> UploadWorkspace:
        self._ensure_open()
        return self

    def __exit__(self, *args: Any) -> None:
        self.close()

    def _ensure_open(self) -> None:
        if self._closed:
            raise UploadInputError("Yükleme oturumu kapandı; dosyaları yeniden seçin.")

    def _read_upload(self, upload: Any) -> bytes | memoryview:
        declared_size = getattr(upload, "size", None)
        if isinstance(declared_size, int) and declared_size > self.max_total_bytes:
            raise UploadInputError("Yüklenen dosya toplam yükleme boyutu sınırını aşıyor.")
        if callable(getattr(upload, "getbuffer", None)):
            data = upload.getbuffer()
        elif callable(getattr(upload, "getvalue", None)):
            data = upload.getvalue()
        else:
            raise UploadInputError("Yüklenen dosyanın içeriği okunamadı.")
        if not isinstance(data, (bytes, memoryview)):
            data = bytes(data)
        if len(data) > self.max_total_bytes:
            raise UploadInputError("Yüklenen dosya toplam yükleme boyutu sınırını aşıyor.")
        return data

    def _check_distribution_size(self, size: int) -> None:
        if size > self.max_distribution_file_bytes:
            limit = self.max_distribution_file_bytes / MIB
            raise UploadInputError(
                f"Bir dağıtım dosyası {limit:g} MB sınırını aşıyor; yükleme yapılmadı."
            )

    def _check_collection(self, count: int, total: int) -> None:
        if count > self.max_files:
            raise UploadInputError(f"Bir yüklemede en fazla {self.max_files} dosya seçilebilir.")
        if total > self.max_total_bytes:
            limit = self.max_total_bytes / MIB
            raise UploadInputError(f"Toplam yükleme boyutu {limit:g} MB sınırını aşıyor.")

    def _distribution_root(self, paths: Iterable[tuple[str, ...]]) -> tuple[str, ...]:
        candidates: set[tuple[str, ...]] = set()
        for parts in paths:
            for index, part in enumerate(parts[:-1]):
                if _region_key(part) in self._region_keys:
                    candidates.add(parts[:index])
                    break
        if not candidates:
            raise UploadInputError(
                "Bölge klasörleri bulunamadı. Dağıtım şirketi klasörlerini içeren ana klasörü "
                "veya aynı klasör yapısını koruyan ZIP dosyasını seçin."
            )
        if len(candidates) != 1:
            raise UploadInputError(
                "Birden fazla dağıtım ana klasörü bulundu. Tek aya ait ana klasörü seçin."
            )
        return next(iter(candidates))

    def _destination(self, directory: Path, parts: tuple[str, ...]) -> Path:
        destination = directory.joinpath(*parts)
        if not destination.resolve().is_relative_to(directory.resolve()):
            raise UploadInputError("Yüklenen dosya çalışma klasörünün dışına çıkamaz.")
        return destination

    def _remove_directory(self, directory: Path) -> None:
        resolved = directory.resolve()
        if resolved == self.root_path or not resolved.is_relative_to(self.root_path):
            raise UploadInputError("Geçici yükleme klasörü güvenle temizlenemedi.")
        if directory.exists():
            shutil.rmtree(directory)

    def _publish(
        self,
        category: str,
        signature: str,
        result_parts: tuple[str, ...],
        write: Callable[[Path], None],
    ) -> str:
        self._ensure_open()
        if not isinstance(category, str) or not re.fullmatch(r"[A-Za-z0-9_-]{1,64}", category):
            raise UploadInputError("Geçersiz yükleme alanı.")
        previous = self._staged.get(category)
        if previous and previous.signature == signature and previous.result.exists():
            return str(previous.result)
        directory = self.root_path / f"{category}-{signature}"
        self._remove_directory(directory)
        directory.mkdir()
        try:
            write(directory)
            result = self._destination(directory, result_parts)
            if not result.exists():
                raise UploadInputError("Yüklenen dosyalar hazırlanamadı.")
        except Exception as error:
            self._remove_directory(directory)
            if isinstance(error, OSError):
                raise UploadInputError("Yüklenen dosyalar geçici çalışma alanına yazılamadı.") from error
            raise
        self._staged[category] = _StagedInput(signature, directory, result)
        if previous and previous.directory != directory:
            self._remove_directory(previous.directory)
        return str(result)

    def stage_file(self, upload: Any, category: str) -> str:
        """Stage a single Excel workbook, retaining its original base name."""
        self._ensure_open()
        parts = _relative_parts(getattr(upload, "name", None))
        filename = parts[-1]
        if _ignored(parts) or Path(filename).suffix.casefold() not in EXCEL_EXTENSIONS:
            raise UploadInputError("Lütfen .xlsx, .xlsm veya .xls biçiminde bir Excel dosyası seçin.")
        data = self._read_upload(upload)
        digest = hashlib.sha256(filename.encode("utf-8") + b"\0")
        digest.update(data)
        signature = digest.hexdigest()
        return self._publish(
            category,
            signature,
            (filename,),
            lambda directory: self._destination(directory, (filename,)).write_bytes(data),
        )

    def stage_directory(self, uploads: Iterable[Any]) -> str:
        """Stage a directory while preserving all relevant relative file paths."""
        self._ensure_open()
        files: list[_InputFile] = []
        seen: set[str] = set()
        total = 0
        for upload in uploads:
            parts = _relative_parts(getattr(upload, "name", None))
            if _ignored(parts) or Path(parts[-1]).suffix.casefold() not in DISTRIBUTION_EXTENSIONS:
                continue
            key = "/".join(parts).casefold()
            if key in seen:
                raise UploadInputError("Aynı göreli dosya yolu yüklemede birden fazla kez bulundu.")
            seen.add(key)
            data = self._read_upload(upload)
            self._check_distribution_size(len(data))
            total += len(data)
            self._check_collection(len(files) + 1, total)
            files.append(_InputFile(parts, data))
        if not files:
            raise UploadInputError("Seçilen klasörde desteklenen dağıtım dosyası bulunamadı.")
        files.sort(key=lambda item: item.parts)
        _check_path_collisions(item.parts for item in files)
        root_parts = self._distribution_root(item.parts for item in files)
        digest = hashlib.sha256()
        for item in files:
            digest.update("/".join(item.parts).encode("utf-8"))
            digest.update(b"\0")
            digest.update(hashlib.sha256(item.data).digest())

        def write(directory: Path) -> None:
            for item in files:
                destination = self._destination(directory, item.parts)
                destination.parent.mkdir(parents=True, exist_ok=True)
                destination.write_bytes(item.data)

        return self._publish("distribution", digest.hexdigest(), root_parts, write)

    def stage_zip(self, upload: Any) -> str:
        """Read a ZIP without extractall, enforcing paths and expanded-size limits."""
        self._ensure_open()
        parts = _relative_parts(getattr(upload, "name", None))
        if Path(parts[-1]).suffix.casefold() != ".zip":
            raise UploadInputError("Lütfen ZIP biçiminde bir dosya seçin.")
        data = self._read_upload(upload)
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as archive:
                entries: list[tuple[zipfile.ZipInfo, tuple[str, ...]]] = []
                seen: set[str] = set()
                total = 0
                file_count = 0
                for entry in archive.infolist():
                    entry_parts = _relative_parts(entry.orig_filename.rstrip("/"))
                    mode = entry.external_attr >> 16
                    file_type = stat.S_IFMT(mode)
                    if file_type not in {0, stat.S_IFREG, stat.S_IFDIR}:
                        raise UploadInputError("ZIP içinde sembolik bağlantı veya özel dosya kullanılamaz.")
                    if entry.is_dir():
                        continue
                    file_count += 1
                    self._check_collection(file_count, 0)
                    if _ignored(entry_parts) or Path(entry_parts[-1]).suffix.casefold() not in DISTRIBUTION_EXTENSIONS:
                        continue
                    if entry.flag_bits & 1:
                        raise UploadInputError("Şifreli ZIP desteklenmiyor; şifresiz ZIP dosyası seçin.")
                    key = "/".join(entry_parts).casefold()
                    if key in seen:
                        raise UploadInputError("ZIP içinde aynı dosya yolu birden fazla kez bulundu.")
                    seen.add(key)
                    self._check_distribution_size(entry.file_size)
                    total += entry.file_size
                    self._check_collection(file_count, total)
                    entries.append((entry, entry_parts))
                if not entries:
                    raise UploadInputError("ZIP içinde desteklenen dağıtım dosyası bulunamadı.")
                _check_path_collisions(parts for _, parts in entries)
                root_parts = self._distribution_root(parts for _, parts in entries)

                def write(directory: Path) -> None:
                    expanded_total = 0
                    for entry, entry_parts in entries:
                        destination = self._destination(directory, entry_parts)
                        destination.parent.mkdir(parents=True, exist_ok=True)
                        file_total = 0
                        with archive.open(entry) as source, destination.open("wb") as target:
                            while chunk := source.read(MIB):
                                file_total += len(chunk)
                                expanded_total += len(chunk)
                                self._check_distribution_size(file_total)
                                self._check_collection(len(entries), expanded_total)
                                target.write(chunk)

                return self._publish(
                    "distribution", hashlib.sha256(data).hexdigest(), root_parts, write
                )
        except (zipfile.BadZipFile, OSError, RuntimeError, NotImplementedError, zlib.error) as error:
            raise UploadInputError("ZIP dosyası okunamadı; sağlam ve şifresiz bir ZIP seçin.") from error

    def clear(self, category: str) -> None:
        """Remove one input category without affecting other selected sources."""
        self._ensure_open()
        previous = self._staged.pop(category, None)
        if previous is not None:
            self._remove_directory(previous.directory)

    def close(self) -> None:
        """Discard all uploaded copies belonging to this workspace."""
        if not self._closed:
            self._temporary.cleanup()
            self._staged.clear()
            self._closed = True
