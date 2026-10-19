"""Synthetic upload tests; no operational workbooks or user paths are needed."""

from __future__ import annotations

import io
import stat
import struct
import zipfile
from pathlib import Path

import pytest

from core.upload_inputs import UploadInputError, UploadWorkspace


class Upload(io.BytesIO):
    def __init__(self, name: str, data: bytes = b"workbook") -> None:
        super().__init__(data)
        self.name = name
        self.size = len(data)


def zipped(entries: dict[str, bytes], name: str = "dagitim.zip") -> Upload:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for path, data in entries.items():
            archive.writestr(path, data)
    return Upload(name, buffer.getvalue())


@pytest.fixture
def workspace():
    with UploadWorkspace() as current:
        yield current


def test_single_workbook_preserves_name_and_reruns_do_not_rewrite(workspace):
    upload = Upload("başka isim.XLSX", b"first")
    first = Path(workspace.stage_file(upload, "invoice"))
    modified = first.stat().st_mtime_ns
    second = Path(workspace.stage_file(Upload(upload.name, b"first"), "invoice"))
    assert first == second
    assert second.name == upload.name
    assert second.read_bytes() == b"first"
    assert second.stat().st_mtime_ns == modified


def test_new_content_gets_new_path_and_removes_old_category(workspace):
    first = Path(workspace.stage_file(Upload("invoice.xlsx", b"first"), "invoice"))
    other = Path(workspace.stage_file(Upload("eno.xls", b"lookup"), "eno"))
    second = Path(workspace.stage_file(Upload("invoice.xlsx", b"second"), "invoice"))
    assert first != second
    assert not first.exists()
    assert second.read_bytes() == b"second"
    assert other.read_bytes() == b"lookup"


def test_directory_resolves_month_wrapper_and_preserves_duplicate_base_names(workspace):
    root = Path(workspace.stage_directory([
        Upload("09-2026/Boğaziçi EDAŞ/invoice.xlsx", b"bogazici"),
        Upload("09-2026/Vangölü EDAŞ/alt/invoice.xlsx", b"vangolu"),
        Upload("09-2026/09 ENO Eylul.xlsx", b"lookup"),
    ]))
    assert root.name == "09-2026"
    assert (root / "Boğaziçi EDAŞ/invoice.xlsx").read_bytes() == b"bogazici"
    assert (root / "Vangölü EDAŞ/alt/invoice.xlsx").read_bytes() == b"vangolu"
    assert (root / "09 ENO Eylul.xlsx").read_bytes() == b"lookup"


def test_selecting_single_region_returns_its_parent(workspace):
    root = Path(workspace.stage_directory([Upload("BEDAŞ/alt/veri.xls")]))
    assert root.parent == workspace.root_path
    assert (root / "BEDAŞ/alt/veri.xls").is_file()


def test_directory_accepts_windows_relative_paths_and_core_extensions(workspace):
    root = Path(workspace.stage_directory([
        Upload("Ay\\GDZ EDAŞ\\a.XLS"),
        Upload("Ay\\GDZ EDAŞ\\alt\\b.XML"),
        Upload("Ay\\GDZ EDAŞ\\c.html"),
    ]))
    assert (root / "GDZ EDAŞ/a.XLS").is_file()
    assert (root / "GDZ EDAŞ/alt/b.XML").is_file()
    assert (root / "GDZ EDAŞ/c.html").is_file()


def test_directory_order_does_not_change_staged_path(workspace):
    files = [Upload("Ay/BEDAŞ/one.xlsx"), Upload("Ay/AYEDAŞ/two.xls")]
    first = workspace.stage_directory(files)
    second = workspace.stage_directory(reversed(files))
    assert first == second


def test_directory_drops_lock_mac_metadata_and_unhandled_files(workspace):
    root = Path(workspace.stage_directory([
        Upload("Ay/BEDAŞ/real.xlsx"),
        Upload("Ay/BEDAŞ/~$real.xlsx"),
        Upload("Ay/__MACOSX/BEDAŞ/metadata.xlsx"),
        Upload("Ay/BEDAŞ/notes.pdf"),
        Upload("Ay/.hidden/BEDAŞ/hidden.xlsx"),
    ]))
    assert [path.name for path in root.rglob("*") if path.is_file()] == ["real.xlsx"]


def test_flat_file_selection_cannot_lose_region_information(workspace):
    with pytest.raises(UploadInputError, match="Bölge klasörleri"):
        workspace.stage_directory([Upload("invoice.xlsx")])


def test_multiple_month_roots_are_rejected(workspace):
    with pytest.raises(UploadInputError, match="Birden fazla"):
        workspace.stage_directory([
            Upload("Eylul/BEDAŞ/a.xlsx"), Upload("Ekim/AYEDAŞ/b.xlsx"),
        ])


def test_custom_region_alias_can_be_added():
    with UploadWorkspace(region_names=["Yeni Bölge"]) as workspace:
        root = Path(workspace.stage_directory([Upload("Ay/Yeni Bölge/a.xlsx")]))
        assert (root / "Yeni Bölge/a.xlsx").is_file()


def test_zip_resolves_nested_wrapper_and_skips_metadata(workspace):
    root = Path(workspace.stage_zip(zipped({
        "arşiv/Ay/BEDAŞ/a.xlsx": b"a",
        "arşiv/Ay/AYEDAŞ/alt/b.html": b"b",
        "arşiv/Ay/BEDAŞ/~$a.xlsx": b"lock",
        "__MACOSX/._a.xlsx": b"metadata",
    })))
    assert root.name == "Ay"
    assert (root / "BEDAŞ/a.xlsx").read_bytes() == b"a"
    assert (root / "AYEDAŞ/alt/b.html").read_bytes() == b"b"
    assert not (root / "BEDAŞ/~$a.xlsx").exists()


def test_zip_and_directory_replace_same_distribution_category(workspace):
    first = Path(workspace.stage_directory([Upload("BEDAŞ/a.xlsx")]))
    second = Path(workspace.stage_zip(zipped({"Ay/AYEDAŞ/b.xlsx": b"new"})))
    assert not first.exists()
    assert (second / "AYEDAŞ/b.xlsx").read_bytes() == b"new"


def test_identical_zip_is_not_rewritten(workspace):
    upload = zipped({"Ay/BEDAŞ/a.xlsx": b"a"})
    first = Path(workspace.stage_zip(upload))
    modified = (first / "BEDAŞ/a.xlsx").stat().st_mtime_ns
    assert workspace.stage_zip(upload) == str(first)
    assert (first / "BEDAŞ/a.xlsx").stat().st_mtime_ns == modified


@pytest.mark.parametrize("name", [
    "../outside.xlsx", "Ay/../outside.xlsx", "/outside.xlsx", "C:/outside.xlsx",
    "C:outside.xlsx", "\\\\server\\share\\outside.xlsx", "Ay\\..\\outside.xlsx",
    "Ay/BEDAŞ/a.xlsx:stream", "Ay/BEDAŞ/a\x00.xlsx", "Ay/BEDAŞ/CON.xlsx",
])
def test_unsafe_paths_are_rejected_for_directory_and_zip(workspace, name):
    with pytest.raises(UploadInputError):
        workspace.stage_directory([Upload(name)])
    # ZIP names containing NUL are normalized/truncated by Python's ZIP writer.
    if "\x00" not in name:
        with pytest.raises(UploadInputError):
            workspace.stage_zip(zipped({name: b"x"}))
    assert not list(workspace.root_path.iterdir())


def test_single_excel_rejects_traversal_and_invalid_types(workspace):
    for name in ["../invoice.xlsx", "invoice.zip", "~$invoice.xlsx"]:
        with pytest.raises(UploadInputError):
            workspace.stage_file(Upload(name), "invoice")


def test_zip_symlink_is_rejected(workspace):
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        entry = zipfile.ZipInfo("Ay/BEDAŞ/link.xlsx")
        entry.create_system = 3
        entry.external_attr = (stat.S_IFLNK | 0o777) << 16
        archive.writestr(entry, "../../outside.xlsx")
    with pytest.raises(UploadInputError, match="sembolik"):
        workspace.stage_zip(Upload("upload.zip", buffer.getvalue()))


def test_corrupt_compressed_zip_payload_is_a_recoverable_input_error(workspace):
    good = Path(workspace.stage_directory([Upload("BEDAŞ/good.xlsx", b"original")]))
    upload = zipped({"BEDAŞ/broken.xlsx": b"compressed contents" * 500})
    payload = bytearray(upload.getvalue())
    name_size, extra_size = struct.unpack_from("<HH", payload, 26)
    compressed_offset = 30 + name_size + extra_size
    payload[compressed_offset] = 0x07  # DEFLATE block type 3 is invalid.
    with pytest.raises(UploadInputError, match="ZIP dosyası okunamadı"):
        workspace.stage_zip(Upload("broken.zip", bytes(payload)))
    assert (good / "BEDAŞ/good.xlsx").read_bytes() == b"original"
    assert len(list(workspace.root_path.iterdir())) == 1


@pytest.mark.parametrize("method", ["directory", "zip"])
def test_expanded_distribution_file_limit(method):
    with UploadWorkspace(max_distribution_file_bytes=4) as workspace:
        with pytest.raises(UploadInputError, match="dağıtım dosyası"):
            if method == "directory":
                workspace.stage_directory([Upload("BEDAŞ/a.xlsx", b"12345")])
            else:
                workspace.stage_zip(zipped({"BEDAŞ/a.xlsx": b"12345"}))
        assert not list(workspace.root_path.iterdir())


@pytest.mark.parametrize("method", ["directory", "zip"])
def test_total_expanded_size_limit(method):
    # ZIP headers exceed a tiny cap, so use compressed, individually valid entries.
    files = {"BEDAŞ/a.xlsx": b"a" * 600, "AYEDAŞ/b.xlsx": b"b" * 600}
    with UploadWorkspace(max_total_bytes=1000) as workspace:
        with pytest.raises(UploadInputError, match="Toplam yükleme"):
            if method == "directory":
                workspace.stage_directory([Upload(name, value) for name, value in files.items()])
            else:
                workspace.stage_zip(zipped(files))


@pytest.mark.parametrize("method", ["directory", "zip"])
def test_file_count_limit(method):
    files = {"BEDAŞ/a.xlsx": b"a", "AYEDAŞ/b.xlsx": b"b"}
    with UploadWorkspace(max_files=1) as workspace:
        with pytest.raises(UploadInputError, match="en fazla 1 dosya"):
            if method == "directory":
                workspace.stage_directory([Upload(name, value) for name, value in files.items()])
            else:
                workspace.stage_zip(zipped(files))


def test_duplicate_relative_paths_are_rejected_case_insensitively(workspace):
    with pytest.raises(UploadInputError, match="birden fazla"):
        workspace.stage_directory([Upload("BEDAŞ/a.xlsx"), Upload("BEDAŞ/A.xlsx")])
    with pytest.raises(UploadInputError, match="birden fazla"):
        workspace.stage_zip(zipped({"BEDAŞ/a.xlsx": b"a", "BEDAŞ/A.xlsx": b"b"}))


def test_conflicting_file_and_directory_paths_are_rejected(workspace):
    files = {"BEDAŞ/a.xlsx": b"a", "BEDAŞ/A.xlsx/child.xls": b"b"}
    with pytest.raises(UploadInputError, match="hem dosya hem klasör"):
        workspace.stage_directory([Upload(name, data) for name, data in files.items()])
    with pytest.raises(UploadInputError, match="hem dosya hem klasör"):
        workspace.stage_zip(zipped(files))


def test_getvalue_only_upload_is_supported(workspace):
    class ValueUpload:
        name = "invoice.xlsx"
        size = 3

        def getvalue(self):
            return b"123"

    assert Path(workspace.stage_file(ValueUpload(), "invoice")).read_bytes() == b"123"


def test_invalid_replacement_retains_previous_input(workspace):
    first = Path(workspace.stage_directory([Upload("BEDAŞ/a.xlsx", b"original")]))
    with pytest.raises(UploadInputError):
        workspace.stage_zip(Upload("bad.zip", b"not a zip"))
    assert (first / "BEDAŞ/a.xlsx").read_bytes() == b"original"


def test_sessions_are_isolated_and_close_cleans_only_own_copies():
    with UploadWorkspace() as first, UploadWorkspace() as second:
        first_file = Path(first.stage_file(Upload("a.xls", b"first"), "invoice"))
        second_file = Path(second.stage_file(Upload("a.xls", b"second"), "invoice"))
        assert first.root_path != second.root_path
        first.close()
        assert not first_file.exists()
        assert second_file.read_bytes() == b"second"
        with pytest.raises(UploadInputError, match="kapandı"):
            first.stage_file(Upload("a.xls"), "invoice")


def test_clear_removes_only_requested_category(workspace):
    invoice = Path(workspace.stage_file(Upload("invoice.xlsx"), "invoice"))
    eno = Path(workspace.stage_file(Upload("eno.xlsx"), "eno"))
    workspace.clear("invoice")
    workspace.clear("invoice")
    assert not invoice.exists()
    assert eno.exists()


def test_invalid_category_cannot_escape_workspace(workspace):
    with pytest.raises(UploadInputError, match="yükleme alanı"):
        workspace.stage_file(Upload("invoice.xlsx"), "../outside")
    assert not list(workspace.root_path.iterdir())
