r"""Build an offline, user-local Windows package for the SKF Streamlit app.

Run this from the repository root with its configured virtual environment:
    .\.venv\Scripts\python.exe tools\build_local_windows_package.py

The bundle contains only application source, approved reference workbooks,
and the Python runtime/dependencies. Local input files, preferences, caches,
test data, and generated reports are deliberately excluded.
"""

from __future__ import annotations

import hashlib
import importlib.metadata
import os
import re
import shutil
import sys
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
APP_SOURCE = ROOT / "06-Haziran"
REQUIREMENTS = APP_SOURCE / "requirements.txt"
PACKAGE_NAME = "SKF-Yerel-Paket-1.2.0"
OUTPUT_DIR = ROOT / "dist"
PACKAGE_ZIP = OUTPUT_DIR / f"{PACKAGE_NAME}.zip"
PACKAGE_CHECKSUM = OUTPUT_DIR / f"{PACKAGE_NAME}.sha256"

APP_FILES = (
    "app.py",
    "extract_and_compare.py",
    "new_mapping_parser.py",
    "run.py",
    "CK_Enerji_-_Yatay.png",
    "SKF Başlıkları.xlsx",
    "_PortalAdmin_Uploads_Content_FastAccess_0bd2da5220923.xlsx",
)
CONFIG_JSON_FILES = ("header_overrides.json", "source_header_layouts.json")


def require(condition: bool, message: str) -> None:
    if not condition:
        raise SystemExit(message)


def check_build_environment() -> Path:
    require(os.name == "nt", "Bu paketleyici Windows üzerinde çalıştırılmalıdır.")
    require(sys.version_info[:2] == (3, 14), "Paket Python 3.14 ile oluşturulmalıdır.")
    require(sys.prefix != sys.base_prefix, "Paketleyiciyi proje .venv ortamından çalıştırın.")
    require(APP_SOURCE.is_dir(), f"Uygulama klasörü bulunamadı: {APP_SOURCE}")
    require(REQUIREMENTS.is_file(), f"Bağımlılık listesi bulunamadı: {REQUIREMENTS}")

    for line in REQUIREMENTS.read_text(encoding="utf-8").splitlines():
        requirement = line.strip()
        if not requirement or requirement.startswith("#"):
            continue
        match = re.fullmatch(r"([A-Za-z0-9_.-]+)==([A-Za-z0-9_.+-]+)", requirement)
        require(match is not None, f"Paketleyici sabitlenmiş bağımlılık bekliyor: {requirement}")
        name, expected_version = match.groups()
        try:
            actual_version = importlib.metadata.version(name)
        except importlib.metadata.PackageNotFoundError:
            actual_version = "kurulu değil"
        require(
            actual_version == expected_version,
            f"{name} sürümü {expected_version} olmalı; ortamda {actual_version} var.",
        )

    for relative_path in APP_FILES:
        require((APP_SOURCE / relative_path).is_file(), f"Paket girdisi bulunamadı: {relative_path}")

    base = Path(sys.base_prefix)
    required_runtime_paths = (
        base / "python.exe",
        base / "Lib",
        base / "DLLs",
        base / f"python{sys.version_info.major}{sys.version_info.minor}.dll",
        base / "python3.dll",
        base / "vcruntime140.dll",
    )
    for runtime_path in required_runtime_paths:
        require(runtime_path.exists(), f"Python taşınabilir çalışma zamanı girdisi yok: {runtime_path}")

    site_packages = Path(sys.prefix) / "Lib" / "site-packages"
    require(site_packages.is_dir(), f"Bağımlılıklar bulunamadı: {site_packages}")
    return site_packages


def copy_application(destination: Path) -> None:
    app_destination = destination / "06-Haziran"
    app_destination.mkdir(parents=True)

    for relative_path in APP_FILES:
        source = APP_SOURCE / relative_path
        shutil.copy2(source, app_destination / relative_path)

    for package_dir in ("core", "analytics", "config"):
        source_dir = APP_SOURCE / package_dir
        target_dir = app_destination / package_dir
        for source in sorted(source_dir.rglob("*.py")):
            if "__pycache__" in source.parts:
                continue
            target = target_dir / source.relative_to(source_dir)
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source, target)

    config_source = APP_SOURCE / "config"
    config_target = app_destination / "config"
    for filename in CONFIG_JSON_FILES:
        source = config_source / filename
        require(source.is_file(), f"Uygulama kural dosyası bulunamadı: {source}")
        shutil.copy2(source, config_target / filename)


def copy_runtime(destination: Path, site_packages: Path) -> None:
    runtime = destination / "runtime"
    runtime.mkdir()
    base = Path(sys.base_prefix)

    def ignore_base_site_packages(_directory: str, names: list[str]) -> set[str]:
        return {name for name in names if name == "site-packages" or name == "__pycache__"}

    shutil.copytree(base / "Lib", runtime / "Lib", ignore=ignore_base_site_packages)
    shutil.copytree(base / "DLLs", runtime / "DLLs")

    runtime_files = [base / "python.exe", base / "python3.dll"]
    runtime_files.extend(base.glob("python*.dll"))
    runtime_files.extend(base.glob("vcruntime140*.dll"))
    runtime_files.extend(base.glob("python*.zip"))
    for source in dict.fromkeys(runtime_files):
        if source.is_file():
            shutil.copy2(source, runtime / source.name)

    target_site_packages = runtime / "Lib" / "site-packages"
    shutil.copytree(
        site_packages,
        target_site_packages,
        ignore=shutil.ignore_patterns("__pycache__", "*.pyc"),
    )

    python_path_file = runtime / f"python{sys.version_info.major}{sys.version_info.minor}._pth"
    python_path_file.write_text(
        "python314.zip\n.\nLib\nDLLs\nLib\\site-packages\nimport site\n",
        encoding="ascii",
        newline="\n",
    )
    shutil.copy2(base / "LICENSE.txt", destination / "Python-LICENSE.txt")


def write_package_readme(destination: Path) -> None:
    readme = """SKF Fatura Veri Portalı — Yerel Windows Paketi

Başlatma
1. Bu ZIP'i kullanıcı profilinizde yazma izni olan bir klasöre çıkarın.
   Önerilen konum: %LOCALAPPDATA%\\SKF-Yerel
2. Klasördeki `SKF'yi Başlat.bat` dosyasına çift tıklayın.
3. Portal tarayıcıda açılır. Durdurmak için açılan siyah pencereyi kapatın.

Kullanım ve veri akışı
- Uygulama yalnızca bu bilgisayarda, 127.0.0.1 adresinde çalışır.
- Bölge klasörü, Bizim Kestiğimiz Faturalar, Uzlaştırma Tüketim ve diğer
  kaynaklar bu bilgisayardaki Windows dosya seçicileriyle seçilir.
- Kaynak dosyalar Streamlit Cloud'a veya başka bir sunucuya yüklenmez.
- Analiz ve Excel çıktısı bu bilgisayarda hazırlanır.
- Streamlit kullanım istatistikleri kapalıdır.
- Kullanıcı dosyaları, seçilmiş dosya yolları ve rapor çıktıları bu ZIP'e dahil
  edilmez. Seçim tercihleri uygulama klasöründeki config dosyalarına yazılır.
- Bu pakette dağıtım faturası, müşteri dosyası veya ETSO listesi bulunmaz.

Güncelleme ve dağıtım
- Şirket içi dağıtım için BT ekibinizin onaylı yazılım dağıtım kanalını kullanın.
- Paketi herkese açık GitHub deposuna veya kişisel dosya paylaşım alanlarına
  koymayın; şirketinizin kurum içi kanal ve imza politikasını izleyin.
- Güncel paket kullanıcı profili altında çalıştırılmalıdır. Program klasörü
  salt okunursa dosya seçimi tercihleri saklanamayabilir.
- Windows uygulama denetimi, SmartScreen veya Forcepoint bu paketi engellerse
  BT ekibinizin onaylı imza/allowlist sürecini izleyin.

Teknik bilgiler
- Python çalışma zamanı ve uygulama bağımlılıkları paketin içindedir.
- Python veya pip kurulumu, yönetici izni ya da internet bağlantısı gerekmez.
- Uygulama yerel web arayüzü için yalnızca loopback portu açar; LAN erişimi
  etkinleştirilmez.
"""
    (destination / "README-Yerel-Kullanim.txt").write_text(readme, encoding="utf-8")


def write_launcher(destination: Path) -> None:
    launcher = r"""@echo off
setlocal
chcp 65001 >nul
title SKF Fatura Veri Portalı (Yerel Kullanım)
cd /d "%~dp0"
if not exist "%~dp0runtime\python.exe" (
  echo Bu paketin runtime klasörü bulunamadı. ZIP'i yeniden çıkarın.
  pause
  exit /b 1
)
set "SKF_FILE_INPUT_MODE=native"
set "STREAMLIT_BROWSER_GATHER_USAGE_STATS=false"
"%~dp0runtime\python.exe" "%~dp006-Haziran\run.py"
set "SKF_EXIT_CODE=%ERRORLEVEL%"
if not "%SKF_EXIT_CODE%"=="0" (
  echo Uygulama hata koduyla kapandı: %SKF_EXIT_CODE%
  pause
)
exit /b %SKF_EXIT_CODE%
"""
    (destination / "SKF'yi Başlat.bat").write_text(
        launcher,
        encoding="utf-8-sig",
        newline="\r\n",
    )


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as package_file:
        for block in iter(lambda: package_file.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def main() -> int:
    site_packages = check_build_environment()
    require(
        not PACKAGE_ZIP.exists() and not PACKAGE_CHECKSUM.exists(),
        f"Çıktı zaten var; üzerine yazmıyorum: {OUTPUT_DIR}",
    )
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    staging_root = ROOT / ".local-package-build"
    staging_root.mkdir(exist_ok=True)
    staging_parent = staging_root.resolve()
    staging = Path(tempfile.mkdtemp(prefix="build-", dir=staging_root)).resolve()
    require(staging_parent in staging.parents, "Güvenli geçici paket klasörü oluşturulamadı.")
    package_root = staging / PACKAGE_NAME
    package_root.mkdir()

    try:
        copy_application(package_root)
        copy_runtime(package_root, site_packages)
        write_launcher(package_root)
        write_package_readme(package_root)
        build_info = (
            f"Package: {PACKAGE_NAME}\n"
            f"Python: {sys.version.split()[0]}\n"
            f"Built: {datetime.now().astimezone().isoformat(timespec='seconds')}\n"
        )
        (package_root / "BUILD-INFO.txt").write_text(build_info, encoding="utf-8")

        with zipfile.ZipFile(
            PACKAGE_ZIP,
            mode="x",
            compression=zipfile.ZIP_DEFLATED,
            compresslevel=6,
            allowZip64=True,
        ) as archive:
            for file_path in sorted(package_root.rglob("*")):
                if file_path.is_file():
                    archive.write(file_path, file_path.relative_to(staging))

        package_hash = sha256(PACKAGE_ZIP)
        PACKAGE_CHECKSUM.write_text(
            f"{package_hash}  {PACKAGE_ZIP.name}\n",
            encoding="ascii",
        )
        size_mb = PACKAGE_ZIP.stat().st_size / (1024 * 1024)
        print(f"Paket: {PACKAGE_ZIP}")
        print(f"SHA-256: {package_hash}")
        print(f"Boyut: {size_mb:.1f} MB")
        print("Arşiv yalnızca kodu ve sabit referans dosyalarını içerir.")
        return 0
    finally:
        resolved_staging = staging.resolve()
        require(
            staging_parent in resolved_staging.parents
            and resolved_staging.name.startswith("build-"),
            "Geçici klasör beklenen paket alanının dışında; temizlik durduruldu.",
        )
        shutil.rmtree(resolved_staging)


if __name__ == "__main__":
    raise SystemExit(main())
