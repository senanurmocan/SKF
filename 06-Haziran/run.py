"""Start the SKF Streamlit app locally or as a hidden Windows background process."""

from __future__ import annotations

import argparse
import ctypes
import hashlib
import json
import os
import signal
import socket
import subprocess
import sys
import time
import urllib.error
import urllib.request
import webbrowser
from pathlib import Path
from typing import Any


FIRST_PORT = 8505
PORT_LIMIT = 8600
HEALTH_TIMEOUT_SECONDS = 60
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000


def find_free_port(start_port: int = FIRST_PORT) -> int:
    """Find a free loopback port without relying on another app's connection state."""
    for port in range(start_port, PORT_LIMIT):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as candidate:
            try:
                candidate.bind(("127.0.0.1", port))
            except OSError:
                continue
            return port
    raise RuntimeError(f"{start_port}-{PORT_LIMIT - 1} aralığında boş port bulunamadı.")


def get_local_ip() -> str:
    """Return the local network address for the explicitly enabled LAN mode."""
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as connection:
            connection.connect(("10.254.254.254", 1))
            return connection.getsockname()[0]
    except OSError:
        return "127.0.0.1"


def package_root() -> Path:
    return Path(__file__).resolve().parent.parent


def background_paths() -> tuple[Path, Path]:
    root = package_root()
    app_data = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
    state_id = hashlib.sha256(str(root).casefold().encode("utf-8")).hexdigest()[:16]
    state_directory = app_data / "SKF-Yerel" / "state"
    log_directory = app_data / "SKF-Yerel" / "logs"
    state_directory.mkdir(parents=True, exist_ok=True)
    log_directory.mkdir(parents=True, exist_ok=True)
    return state_directory / f"server-{state_id}.json", log_directory / f"server-{state_id}.log"


def read_server_state(path: Path) -> dict[str, Any] | None:
    try:
        state = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    if not isinstance(state, dict):
        return None
    try:
        state["pid"] = int(state["pid"])
        state["port"] = int(state["port"])
        state["executable"] = str(state["executable"])
        state["package_root"] = str(state["package_root"])
    except (KeyError, TypeError, ValueError):
        return None
    if not 1 <= state["port"] <= 65535 or state["pid"] <= 0:
        return None
    return state


def process_image(pid: int) -> str | None:
    """Return a Windows process image path so stale/reused PIDs are not stopped."""
    if os.name != "nt":
        return None
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.OpenProcess.restype = ctypes.c_void_p
    kernel32.OpenProcess.argtypes = [ctypes.c_ulong, ctypes.c_int, ctypes.c_ulong]
    handle = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not handle:
        return None
    try:
        buffer = ctypes.create_unicode_buffer(32768)
        size = ctypes.c_ulong(len(buffer))
        query = kernel32.QueryFullProcessImageNameW
        query.argtypes = [ctypes.c_void_p, ctypes.c_ulong, ctypes.c_wchar_p, ctypes.POINTER(ctypes.c_ulong)]
        if not query(handle, 0, buffer, ctypes.byref(size)):
            return None
        return str(Path(buffer.value).resolve())
    finally:
        kernel32.CloseHandle(ctypes.c_void_p(handle))


def process_is_running(pid: int, expected_executable: str) -> bool:
    actual = process_image(pid)
    if not actual:
        return False
    try:
        return os.path.normcase(actual) == os.path.normcase(str(Path(expected_executable).resolve()))
    except OSError:
        return False


def server_is_healthy(port: int, timeout: float = 0.8) -> bool:
    try:
        with urllib.request.urlopen(
            f"http://127.0.0.1:{port}/_stcore/health", timeout=timeout
        ) as response:
            return response.status == 200
    except (OSError, urllib.error.URLError):
        return False


def show_message(message: str, *, error: bool = False) -> None:
    if os.name == "nt":
        icon = 0x10 if error else 0x40
        ctypes.windll.user32.MessageBoxW(None, message, "SKF Fatura Veri Portalı", icon)
    else:
        print(message, file=sys.stderr if error else sys.stdout)


def write_server_state(path: Path, state: dict[str, Any]) -> None:
    temporary = path.with_suffix(".tmp")
    temporary.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
    temporary.replace(path)


def launch_command(
    app_path: Path,
    port: int,
    *,
    address: str = "127.0.0.1",
    headless: bool = True,
    xsrf: bool = False,
) -> list[str]:
    return [
        sys.executable,
        "-m",
        "streamlit",
        "run",
        str(app_path),
        "--server.port",
        str(port),
        "--server.address",
        address,
        "--server.headless",
        str(headless).lower(),
        "--server.enableXsrfProtection",
        str(xsrf).lower(),
        "--server.enableCORS",
        "false",
        "--browser.gatherUsageStats",
        "false",
    ]


def open_app(port: int, *, open_browser: bool) -> None:
    url = f"http://127.0.0.1:{port}"
    if open_browser:
        webbrowser.open(url, new=2)


def start_background(port: int | None = None, *, open_browser: bool = True) -> int:
    if os.name != "nt":
        show_message("Arka plan başlatıcısı yalnızca Windows paketinde kullanılabilir.", error=True)
        return 2

    root = package_root()
    try:
        state_path, log_path = background_paths()
    except OSError as error:
        show_message(f"SKF çalışma klasörü oluşturulamadı.\n\n{error}", error=True)
        return 1
    existing = read_server_state(state_path)
    if existing and os.path.normcase(existing["package_root"]) == os.path.normcase(str(root)):
        expected_executable = existing["executable"]
        if process_is_running(existing["pid"], expected_executable):
            deadline = time.monotonic() + 12
            while time.monotonic() < deadline:
                if server_is_healthy(existing["port"]):
                    open_app(existing["port"], open_browser=open_browser)
                    return 0
                if not process_is_running(existing["pid"], expected_executable):
                    break
                time.sleep(0.4)
            if process_is_running(existing["pid"], expected_executable):
                show_message("Uygulama başlatılıyor. Biraz bekleyip tarayıcıyı yenileyin.")
                return 0
        state_path.unlink(missing_ok=True)

    try:
        selected_port = port or find_free_port()
    except RuntimeError as error:
        show_message(str(error), error=True)
        return 2

    app_path = Path(__file__).resolve().parent / "app.py"
    environment = os.environ.copy()
    environment["SKF_FILE_INPUT_MODE"] = "native"
    environment["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"
    flags = getattr(subprocess, "DETACHED_PROCESS", 0x00000008) | getattr(
        subprocess, "CREATE_NO_WINDOW", 0x08000000
    )
    try:
        with log_path.open("a", encoding="utf-8") as app_log:
            app_log.write(
                f"\n[{time.strftime('%Y-%m-%d %H:%M:%S')}] Uygulama başlatılıyor; port={selected_port}\n"
            )
            app_log.flush()
            process = subprocess.Popen(
                launch_command(app_path, selected_port),
                cwd=str(app_path.parent),
                env=environment,
                stdin=subprocess.DEVNULL,
                stdout=app_log,
                stderr=subprocess.STDOUT,
                creationflags=flags,
                close_fds=True,
            )
    except OSError as error:
        show_message(f"Uygulama arka planda başlatılamadı.\n\n{error}\n\nGünlük: {log_path}", error=True)
        return 1

    state = {
        "pid": process.pid,
        "port": selected_port,
        "url": f"http://127.0.0.1:{selected_port}",
        "executable": sys.executable,
        "package_root": str(root),
        "started_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
    }
    try:
        write_server_state(state_path, state)
    except OSError as error:
        process.terminate()
        show_message(f"SKF başlatma bilgisi kaydedilemedi.\n\n{error}", error=True)
        return 1

    deadline = time.monotonic() + HEALTH_TIMEOUT_SECONDS
    while time.monotonic() < deadline:
        if server_is_healthy(selected_port):
            open_app(selected_port, open_browser=open_browser)
            return 0
        if process.poll() is not None:
            break
        time.sleep(0.5)

    if process.poll() is None:
        process.terminate()
    state_path.unlink(missing_ok=True)
    show_message(
        "Uygulama başlayamadı. Günlük dosyasındaki hata ayrıntılarını kontrol edin:\n\n"
        f"{log_path}",
        error=True,
    )
    return 1


def stop_background() -> int:
    if os.name != "nt":
        show_message("Arka plan başlatıcısı yalnızca Windows paketinde kullanılabilir.", error=True)
        return 2

    try:
        state_path, log_path = background_paths()
    except OSError as error:
        show_message(f"SKF çalışma klasörüne erişilemedi.\n\n{error}", error=True)
        return 1
    state = read_server_state(state_path)
    if state is None or os.path.normcase(state["package_root"]) != os.path.normcase(str(package_root())):
        state_path.unlink(missing_ok=True)
        show_message("Çalışan SKF arka plan uygulaması bulunamadı.")
        return 0

    if not process_is_running(state["pid"], state["executable"]):
        state_path.unlink(missing_ok=True)
        show_message("SKF uygulaması zaten kapalı.")
        return 0

    try:
        os.kill(state["pid"], signal.SIGTERM)
    except OSError as error:
        show_message(f"Uygulama durdurulamadı.\n\n{error}\n\nGünlük: {log_path}", error=True)
        return 1

    deadline = time.monotonic() + 8
    while time.monotonic() < deadline and process_is_running(state["pid"], state["executable"]):
        time.sleep(0.2)
    state_path.unlink(missing_ok=True)
    show_message("SKF uygulaması durduruldu.")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="SKF Fatura Veri Portalı başlatıcısı")
    parser.add_argument("--lan", action="store_true", help="Güvenilir LAN erişimini aç (XSRF korumalı).")
    parser.add_argument("--lan-open", action="store_true", help="LAN erişimini XSRF koruması olmadan aç (geliştirme).")
    parser.add_argument("--port", type=int, default=None, help="Sabit port numarası (varsayılan: otomatik).")
    parser.add_argument("--background", action="store_true", help="Windows uygulamasını görünmeyen arka planda başlat.")
    parser.add_argument("--stop", action="store_true", help="Bu paketin arka plan uygulamasını durdur.")
    parser.add_argument("--no-browser", action="store_true", help=argparse.SUPPRESS)
    args = parser.parse_args()

    if args.stop:
        return stop_background()
    if args.background:
        if args.lan or args.lan_open:
            parser.error("Arka plan modu yalnızca yerel 127.0.0.1 erişimini destekler.")
        return start_background(args.port, open_browser=not args.no_browser)

    base_dir = Path(__file__).resolve().parent
    app_path = base_dir / "app.py"
    port = args.port or find_free_port()
    address = "0.0.0.0" if (args.lan or args.lan_open) else "127.0.0.1"
    url = f"http://{get_local_ip() if address == '0.0.0.0' else '127.0.0.1'}:{port}"

    print("=" * 70)
    print("SKF Fatura Veri Portalı başlatılıyor…")
    print(f"Erişim adresi: {url}")
    print("Pencereyi kapatmak uygulamayı durdurur.")
    print("=" * 70)

    cmd = launch_command(
        app_path,
        port,
        address=address,
        headless=bool(args.lan or args.lan_open),
        xsrf=bool(args.lan),
    )
    environment = os.environ.copy()
    if args.lan or args.lan_open:
        environment["SKF_FILE_INPUT_MODE"] = "browser"
    else:
        environment["SKF_FILE_INPUT_MODE"] = "native"
    return subprocess.run(cmd, cwd=str(base_dir), env=environment, check=False).returncode


if __name__ == "__main__":
    raise SystemExit(main())
