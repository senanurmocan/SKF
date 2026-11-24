"""
Launcher script for EDAS Bill Extraction & Analysis Portal.
Automatically finds an available port (e.g. 8505) to prevent conflicts with other local projects.

Güvenlik Modları:
  --local     : Yalnız localhost (varsayılan, en güvenli)
  --lan       : LAN erişimi açık (0.0.0.0), XSRF koruması aktif
  --lan-open  : LAN erişimi açık, XSRF kapalı (geliştirme ortamı)
"""

import argparse
import os
import socket
import subprocess
import sys
import threading
import time
import traceback
import webbrowser


def show_message_box(message: str, title: str, flags: int = 0x10010) -> None:
    try:
        import ctypes

        message_box = ctypes.windll.user32.MessageBoxW
        message_box.argtypes = (
            ctypes.c_void_p,
            ctypes.c_wchar_p,
            ctypes.c_wchar_p,
            ctypes.c_uint,
        )
        message_box.restype = ctypes.c_int
        message_box(None, message, title, flags)
    except Exception:
        print(message, file=sys.stderr)


def show_startup_error(details: str, exit_code: int | None = None) -> None:
    message = "SKF başlatılamadı."
    if exit_code is not None:
        message += f"\nÇıkış kodu: {exit_code}"

    details = details.strip()
    if details:
        if len(details) > 4000:
            details = "... (son bölüm gösteriliyor) ...\n" + details[-4000:]
        message += f"\n\nHata ayrıntısı:\n{details}"
    else:
        message += "\n\nUygulama hata ayrıntısı üretmedi."

    show_message_box(message, "SKF - Başlatma Hatası")


def find_chrome_executable() -> str | None:
    candidates = [
        os.path.join(os.environ.get("PROGRAMFILES", ""), "Google", "Chrome", "Application", "chrome.exe"),
        os.path.join(os.environ.get("PROGRAMFILES(X86)", ""), "Google", "Chrome", "Application", "chrome.exe"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Google", "Chrome", "Application", "chrome.exe"),
    ]
    for candidate in candidates:
        if candidate and os.path.isfile(candidate):
            return candidate

    if os.name == "nt":
        try:
            import winreg

            app_path_keys = (
                (winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe"),
                (winreg.HKEY_LOCAL_MACHINE, r"Software\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe"),
                (winreg.HKEY_LOCAL_MACHINE, r"Software\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe"),
            )
            for root, key_path in app_path_keys:
                try:
                    with winreg.OpenKey(root, key_path) as key:
                        candidate, _ = winreg.QueryValueEx(key, None)
                    if os.path.isfile(candidate):
                        return candidate
                except OSError:
                    continue
        except ImportError:
            pass
    return None


def open_app_browser(url: str) -> bool:
    chrome_executable = find_chrome_executable()
    if chrome_executable:
        try:
            subprocess.Popen(
                [chrome_executable, url],
                stdin=subprocess.DEVNULL,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            return True
        except OSError:
            pass
    return webbrowser.open_new_tab(url)


def open_local_app_when_ready(port: int, stop_event: threading.Event) -> None:
    url = f"http://localhost:{port}"
    deadline = time.monotonic() + 60
    while not stop_event.is_set() and time.monotonic() < deadline:
        try:
            with socket.create_connection(("127.0.0.1", port), timeout=0.5):
                break
        except OSError:
            stop_event.wait(0.25)
    else:
        if not stop_event.is_set():
            show_browser_error(url, "Yerel uygulama sunucusu 60 saniye içinde yanıt vermedi.")
        return

    if stop_event.is_set():
        return
    try:
        opened = open_app_browser(url)
    except Exception as exc:
        show_browser_error(url, str(exc))
        return
    if not opened:
        show_browser_error(url)


def show_browser_error(url: str, details: str = "") -> None:
    message = (
        "SKF sunucusu çalışıyor ancak tarayıcı açılamadı. "
        "Google Chrome'da şu yerel adresi açın:\n\n"
        f"{url}"
    )
    if details:
        message += f"\n\nAyrıntı: {details}"
    show_message_box(message, "SKF - Tarayıcı Açılamadı", 0x10030)


def find_free_port(start_port=8505):
    """Find a free TCP port starting from start_port."""
    port = start_port
    while port < 8600:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(('127.0.0.1', port)) != 0:
                return port
        port += 1
    raise OSError("8505-8599 aralığında kullanılabilir yerel bağlantı noktası bulunamadı.")


def get_local_ip():
    """Makinenin LAN IP adresini bul."""
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
            s.connect(('10.254.254.254', 1))
            return s.getsockname()[0]
    except Exception:
        return '127.0.0.1'


def main() -> int:
    parser = argparse.ArgumentParser(description='EDAŞ Fatura Portalı Launcher')
    parser.add_argument('--lan', action='store_true',
                        help='LAN erişimi aç (0.0.0.0), XSRF korumalı')
    parser.add_argument('--lan-open', action='store_true',
                        help='LAN erişimi aç, XSRF kapalı (geliştirme)')
    parser.add_argument('--port', type=int, default=None,
                        help='Sabit port numarası (varsayılan: otomatik)')
    args = parser.parse_args()

    base_dir = os.path.dirname(os.path.abspath(__file__))
    app_path = os.path.join(base_dir, 'app.py')

    port = args.port or find_free_port(8505)
    address = '0.0.0.0' if (args.lan or args.lan_open) else '127.0.0.1'

    print("=" * 70)
    print(f"⚡ EDAŞ Fatura Portalı başlatılıyor...")

    if args.lan or args.lan_open:
        local_ip = get_local_ip()
        print(f"🌐 LAN Erişimi: http://{local_ip}:{port}")
        print(f"🔒 XSRF Koruması: {'AÇIK' if args.lan else 'KAPALI (geliştirme)'}")
    else:
        print(f"🌐 Yerel Erişim: http://localhost:{port}")
        print(f"🔒 Mod: Yalnız localhost (en güvenli)")
    print("=" * 70)

    cmd = [
        sys.executable, "-m", "streamlit", "run", app_path,
        "--server.port", str(port),
        "--server.address", address,
        "--server.headless", "true",
        "--server.enableXsrfProtection", "true" if args.lan else "false",
        "--server.enableCORS", "false",
        "--browser.gatherUsageStats", "false",
    ]
    app_environment = os.environ.copy()
    if args.lan or args.lan_open:
        app_environment["SKF_FILE_INPUT_MODE"] = "browser"
    browser_stop_event = None
    browser_thread = None
    if not args.lan and not args.lan_open:
        browser_stop_event = threading.Event()
        browser_thread = threading.Thread(
            target=open_local_app_when_ready,
            args=(port, browser_stop_event),
            daemon=True,
        )
        browser_thread.start()
    try:
        result = subprocess.run(
            cmd,
            cwd=os.path.dirname(app_path),
            env=app_environment,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            errors="replace",
            check=False,
        )
    finally:
        if browser_stop_event is not None:
            browser_stop_event.set()
            browser_thread.join(timeout=1)
    if result.returncode != 0:
        show_startup_error(result.stdout, result.returncode)
    return result.returncode


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except SystemExit:
        raise
    except Exception:
        show_startup_error(traceback.format_exc())
        raise SystemExit(1)
