"""
Launcher script for EDAS Bill Extraction & Analysis Portal.
Automatically finds an available port (e.g. 8505) to prevent conflicts with other local projects.

Güvenlik Modları:
  --local     : Yalnız localhost (varsayılan, en güvenli)
  --lan       : LAN erişimi açık (0.0.0.0), XSRF koruması aktif
  --lan-open  : LAN erişimi açık, XSRF kapalı (geliştirme ortamı)
"""

import os
import socket
import subprocess
import sys
import argparse


def find_free_port(start_port=8505):
    """Find a free TCP port starting from start_port."""
    port = start_port
    while port < 8600:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(('127.0.0.1', port)) != 0:
                return port
        port += 1
    return start_port


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
        "--server.headless", "true" if (args.lan or args.lan_open) else "false",
        "--server.enableXsrfProtection", "true" if args.lan else "false",
        "--server.enableCORS", "false",
        "--browser.gatherUsageStats", "false",
    ]
    return subprocess.run(cmd, cwd=os.path.dirname(app_path), check=False).returncode


if __name__ == "__main__":
    raise SystemExit(main())
