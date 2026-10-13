"""
Launcher script for EDAS Bill Extraction & Analysis Portal.
Automatically finds an available port (e.g. 8505) to prevent conflicts with other local projects.

Güvenlik Modları:
  varsayılan  : Yalnız localhost (yerel kullanım)
  --server    : IIS arkasında localhost arka ucu (şirket içi dağıtım)
"""

import os
import socket
import subprocess
import sys
import argparse
from urllib.parse import urlsplit


def find_free_port(start_port=8505):
    """Find a free TCP port starting from start_port."""
    port = start_port
    while port < 8600:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(('127.0.0.1', port)) != 0:
                return port
        port += 1
    return start_port


def main() -> int:
    parser = argparse.ArgumentParser(description='EDAŞ Fatura Portalı Launcher')
    parser.add_argument('--server', action='store_true',
                        help='Şirket içi dağıtım: localhost arka ucu, IIS ters proxy ve Windows kimlik doğrulaması')
    parser.add_argument('--lan', action='store_true',
                        help='Eski parametre; güvenli şirket sunucusu moduna yönlendirir')
    parser.add_argument('--lan-open', action='store_true',
                        help=argparse.SUPPRESS)
    parser.add_argument('--port', type=int, default=None,
                        help='Sabit port numarası (varsayılan: otomatik)')
    args = parser.parse_args()

    if args.lan_open:
        parser.error('--lan-open kaldırıldı; XSRF korumasını kapatarak ağda sunucu başlatılamaz.')

    internal_server = args.server or args.lan
    process_env = os.environ.copy()
    if internal_server:
        process_env['SKF_INTERNAL_SERVER'] = '1'
        port = args.port or 8505
        address = '127.0.0.1'
        public_url = os.environ.get('SKF_PUBLIC_URL', '').strip()
        public_uri = urlsplit(public_url) if public_url else None
        if public_uri and public_uri.scheme != 'https':
            parser.error('SKF_PUBLIC_URL şirket içi sunucu modunda HTTPS olmalıdır.')
        public_host = public_uri.hostname if public_uri else None
        allowed_hosts = os.environ.get('SKF_ALLOWED_HOSTS', '').strip()
        if not allowed_hosts:
            allowed_hosts = public_host or 'localhost,127.0.0.1'
        print('=' * 70)
        print('CK Enerji SKF şirket içi sunucu arka ucu başlatılıyor...')
        print(f'Yalnızca bu sunucudan erişim: http://127.0.0.1:{port}')
        print('Kullanıcı erişimi IIS Windows Authentication ve AD yetki grubundan geçmelidir.')
        print('=' * 70)
    else:
        process_env.pop('SKF_INTERNAL_SERVER', None)
        port = args.port or find_free_port(8505)
        address = '127.0.0.1'
        print('=' * 70)
        print('EDAŞ Fatura Portalı yerel modda başlatılıyor...')
        print(f'Yerel erişim: http://localhost:{port}')
        print('Mod: yalnızca bu bilgisayar')
        print('=' * 70)

    base_dir = os.path.dirname(os.path.abspath(__file__))
    app_path = os.path.join(base_dir, 'app.py')

    cmd = [
        sys.executable, "-m", "streamlit", "run", app_path,
        "--server.port", str(port),
        "--server.address", address,
        "--server.headless", "true" if internal_server else "false",
        "--server.enableXsrfProtection", "true",
        "--server.enableCORS", "true",
        "--server.enableStaticServing", "false",
        "--server.disconnectedSessionTTL", "120",
        "--browser.gatherUsageStats", "false",
    ]
    if internal_server:
        cmd.extend(["--server.allowedHosts", allowed_hosts])
        public_url = os.environ.get('SKF_PUBLIC_URL', '').strip()
        if public_url:
            public_uri = urlsplit(public_url)
            public_origin = f"{public_uri.scheme}://{public_uri.netloc}"
            cmd.extend(["--server.corsAllowedOrigins", public_origin])
        cmd.extend(["--client.showErrorDetails", "none"])
    return subprocess.run(
        cmd,
        cwd=os.path.dirname(app_path),
        env=process_env,
        check=False,
    ).returncode


if __name__ == "__main__":
    raise SystemExit(main())
