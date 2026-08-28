#!/usr/bin/env python3
"""
AuditWorkpaper Pro — Local Python Web Server
Menjalankan aplikasi web secara lokal menggunakan Python built-in http.server.

Cara Penggunaan:
    python server.py
    python server.py --port 8080
    python server.py --no-browser
"""

import os
import sys
import socket
import argparse
import webbrowser
import threading
from http.server import HTTPServer, SimpleHTTPRequestHandler

# Pastikan encoding stdout aman di Windows
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

# Direktori root aplikasi
PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))

# MIME types tambahan
EXTRA_MIME_TYPES = {
    '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    '.xls': 'application/vnd.ms-excel',
    '.js': 'application/javascript',
    '.mjs': 'application/javascript',
    '.json': 'application/json',
    '.css': 'text/css',
    '.svg': 'image/svg+xml',
    '.png': 'image/png',
    '.ico': 'image/x-icon',
    '.pdf': 'application/pdf',
    '.wasm': 'application/wasm',
}


class AuditWorkpaperHTTPHandler(SimpleHTTPRequestHandler):
    """Custom HTTP Request Handler dengan dukungan CORS dan MIME types lengkap."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=PROJECT_DIR, **kwargs)

    def guess_type(self, path):
        ext = os.path.splitext(path)[1].lower()
        if ext in EXTRA_MIME_TYPES:
            return EXTRA_MIME_TYPES[ext]
        return super().guess_type(path)

    def end_headers(self):
        # Header CORS & No-Cache untuk development
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type, Range')
        self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
        self.send_header('Pragma', 'no-cache')
        self.send_header('Expires', '0')
        super().end_headers()

    def log_message(self, format, *args):
        # Format log yang rapi dan ringkas
        try:
            sys.stdout.write(f"  [HTTP] {self.address_string()} - {args[0]} {args[1]}\n")
            sys.stdout.flush()
        except Exception:
            pass


def find_available_port(start_port=8000, max_attempts=50):
    """Mencari port yang kosong jika port default sedang digunakan."""
    for port in range(start_port, start_port + max_attempts):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(('127.0.0.1', port)) != 0:
                return port
    return start_port


def main():
    parser = argparse.ArgumentParser(description='Jalankan Server Lokal AuditWorkpaper Pro')
    parser.add_argument('--port', '-p', type=int, default=8000, help='Port server (default: 8000)')
    parser.add_argument('--host', '-H', type=str, default='127.0.0.1', help='Host server (default: 127.0.0.1)')
    parser.add_argument('--no-browser', action='store_true', help='Jangan buka browser secara otomatis')
    args = parser.parse_args()

    port = find_available_port(args.port)
    server_address = (args.host, port)
    url = f"http://{args.host}:{port}/"

    os.chdir(PROJECT_DIR)

    try:
        httpd = HTTPServer(server_address, AuditWorkpaperHTTPHandler)
    except Exception as e:
        print(f"\n[!] Gagal memulai server pada port {port}: {e}")
        sys.exit(1)

    print("\n" + "=" * 60)
    print("  [>] AuditWorkpaper Pro — Server Python Aktif")
    print("=" * 60)
    print(f"  * Direktori  : {PROJECT_DIR}")
    print(f"  * Alamat URL : {url}")
    print("=" * 60)
    print("  Tekan Ctrl+C di terminal untuk menghentikan server.\n")

    if not args.no_browser:
        def open_browser():
            import time
            time.sleep(0.5)
            webbrowser.open(url)

        threading.Thread(target=open_browser, daemon=True).start()

    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n\n[-] Server dihentikan oleh pengguna.")
        httpd.server_close()
        sys.exit(0)


if __name__ == '__main__':
    main()

