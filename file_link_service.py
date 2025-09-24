"""Signed URL generator and optional HTTP server for boleto PDFs."""

from __future__ import annotations

import argparse
import base64
import hmac
import hashlib
import http.server
import json
import logging
import os
import threading
import time
from pathlib import Path
from typing import Optional
from urllib.parse import parse_qs, quote, urlencode, urlparse, urlunparse


LOGGER = logging.getLogger(__name__)


class FileLinkService:
    """Generates signed URLs to serve files through a lightweight HTTP endpoint."""

    def __init__(
        self,
        downloads_dir: Path,
        base_url: str,
        secret_key: str,
        expiry_minutes: int,
    ) -> None:
        self.downloads_dir = downloads_dir
        self.base_url = base_url.rstrip('/')
        self.secret_key = secret_key.encode('utf-8')
        self.expiry_minutes = expiry_minutes

    def _relative_path(self, file_path: Path) -> Path:
        file_path = file_path.resolve()
        try:
            return file_path.relative_to(self.downloads_dir.resolve())
        except ValueError:
            raise ValueError("File is outside of the configured downloads directory")

    def generate_signed_url(self, file_path: Path) -> str:
        relative = self._relative_path(file_path)
        expires = int(time.time() + self.expiry_minutes * 60)
        payload = f"{relative}|{expires}"
        signature = hmac.new(self.secret_key, payload.encode('utf-8'), hashlib.sha256).hexdigest()

        # Build query parameters
        query = urlencode(
            {
                "path": str(relative),
                "expires": str(expires),
                "sig": signature,
            }
        )

        parsed = urlparse(self.base_url)
        url = urlunparse(
            (
                parsed.scheme,
                parsed.netloc,
                parsed.path,
                parsed.params,
                query,
                parsed.fragment,
            )
        )
        return url

    def validate_request(self, path: str, expires: int, sig: str) -> Optional[Path]:
        payload = f"{path}|{expires}"
        expected = hmac.new(self.secret_key, payload.encode('utf-8'), hashlib.sha256).hexdigest()
        if not hmac.compare_digest(expected, sig):
            LOGGER.warning("Invalid signature for %s", path)
            return None
        if time.time() > expires:
            LOGGER.warning("Link expired for %s", path)
            return None
        file_path = (self.downloads_dir / Path(path)).resolve()
        try:
            file_path.relative_to(self.downloads_dir.resolve())
        except ValueError:
            LOGGER.warning("Attempt to access file outside downloads directory: %s", file_path)
            return None
        if not file_path.exists() or not file_path.is_file():
            LOGGER.warning("Requested file not found: %s", file_path)
            return None
        return file_path


class SignedFileRequestHandler(http.server.SimpleHTTPRequestHandler):
    """HTTP handler that validates signed URLs before serving files."""

    service: FileLinkService

    def do_GET(self):  # noqa: N802 (follow BaseHTTPRequestHandler naming)
        parsed = urlparse(self.path)
        params = parse_qs(parsed.query)
        path = params.get("path", [None])[0]
        expires = params.get("expires", [None])[0]
        sig = params.get("sig", [None])[0]

        if not all([path, expires, sig]):
            self.send_error(400, "Missing signature parameters")
            return

        try:
            expires_int = int(expires)
        except ValueError:
            self.send_error(400, "Invalid expires parameter")
            return

        file_path = self.service.validate_request(path, expires_int, sig)
        if not file_path:
            self.send_error(403, "Invalid or expired link")
            return

        self.path = "/" + str(file_path.relative_to(self.service.downloads_dir))
        return http.server.SimpleHTTPRequestHandler.do_GET(self)


def run_server(
    downloads_dir: Path,
    host: str,
    port: int,
    secret_key: str,
    debug: bool = False,
):
    logging.basicConfig(level=logging.DEBUG if debug else logging.INFO)
    service = FileLinkService(downloads_dir, base_url="", secret_key=secret_key, expiry_minutes=30)

    handler_class = SignedFileRequestHandler
    handler_class.service = service
    os.chdir(str(downloads_dir))

    with http.server.ThreadingHTTPServer((host, port), handler_class) as httpd:
        LOGGER.info("Serving signed files from %s on %s:%s", downloads_dir, host, port)
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            LOGGER.info("Shutting down signed file server")


def main():
    parser = argparse.ArgumentParser(description="Signed file server for boleto PDFs")
    parser.add_argument("downloads_dir", help="Directory containing downloadable files")
    parser.add_argument("secret_key", help="Secret key for HMAC validation")
    parser.add_argument("--host", default="0.0.0.0")
    parser.add_argument("--port", type=int, default=8080)
    parser.add_argument("--debug", action="store_true")

    args = parser.parse_args()
    run_server(Path(args.downloads_dir), args.host, args.port, args.secret_key, debug=args.debug)


if __name__ == "__main__":
    main()
