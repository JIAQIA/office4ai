"""
Certificate path discovery and SSL context creation.

Provides constants and utilities for locating certificate files
and creating SSL contexts for the HTTPS Socket.IO server.
"""

from __future__ import annotations

import os
import ssl
from pathlib import Path

# Certificate file names
CA_CERT_FILE = "ca.pem"
CA_KEY_FILE = "ca-key.pem"
SERVER_CERT_FILE = "cert.pem"
SERVER_KEY_FILE = "key.pem"

# CA subject
CA_COMMON_NAME = "Office4AI Local CA"

# Environment variable for custom cert directory
CERT_DIR_ENV = "OFFICE4AI_CERT_DIR"


def get_cert_dir() -> Path:
    """
    Get the certificate directory path.

    Priority:
    1. ``OFFICE4AI_CERT_DIR`` environment variable
    2. ``~/.office4ai/certs/`` (default)

    Returns:
        Absolute path to the certificate directory.
    """
    env_dir = os.environ.get(CERT_DIR_ENV)
    if env_dir:
        return Path(env_dir).resolve()
    return Path.home() / ".office4ai" / "certs"


def get_cert_paths(cert_dir: Path | None = None) -> dict[str, Path]:
    """
    Return paths for all four certificate files.

    Args:
        cert_dir: Certificate directory. Uses :func:`get_cert_dir` when *None*.

    Returns:
        Dictionary with keys ``ca_cert``, ``ca_key``, ``server_cert``, ``server_key``.
    """
    if cert_dir is None:
        cert_dir = get_cert_dir()
    return {
        "ca_cert": cert_dir / CA_CERT_FILE,
        "ca_key": cert_dir / CA_KEY_FILE,
        "server_cert": cert_dir / SERVER_CERT_FILE,
        "server_key": cert_dir / SERVER_KEY_FILE,
    }


def create_ssl_context(cert_dir: Path | None = None) -> ssl.SSLContext:
    """
    Create an SSL context using server certificate and key.

    Args:
        cert_dir: Certificate directory. Uses :func:`get_cert_dir` when *None*.

    Returns:
        Configured :class:`ssl.SSLContext` for TLS server use.

    Raises:
        FileNotFoundError: If cert.pem or key.pem is missing.
    """
    if cert_dir is None:
        cert_dir = get_cert_dir()

    cert_path = cert_dir / SERVER_CERT_FILE
    key_path = cert_dir / SERVER_KEY_FILE

    if not cert_path.exists():
        raise FileNotFoundError(f"Server certificate not found: {cert_path}")
    if not key_path.exists():
        raise FileNotFoundError(f"Server key not found: {key_path}")

    ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    ctx.load_cert_chain(str(cert_path), str(key_path))
    return ctx
