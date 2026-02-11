"""
Certificate validation utilities.

Checks certificate existence, parsability, and expiry status.
"""

from __future__ import annotations

import datetime
import enum
from pathlib import Path

from cryptography import x509

from .paths import CA_CERT_FILE, CA_KEY_FILE, SERVER_CERT_FILE, SERVER_KEY_FILE


class CertStatus(enum.Enum):
    """Result of :func:`validate_certs`."""

    NO_CERTS = "no_certs"
    CA_ONLY = "ca_only"
    ALL_VALID = "all_valid"
    SERVER_EXPIRED = "server_expired"
    CA_EXPIRED = "ca_expired"
    INVALID = "invalid"


def _load_cert(path: Path) -> x509.Certificate | None:
    """Load a PEM certificate, returning *None* on failure."""
    try:
        return x509.load_pem_x509_certificate(path.read_bytes())
    except Exception:
        return None


def _is_expired(cert: x509.Certificate) -> bool:
    now = datetime.datetime.now(datetime.timezone.utc)
    return now > cert.not_valid_after_utc


def validate_certs(cert_dir: Path) -> CertStatus:
    """
    Validate certificates in *cert_dir*.

    Returns:
        :class:`CertStatus` describing the current state.
    """
    ca_cert_path = cert_dir / CA_CERT_FILE
    ca_key_path = cert_dir / CA_KEY_FILE
    server_cert_path = cert_dir / SERVER_CERT_FILE
    server_key_path = cert_dir / SERVER_KEY_FILE

    # Check CA files
    if not ca_cert_path.exists() or not ca_key_path.exists():
        # If no CA files exist at all → NO_CERTS
        if not server_cert_path.exists() and not server_key_path.exists():
            return CertStatus.NO_CERTS
        # Server files without CA is invalid
        return CertStatus.INVALID

    # Try to parse CA cert
    ca_cert = _load_cert(ca_cert_path)
    if ca_cert is None:
        return CertStatus.INVALID

    # Check CA expiry
    if _is_expired(ca_cert):
        return CertStatus.CA_EXPIRED

    # Check server files
    if not server_cert_path.exists() or not server_key_path.exists():
        return CertStatus.CA_ONLY

    # Try to parse server cert
    server_cert = _load_cert(server_cert_path)
    if server_cert is None:
        return CertStatus.INVALID

    # Check server cert expiry
    if _is_expired(server_cert):
        return CertStatus.SERVER_EXPIRED

    return CertStatus.ALL_VALID


def get_cert_expiry_info(cert_dir: Path) -> dict[str, str]:
    """
    Return human-readable expiry information.

    Returns:
        Dictionary with optional keys ``ca_valid_until`` and ``server_valid_until``.
    """
    info: dict[str, str] = {}

    ca_cert = _load_cert(cert_dir / CA_CERT_FILE)
    if ca_cert is not None:
        info["ca_valid_until"] = ca_cert.not_valid_after_utc.strftime("%Y-%m-%d")

    server_cert = _load_cert(cert_dir / SERVER_CERT_FILE)
    if server_cert is not None:
        info["server_valid_until"] = server_cert.not_valid_after_utc.strftime("%Y-%m-%d")

    return info
