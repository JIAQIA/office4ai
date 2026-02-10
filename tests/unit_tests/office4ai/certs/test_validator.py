"""Tests for office4ai.certs.validator."""

from __future__ import annotations

import datetime
from pathlib import Path

from cryptography import x509
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import rsa
from cryptography.x509.oid import NameOID

from office4ai.certs.generator import generate_ca, generate_server_cert
from office4ai.certs.paths import CA_CERT_FILE, CA_KEY_FILE, SERVER_CERT_FILE, SERVER_KEY_FILE
from office4ai.certs.validator import CertStatus, get_cert_expiry_info, validate_certs


def _write_expired_cert(path: Path, cn: str = "test", *, ca: bool = False) -> None:
    """Write a PEM certificate that is already expired."""
    key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    now = datetime.datetime.now(datetime.timezone.utc)
    builder = (
        x509.CertificateBuilder()
        .subject_name(x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, cn)]))
        .issuer_name(x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, cn)]))
        .public_key(key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(now - datetime.timedelta(days=10))
        .not_valid_after(now - datetime.timedelta(days=1))
        .add_extension(x509.BasicConstraints(ca=ca, path_length=0 if ca else None), critical=True)
    )
    cert = builder.sign(key, hashes.SHA256())
    path.write_bytes(cert.public_bytes(serialization.Encoding.PEM))
    # Write key too (needed for CA validation)
    key_path = path.parent / (path.stem.replace("ca", "ca-key").replace("cert", "key") + path.suffix)
    if "ca" in path.name:
        key_path = path.parent / CA_KEY_FILE
    else:
        key_path = path.parent / SERVER_KEY_FILE
    key_path.write_bytes(
        key.private_bytes(
            serialization.Encoding.PEM,
            serialization.PrivateFormat.TraditionalOpenSSL,
            serialization.NoEncryption(),
        )
    )


class TestValidateCerts:
    def test_no_certs(self, tmp_path: Path) -> None:
        assert validate_certs(tmp_path) == CertStatus.NO_CERTS

    def test_all_valid(self, tmp_path: Path) -> None:
        ca_cert, ca_key = generate_ca(tmp_path)
        generate_server_cert(tmp_path, ca_cert, ca_key)
        assert validate_certs(tmp_path) == CertStatus.ALL_VALID

    def test_ca_only(self, tmp_path: Path) -> None:
        generate_ca(tmp_path)
        assert validate_certs(tmp_path) == CertStatus.CA_ONLY

    def test_server_expired(self, tmp_path: Path) -> None:
        ca_cert, ca_key = generate_ca(tmp_path)
        generate_server_cert(tmp_path, ca_cert, ca_key)
        # Overwrite server cert with expired one
        _write_expired_cert(tmp_path / SERVER_CERT_FILE, cn="localhost")
        assert validate_certs(tmp_path) == CertStatus.SERVER_EXPIRED

    def test_ca_expired(self, tmp_path: Path) -> None:
        # Generate valid set first so all files exist
        ca_cert, ca_key = generate_ca(tmp_path)
        generate_server_cert(tmp_path, ca_cert, ca_key)
        # Overwrite CA cert with expired one
        _write_expired_cert(tmp_path / CA_CERT_FILE, cn="Office4AI Local CA", ca=True)
        assert validate_certs(tmp_path) == CertStatus.CA_EXPIRED

    def test_invalid_unparseable(self, tmp_path: Path) -> None:
        (tmp_path / CA_CERT_FILE).write_text("not-a-cert")
        (tmp_path / CA_KEY_FILE).write_text("not-a-key")
        assert validate_certs(tmp_path) == CertStatus.INVALID

    def test_invalid_server_without_ca(self, tmp_path: Path) -> None:
        (tmp_path / SERVER_CERT_FILE).write_text("dummy")
        (tmp_path / SERVER_KEY_FILE).write_text("dummy")
        assert validate_certs(tmp_path) == CertStatus.INVALID


class TestGetCertExpiryInfo:
    def test_returns_dates(self, tmp_path: Path) -> None:
        ca_cert, ca_key = generate_ca(tmp_path)
        generate_server_cert(tmp_path, ca_cert, ca_key)
        info = get_cert_expiry_info(tmp_path)
        assert "ca_valid_until" in info
        assert "server_valid_until" in info

    def test_empty_for_no_certs(self, tmp_path: Path) -> None:
        info = get_cert_expiry_info(tmp_path)
        assert info == {}
