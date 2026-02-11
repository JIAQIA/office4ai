"""
Certificate generation using the ``cryptography`` library.

Generates a local CA certificate and a server certificate for localhost
HTTPS connections. No external tools (e.g. mkcert) are required.
"""

from __future__ import annotations

import datetime
import ipaddress
import platform
import stat
from pathlib import Path

from cryptography import x509
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import rsa
from cryptography.x509.oid import ExtendedKeyUsageOID, NameOID

from .paths import (
    CA_CERT_FILE,
    CA_COMMON_NAME,
    CA_KEY_FILE,
    SERVER_CERT_FILE,
    SERVER_KEY_FILE,
)

# Validity periods
_CA_VALIDITY_DAYS = 365 * 10  # 10 years
_SERVER_VALIDITY_DAYS = 825  # macOS WKWebView upper limit


def _secure_directory(path: Path) -> None:
    """Ensure *path* exists as a directory with restricted permissions."""
    path.mkdir(parents=True, exist_ok=True)
    if platform.system() != "Windows":
        path.chmod(stat.S_IRWXU)  # 0700
    else:
        import subprocess

        subprocess.run(
            ["icacls", str(path), "/inheritance:r", "/grant:r", f"{_win_user()}:(OI)(CI)F"],
            check=True,
            capture_output=True,
        )


def _secure_file(path: Path) -> None:
    """Set owner-only read/write on *path*."""
    if platform.system() != "Windows":
        path.chmod(stat.S_IRUSR | stat.S_IWUSR)  # 0600
    else:
        import subprocess

        subprocess.run(
            ["icacls", str(path), "/inheritance:r", "/grant:r", f"{_win_user()}:F"],
            check=True,
            capture_output=True,
        )


def _win_user() -> str:
    """Return ``DOMAIN\\username`` for icacls on Windows."""
    import os

    return os.environ.get("USERDOMAIN", "") + "\\" + os.environ.get("USERNAME", "")


def generate_ca(cert_dir: Path) -> tuple[x509.Certificate, rsa.RSAPrivateKey]:
    """
    Generate a self-signed CA certificate.

    The CA includes **NameConstraints** limiting it to ``localhost``,
    ``127.0.0.1``, and ``::1`` so that even if the private key leaks
    an attacker cannot sign certificates for external domains.

    Args:
        cert_dir: Directory where ``ca.pem`` and ``ca-key.pem`` are written.

    Returns:
        Tuple of (CA certificate, CA private key).
    """
    _secure_directory(cert_dir)

    key = rsa.generate_private_key(public_exponent=65537, key_size=2048)

    now = datetime.datetime.now(datetime.timezone.utc)
    subject = issuer = x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, CA_COMMON_NAME)])

    cert = (
        x509.CertificateBuilder()
        .subject_name(subject)
        .issuer_name(issuer)
        .public_key(key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(now)
        .not_valid_after(now + datetime.timedelta(days=_CA_VALIDITY_DAYS))
        .add_extension(x509.BasicConstraints(ca=True, path_length=0), critical=True)
        .add_extension(
            x509.KeyUsage(
                digital_signature=False,
                content_commitment=False,
                key_encipherment=False,
                data_encipherment=False,
                key_agreement=False,
                key_cert_sign=True,
                crl_sign=True,
                encipher_only=False,
                decipher_only=False,
            ),
            critical=True,
        )
        .add_extension(
            x509.NameConstraints(
                permitted_subtrees=[
                    x509.DNSName("localhost"),
                    x509.IPAddress(ipaddress.IPv4Network("127.0.0.1/32")),
                    x509.IPAddress(ipaddress.IPv6Network("::1/128")),
                ],
                excluded_subtrees=None,
            ),
            critical=True,
        )
        .add_extension(x509.SubjectKeyIdentifier.from_public_key(key.public_key()), critical=False)
        .sign(key, hashes.SHA256())
    )

    # Write CA certificate
    ca_cert_path = cert_dir / CA_CERT_FILE
    ca_cert_path.write_bytes(cert.public_bytes(serialization.Encoding.PEM))

    # Write CA private key
    ca_key_path = cert_dir / CA_KEY_FILE
    ca_key_path.write_bytes(
        key.private_bytes(
            serialization.Encoding.PEM,
            serialization.PrivateFormat.TraditionalOpenSSL,
            serialization.NoEncryption(),
        )
    )
    _secure_file(ca_key_path)

    return cert, key


def generate_server_cert(
    cert_dir: Path,
    ca_cert: x509.Certificate,
    ca_key: rsa.RSAPrivateKey,
) -> x509.Certificate:
    """
    Generate a server certificate signed by the given CA.

    The certificate is valid for 825 days (macOS WKWebView limit) and
    includes SANs for ``localhost``, ``127.0.0.1``, and ``::1``.

    Args:
        cert_dir: Directory where ``cert.pem`` and ``key.pem`` are written.
        ca_cert: CA certificate used as issuer.
        ca_key: CA private key used to sign.

    Returns:
        The generated server certificate.
    """
    _secure_directory(cert_dir)

    key = rsa.generate_private_key(public_exponent=65537, key_size=2048)

    now = datetime.datetime.now(datetime.timezone.utc)
    subject = x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, "localhost")])

    cert = (
        x509.CertificateBuilder()
        .subject_name(subject)
        .issuer_name(ca_cert.subject)
        .public_key(key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(now)
        .not_valid_after(now + datetime.timedelta(days=_SERVER_VALIDITY_DAYS))
        .add_extension(x509.BasicConstraints(ca=False, path_length=None), critical=True)
        .add_extension(
            x509.KeyUsage(
                digital_signature=True,
                content_commitment=False,
                key_encipherment=True,
                data_encipherment=False,
                key_agreement=False,
                key_cert_sign=False,
                crl_sign=False,
                encipher_only=False,
                decipher_only=False,
            ),
            critical=True,
        )
        .add_extension(
            x509.ExtendedKeyUsage([ExtendedKeyUsageOID.SERVER_AUTH]),
            critical=False,
        )
        .add_extension(
            x509.SubjectAlternativeName(
                [
                    x509.DNSName("localhost"),
                    x509.IPAddress(ipaddress.IPv4Address("127.0.0.1")),
                    x509.IPAddress(ipaddress.IPv6Address("::1")),
                ]
            ),
            critical=False,
        )
        .add_extension(
            x509.AuthorityKeyIdentifier.from_issuer_public_key(ca_key.public_key()),
            critical=False,
        )
        .sign(ca_key, hashes.SHA256())
    )

    # Write server certificate
    server_cert_path = cert_dir / SERVER_CERT_FILE
    server_cert_path.write_bytes(cert.public_bytes(serialization.Encoding.PEM))

    # Write server private key
    server_key_path = cert_dir / SERVER_KEY_FILE
    server_key_path.write_bytes(
        key.private_bytes(
            serialization.Encoding.PEM,
            serialization.PrivateFormat.TraditionalOpenSSL,
            serialization.NoEncryption(),
        )
    )
    _secure_file(server_key_path)

    return cert
