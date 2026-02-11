"""Tests for office4ai.certs.generator."""

from __future__ import annotations

import datetime
import ipaddress
import platform
import stat
from pathlib import Path

from cryptography import x509
from cryptography.x509.oid import ExtendedKeyUsageOID, NameOID

from office4ai.certs.generator import generate_ca, generate_server_cert
from office4ai.certs.paths import CA_CERT_FILE, CA_KEY_FILE, SERVER_CERT_FILE, SERVER_KEY_FILE


class TestGenerateCA:
    def test_generates_four_files(self, tmp_path: Path) -> None:
        cert_dir = tmp_path / "certs"
        generate_ca(cert_dir)
        assert (cert_dir / CA_CERT_FILE).exists()
        assert (cert_dir / CA_KEY_FILE).exists()

    def test_ca_attributes(self, tmp_path: Path) -> None:
        cert_dir = tmp_path / "certs"
        ca_cert, _ = generate_ca(cert_dir)

        # CN
        cn = ca_cert.subject.get_attributes_for_oid(NameOID.COMMON_NAME)[0].value
        assert cn == "Office4AI Local CA"

        # Self-signed
        assert ca_cert.issuer == ca_cert.subject

        # BasicConstraints CA:TRUE pathlen:0
        bc = ca_cert.extensions.get_extension_for_class(x509.BasicConstraints).value
        assert bc.ca is True
        assert bc.path_length == 0

        # KeyUsage: keyCertSign + cRLSign
        ku = ca_cert.extensions.get_extension_for_class(x509.KeyUsage).value
        assert ku.key_cert_sign is True
        assert ku.crl_sign is True
        assert ku.digital_signature is False

        # NameConstraints: localhost, 127.0.0.1/32, ::1/128
        nc = ca_cert.extensions.get_extension_for_class(x509.NameConstraints).value
        assert nc.permitted_subtrees is not None
        dns_names = [s.value for s in nc.permitted_subtrees if isinstance(s, x509.DNSName)]
        assert "localhost" in dns_names
        ip_nets = [s.value for s in nc.permitted_subtrees if isinstance(s, x509.IPAddress)]
        assert ipaddress.IPv4Network("127.0.0.1/32") in ip_nets
        assert ipaddress.IPv6Network("::1/128") in ip_nets

    def test_ca_validity_10_years(self, tmp_path: Path) -> None:
        cert_dir = tmp_path / "certs"
        ca_cert, _ = generate_ca(cert_dir)

        now = datetime.datetime.now(datetime.timezone.utc)
        delta = ca_cert.not_valid_after_utc - now
        # Should be approximately 3650 days (10 years)
        assert 3640 <= delta.days <= 3660

    def test_private_key_permissions_unix(self, tmp_path: Path) -> None:
        if platform.system() == "Windows":
            return
        cert_dir = tmp_path / "certs"
        generate_ca(cert_dir)
        key_stat = (cert_dir / CA_KEY_FILE).stat()
        # Owner read+write only (0600)
        assert stat.S_IMODE(key_stat.st_mode) == 0o600


class TestGenerateServerCert:
    def test_generates_server_files(self, tmp_path: Path) -> None:
        cert_dir = tmp_path / "certs"
        ca_cert, ca_key = generate_ca(cert_dir)
        generate_server_cert(cert_dir, ca_cert, ca_key)
        assert (cert_dir / SERVER_CERT_FILE).exists()
        assert (cert_dir / SERVER_KEY_FILE).exists()

    def test_server_cert_attributes(self, tmp_path: Path) -> None:
        cert_dir = tmp_path / "certs"
        ca_cert, ca_key = generate_ca(cert_dir)
        server_cert = generate_server_cert(cert_dir, ca_cert, ca_key)

        # CN=localhost
        cn = server_cert.subject.get_attributes_for_oid(NameOID.COMMON_NAME)[0].value
        assert cn == "localhost"

        # Issuer = CA
        assert server_cert.issuer == ca_cert.subject

        # Not a CA
        bc = server_cert.extensions.get_extension_for_class(x509.BasicConstraints).value
        assert bc.ca is False

        # SAN
        san = server_cert.extensions.get_extension_for_class(x509.SubjectAlternativeName).value
        dns_names = san.get_values_for_type(x509.DNSName)
        assert "localhost" in dns_names
        ips = san.get_values_for_type(x509.IPAddress)
        assert ipaddress.IPv4Address("127.0.0.1") in ips
        assert ipaddress.IPv6Address("::1") in ips

        # ExtendedKeyUsage: serverAuth
        eku = server_cert.extensions.get_extension_for_class(x509.ExtendedKeyUsage).value
        assert ExtendedKeyUsageOID.SERVER_AUTH in eku

    def test_server_cert_validity_825_days(self, tmp_path: Path) -> None:
        cert_dir = tmp_path / "certs"
        ca_cert, ca_key = generate_ca(cert_dir)
        server_cert = generate_server_cert(cert_dir, ca_cert, ca_key)

        now = datetime.datetime.now(datetime.timezone.utc)
        delta = server_cert.not_valid_after_utc - now
        assert 820 <= delta.days <= 830

    def test_server_key_permissions_unix(self, tmp_path: Path) -> None:
        if platform.system() == "Windows":
            return
        cert_dir = tmp_path / "certs"
        ca_cert, ca_key = generate_ca(cert_dir)
        generate_server_cert(cert_dir, ca_cert, ca_key)
        key_stat = (cert_dir / SERVER_KEY_FILE).stat()
        assert stat.S_IMODE(key_stat.st_mode) == 0o600
