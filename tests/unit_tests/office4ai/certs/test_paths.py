"""Tests for office4ai.certs.paths."""

from __future__ import annotations

from pathlib import Path

import pytest

from office4ai.certs.paths import get_cert_dir


class TestGetCertDir:
    def test_default_path(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.delenv("OFFICE4AI_CERT_DIR", raising=False)
        result = get_cert_dir()
        assert result == Path.home() / ".office4ai" / "certs"

    def test_env_override(self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
        monkeypatch.setenv("OFFICE4AI_CERT_DIR", str(tmp_path / "custom"))
        result = get_cert_dir()
        assert result == (tmp_path / "custom").resolve()


class TestCreateSslContext:
    def test_missing_cert_raises(self, tmp_path: Path) -> None:
        from office4ai.certs.paths import create_ssl_context

        with pytest.raises(FileNotFoundError, match="Server certificate not found"):
            create_ssl_context(tmp_path)

    def test_missing_key_raises(self, tmp_path: Path) -> None:
        from office4ai.certs.paths import create_ssl_context

        # Create cert.pem but not key.pem
        (tmp_path / "cert.pem").write_text("dummy")
        with pytest.raises(FileNotFoundError, match="Server key not found"):
            create_ssl_context(tmp_path)
