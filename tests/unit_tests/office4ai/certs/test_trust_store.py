"""Tests for office4ai.certs.trust_store."""

from __future__ import annotations

import platform
from pathlib import Path
from unittest.mock import MagicMock, patch

from office4ai.certs.trust_store import (
    MacOSTrustStore,
    UnsupportedTrustStore,
    WindowsTrustStore,
    get_trust_store,
)


class TestGetTrustStore:
    def test_macos(self) -> None:
        with patch("office4ai.certs.trust_store.platform") as mock_platform:
            mock_platform.system.return_value = "Darwin"
            store = get_trust_store()
            assert isinstance(store, MacOSTrustStore)

    def test_windows(self) -> None:
        with patch("office4ai.certs.trust_store.platform") as mock_platform:
            mock_platform.system.return_value = "Windows"
            store = get_trust_store()
            assert isinstance(store, WindowsTrustStore)

    def test_linux_unsupported(self) -> None:
        with patch("office4ai.certs.trust_store.platform") as mock_platform:
            mock_platform.system.return_value = "Linux"
            store = get_trust_store()
            assert isinstance(store, UnsupportedTrustStore)


class TestMacOSTrustStore:
    def test_install_command(self) -> None:
        store = MacOSTrustStore()
        ca_path = Path("/tmp/ca.pem")
        with patch("office4ai.certs.trust_store.subprocess.run") as mock_run:
            mock_run.return_value = MagicMock(returncode=0)
            result = store.install(ca_path)
            assert result is True
            args = mock_run.call_args[0][0]
            assert args[0] == "sudo"
            assert "add-trusted-cert" in args

    def test_uninstall_command(self) -> None:
        store = MacOSTrustStore()
        ca_path = Path("/tmp/ca.pem")
        with patch("office4ai.certs.trust_store.subprocess.run") as mock_run:
            mock_run.return_value = MagicMock(returncode=0)
            result = store.uninstall(ca_path)
            assert result is True
            args = mock_run.call_args[0][0]
            assert "remove-trusted-cert" in args

    def test_is_installed(self) -> None:
        store = MacOSTrustStore()
        with patch("office4ai.certs.trust_store.subprocess.run") as mock_run:
            mock_run.return_value = MagicMock(returncode=0)
            assert store.is_installed() is True

            mock_run.return_value = MagicMock(returncode=1)
            assert store.is_installed() is False


class TestWindowsTrustStore:
    def test_install_command(self) -> None:
        store = WindowsTrustStore()
        ca_path = Path("C:\\certs\\ca.pem")
        with patch("office4ai.certs.trust_store.subprocess.run") as mock_run:
            mock_run.return_value = MagicMock(returncode=0)
            result = store.install(ca_path)
            assert result is True
            args = mock_run.call_args[0][0]
            assert args[0] == "certutil"
            assert "-addstore" in args

    def test_uninstall_command(self) -> None:
        store = WindowsTrustStore()
        ca_path = Path("C:\\certs\\ca.pem")
        with patch("office4ai.certs.trust_store.subprocess.run") as mock_run:
            mock_run.return_value = MagicMock(returncode=0)
            result = store.uninstall(ca_path)
            assert result is True
            args = mock_run.call_args[0][0]
            assert "certutil" in args
            assert "-delstore" in args


class TestUnsupportedTrustStore:
    def test_always_fails(self) -> None:
        store = UnsupportedTrustStore()
        ca_path = Path("/tmp/ca.pem")
        assert store.install(ca_path) is False
        assert store.uninstall(ca_path) is False
        assert store.is_installed() is False

    def test_manual_commands(self) -> None:
        store = UnsupportedTrustStore()
        ca_path = Path("/tmp/ca.pem")
        assert "Unsupported" in store.get_manual_install_command(ca_path)
        assert "Unsupported" in store.get_manual_uninstall_command(ca_path)


class TestCurrentPlatform:
    """Smoke test: get_trust_store() should work on the current platform."""

    def test_returns_a_trust_store(self) -> None:
        store = get_trust_store()
        if platform.system() == "Darwin":
            assert isinstance(store, MacOSTrustStore)
        elif platform.system() == "Windows":
            assert isinstance(store, WindowsTrustStore)
        else:
            assert isinstance(store, UnsupportedTrustStore)
