"""
Platform-specific system trust store management.

Provides an abstract :class:`TrustStore` base class and concrete
implementations for macOS (Keychain) and Windows (certutil).
"""

from __future__ import annotations

import platform
import subprocess
from abc import ABC, abstractmethod
from pathlib import Path

from .paths import CA_COMMON_NAME


class TrustStore(ABC):
    """Abstract base class for system trust store operations."""

    @abstractmethod
    def install(self, ca_cert_path: Path) -> bool:
        """Install the CA certificate into the system trust store.

        Returns:
            *True* if installation succeeded.
        """

    @abstractmethod
    def uninstall(self, ca_cert_path: Path) -> bool:
        """Remove the CA certificate from the system trust store.

        Returns:
            *True* if removal succeeded.
        """

    @abstractmethod
    def is_installed(self) -> bool:
        """Check whether the CA is currently installed."""

    @abstractmethod
    def get_manual_install_command(self, ca_cert_path: Path) -> str:
        """Return a shell command the user can copy-paste to install manually."""

    @abstractmethod
    def get_manual_uninstall_command(self, ca_cert_path: Path) -> str:
        """Return a shell command the user can copy-paste to uninstall manually."""


class MacOSTrustStore(TrustStore):
    """macOS Keychain trust store."""

    _KEYCHAIN = "/Library/Keychains/System.keychain"

    def install(self, ca_cert_path: Path) -> bool:
        result = subprocess.run(
            [
                "sudo",
                "security",
                "add-trusted-cert",
                "-d",
                "-r",
                "trustRoot",
                "-k",
                self._KEYCHAIN,
                str(ca_cert_path),
            ],
            capture_output=False,
        )
        return result.returncode == 0

    def uninstall(self, ca_cert_path: Path) -> bool:
        result = subprocess.run(
            ["sudo", "security", "remove-trusted-cert", "-d", str(ca_cert_path)],
            capture_output=False,
        )
        return result.returncode == 0

    def is_installed(self) -> bool:
        result = subprocess.run(
            ["security", "find-certificate", "-c", CA_COMMON_NAME, self._KEYCHAIN],
            capture_output=True,
        )
        return result.returncode == 0

    def get_manual_install_command(self, ca_cert_path: Path) -> str:
        return (
            f"sudo security add-trusted-cert -d -r trustRoot "
            f"-k {self._KEYCHAIN} {ca_cert_path}"
        )

    def get_manual_uninstall_command(self, ca_cert_path: Path) -> str:
        return f"sudo security remove-trusted-cert -d {ca_cert_path}"


class WindowsTrustStore(TrustStore):
    """Windows certificate store (certutil)."""

    def install(self, ca_cert_path: Path) -> bool:
        result = subprocess.run(
            ["certutil", "-addstore", "Root", str(ca_cert_path)],
            capture_output=True,
        )
        return result.returncode == 0

    def uninstall(self, ca_cert_path: Path) -> bool:
        result = subprocess.run(
            ["certutil", "-delstore", "Root", CA_COMMON_NAME],
            capture_output=True,
        )
        return result.returncode == 0

    def is_installed(self) -> bool:
        result = subprocess.run(
            ["certutil", "-store", "Root", CA_COMMON_NAME],
            capture_output=True,
        )
        return result.returncode == 0

    def get_manual_install_command(self, ca_cert_path: Path) -> str:
        return f'certutil -addstore Root "{ca_cert_path}"'

    def get_manual_uninstall_command(self, ca_cert_path: Path) -> str:
        return f'certutil -delstore Root "{CA_COMMON_NAME}"'


class UnsupportedTrustStore(TrustStore):
    """Fallback for unsupported platforms (e.g. Linux)."""

    def install(self, ca_cert_path: Path) -> bool:
        return False

    def uninstall(self, ca_cert_path: Path) -> bool:
        return False

    def is_installed(self) -> bool:
        return False

    def get_manual_install_command(self, ca_cert_path: Path) -> str:
        return f"# Unsupported platform. Manually install {ca_cert_path} into your system trust store."

    def get_manual_uninstall_command(self, ca_cert_path: Path) -> str:
        return f"# Unsupported platform. Manually remove '{CA_COMMON_NAME}' from your system trust store."


def get_trust_store() -> TrustStore:
    """Return the appropriate :class:`TrustStore` for the current platform."""
    system = platform.system()
    if system == "Darwin":
        return MacOSTrustStore()
    if system == "Windows":
        return WindowsTrustStore()
    return UnsupportedTrustStore()
