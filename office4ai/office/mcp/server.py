# filename: server.py
# @Time    : 2025/12/18 16:07
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm

from __future__ import annotations

import asyncio
import shutil
import sys
from pathlib import Path

from loguru import logger

from office4ai.a2c_smcp.config import MCPServerConfig
from office4ai.a2c_smcp.server import BaseMCPServer
from office4ai.certs.paths import get_cert_dir
from office4ai.certs.trust_store import get_trust_store
from office4ai.certs.validator import CertStatus, get_cert_expiry_info, validate_certs
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


class OfficeMCPServer(BaseMCPServer):
    """
    Office 级别的统一 MCP Server | Office-level unified MCP Server

    一个 Server 实例同时处理 Word、PPT、Excel 三种文档类型。
    Socket.IO Server 的生命周期与 MCP Server 一致。
    """

    def __init__(self, config: MCPServerConfig, cert_dir: Path | None = None) -> None:
        # 同步：创建 workspace 实例 (未启动)
        self.workspace = OfficeWorkspace(
            host=config.host,
            port=config.socketio_port,
            cert_dir=cert_dir,
        )
        super().__init__(config=config, server_name="office4ai")

    def _register_tools(self) -> None:
        """注册所有平台的工具 | Register all platform tools"""
        from office4ai.a2c_smcp.tools.word import (
            WordAppendTextTool,
            WordGetSelectedContentTool,
            WordGetVisibleContentTool,
            WordInsertEquationTool,
            WordInsertImageTool,
            WordInsertTableTool,
            WordInsertTextTool,
            WordInsertTOCTool,
            WordReplaceTextTool,
        )

        word_tools = [
            WordGetSelectedContentTool(self.workspace),
            WordGetVisibleContentTool(self.workspace),
            WordInsertTextTool(self.workspace),
            WordAppendTextTool(self.workspace),
            WordReplaceTextTool(self.workspace),
            WordInsertImageTool(self.workspace),
            WordInsertTableTool(self.workspace),
            WordInsertEquationTool(self.workspace),
            WordInsertTOCTool(self.workspace),
        ]

        for tool in word_tools:
            self.tools[tool.name] = tool

        logger.info(f"已注册 {len(word_tools)} 个 Word 工具 | Registered {len(word_tools)} Word tools")

        # PPT 工具 (未来) | PPT tools (future)
        # Excel 工具 (未来) | Excel tools (future)

    def _register_resources(self) -> None:
        """注册资源 | Register resources"""
        from office4ai.a2c_smcp.resources.connected_documents import ConnectedDocumentsResource

        docs_resource = ConnectedDocumentsResource(self.workspace)
        self.resources[docs_resource.base_uri] = docs_resource

    async def _async_startup(self) -> None:
        """启动 OfficeWorkspace (Socket.IO Server) | Start OfficeWorkspace"""
        logger.info("启动 OfficeWorkspace | Starting OfficeWorkspace...")
        await self.workspace.start()

    async def _async_shutdown(self) -> None:
        """停止 OfficeWorkspace (Socket.IO Server) | Stop OfficeWorkspace"""
        logger.info("停止 OfficeWorkspace | Stopping OfficeWorkspace...")
        await self.workspace.stop()


# ---------------------------------------------------------------------------
# CLI sub-commands
# ---------------------------------------------------------------------------


def _cmd_setup() -> None:
    """Generate certificates and install CA into system trust store."""
    from office4ai.certs.generator import generate_ca, generate_server_cert

    cert_dir = get_cert_dir()
    status = validate_certs(cert_dir)

    if status == CertStatus.ALL_VALID:
        info = get_cert_expiry_info(cert_dir)
        print(f"Certificates are already valid at {cert_dir}")
        print(f"  CA valid until:     {info.get('ca_valid_until', 'N/A')}")
        print(f"  Server valid until: {info.get('server_valid_until', 'N/A')}")
        return

    # Determine what needs to happen
    need_ca = status in (CertStatus.NO_CERTS, CertStatus.CA_EXPIRED, CertStatus.INVALID)
    need_server = True  # always regenerate server cert when setup runs

    print()
    print("Office4AI Certificate Setup")
    print("=" * 40)
    if need_ca:
        print("This will:")
        print("  1. Generate a local CA certificate (Name Constraints: localhost/127.0.0.1 only)")
        print("  2. Generate a server certificate for localhost and 127.0.0.1")
        print("  3. Install the CA into your system trust store (requires admin privileges)")
    else:
        print("CA certificate is still valid. This will:")
        print("  1. Regenerate the server certificate (no admin privileges needed)")
    print(f"\nCertificate location: {cert_dir}")
    print()

    answer = input("Proceed? [y/N]: ").strip().lower()
    if answer != "y":
        print("Aborted.")
        return

    if need_ca:
        ca_cert, ca_key = generate_ca(cert_dir)
        print("CA certificate generated (valid 10 years)")
    else:
        # Load existing CA
        from cryptography import x509
        from cryptography.hazmat.primitives.serialization import load_pem_private_key

        from office4ai.certs.paths import CA_CERT_FILE, CA_KEY_FILE

        ca_cert = x509.load_pem_x509_certificate((cert_dir / CA_CERT_FILE).read_bytes())
        ca_key = load_pem_private_key((cert_dir / CA_KEY_FILE).read_bytes(), password=None)  # type: ignore[assignment]

    if need_server:
        generate_server_cert(cert_dir, ca_cert, ca_key)
        print("Server certificate generated (valid 825 days)")

    # Install CA to system trust store (only if CA was regenerated)
    if need_ca:
        from office4ai.certs.paths import CA_CERT_FILE

        ca_cert_path = cert_dir / CA_CERT_FILE
        trust_store = get_trust_store()

        print("Installing CA to system trust store...")
        if trust_store.install(ca_cert_path):
            print("CA installed to system trust store")
        else:
            print("Failed to install CA to system trust store.")
            print("You can install it manually:")
            print(f"  {trust_store.get_manual_install_command(ca_cert_path)}")

    print()
    print("Setup complete! Start the server with: office4ai-mcp serve")


def _cmd_cleanup() -> None:
    """Remove CA from system trust store and delete certificate files."""
    cert_dir = get_cert_dir()

    if not cert_dir.exists():
        print(f"No certificates found at {cert_dir}")
        return

    print()
    print("Office4AI Certificate Cleanup")
    print("=" * 40)
    print("This will:")
    print("  1. Remove CA from system trust store (requires admin privileges)")
    print(f"  2. Delete all certificate files from {cert_dir}")
    print()

    answer = input("Proceed? [y/N]: ").strip().lower()
    if answer != "y":
        print("Aborted.")
        return

    from office4ai.certs.paths import CA_CERT_FILE

    ca_cert_path = cert_dir / CA_CERT_FILE
    trust_store = get_trust_store()

    if ca_cert_path.exists():
        if trust_store.uninstall(ca_cert_path):
            print("CA removed from system trust store")
        else:
            print("Failed to remove CA from system trust store.")
            print("You can remove it manually:")
            print(f"  {trust_store.get_manual_uninstall_command(ca_cert_path)}")

    shutil.rmtree(cert_dir)
    print("Certificate files deleted")
    print()
    print("Cleanup complete")


async def async_main() -> None:
    cert_dir = get_cert_dir()
    status = validate_certs(cert_dir)

    if status != CertStatus.ALL_VALID:
        if status == CertStatus.NO_CERTS:
            logger.error(f"SSL certificates not found at {cert_dir}")
        elif status == CertStatus.SERVER_EXPIRED:
            info = get_cert_expiry_info(cert_dir)
            logger.error(f"Server certificate expired ({info.get('server_valid_until', 'unknown')})")
        elif status == CertStatus.CA_EXPIRED:
            info = get_cert_expiry_info(cert_dir)
            logger.error(f"CA certificate expired ({info.get('ca_valid_until', 'unknown')})")
        else:
            logger.error(f"Certificate validation failed: {status.value}")

        logger.error("Run `office4ai-mcp setup` to generate and install certificates.")
        sys.exit(1)

    info = get_cert_expiry_info(cert_dir)
    logger.info(f"Loaded certificates from {cert_dir}")
    logger.info(f"  CA valid until:     {info.get('ca_valid_until', 'N/A')}")
    logger.info(f"  Server valid until: {info.get('server_valid_until', 'N/A')}")

    config = MCPServerConfig()
    logger.info(
        f"启动 MCP Server | Starting MCP Server: transport={config.transport}, host={config.host}, port={config.port}",
    )

    server = OfficeMCPServer(config, cert_dir=cert_dir)
    await server.run()


def main() -> None:
    # Extract subcommand before confz sees sys.argv
    # Use simple argv inspection: first non-flag arg is the subcommand
    args = sys.argv[1:]
    subcommand = "serve"

    if args and args[0] in ("serve", "setup", "cleanup"):
        subcommand = args[0]
        sys.argv = [sys.argv[0]] + args[1:]

    if subcommand == "setup":
        _cmd_setup()
    elif subcommand == "cleanup":
        _cmd_cleanup()
    else:
        asyncio.run(async_main())


if __name__ == "__main__":
    main()
