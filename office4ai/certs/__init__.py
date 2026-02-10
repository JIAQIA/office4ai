"""
Auto certificate management for Office4AI HTTPS Socket.IO connections.

Public API:
    - Path discovery: :func:`get_cert_dir`, :func:`get_cert_paths`, :func:`create_ssl_context`
    - Generation: :func:`generate_ca`, :func:`generate_server_cert`
    - Validation: :func:`validate_certs`, :class:`CertStatus`, :func:`get_cert_expiry_info`
    - Trust store: :func:`get_trust_store`
"""

from .generator import generate_ca, generate_server_cert
from .paths import create_ssl_context, get_cert_dir, get_cert_paths
from .trust_store import get_trust_store
from .validator import CertStatus, get_cert_expiry_info, validate_certs

__all__ = [
    "CertStatus",
    "create_ssl_context",
    "generate_ca",
    "generate_server_cert",
    "get_cert_dir",
    "get_cert_expiry_info",
    "get_cert_paths",
    "get_trust_store",
    "validate_certs",
]
