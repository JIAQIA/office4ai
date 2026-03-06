"""
Centralized logging configuration for Office4AI.

Bridges stdlib logging → loguru so that all loggers (including third-party
libraries like socketio, engineio, aiohttp, uvicorn) are unified under a
single format and routed to both console and file sinks.

Environment variables:
    OFFICE4AI_LOG_DIR     – directory for log files (default: "./logs", empty string disables file logging)
    OFFICE4AI_LOG_LEVEL   – minimum log level (default: "INFO")
    OFFICE4AI_LOG_CONSOLE – enable console output (default: "true")
"""

from __future__ import annotations

import logging
import os
import sys

from loguru import logger

# Shared format string (no color tags – loguru adds color automatically for console sinks)
_LOG_FORMAT = "{time:YYYY-MM-DD HH:mm:ss.SSS} | {level: <8} | {name}:{function}:{line} - {message}"

# Third-party loggers to intercept
_INTERCEPTED_LOGGERS = ("socketio", "engineio", "aiohttp", "uvicorn", "uvicorn.access", "uvicorn.error")


class InterceptHandler(logging.Handler):
    """Route stdlib logging records to loguru, preserving caller information."""

    def emit(self, record: logging.LogRecord) -> None:
        # Map stdlib level to loguru level name
        try:
            level: str | int = logger.level(record.levelname).name
        except ValueError:
            level = record.levelno

        # Find caller frame outside the logging/loguru stack
        frame = sys._getframe(6)
        depth = 6
        while frame and frame.f_code.co_filename == logging.__file__:
            frame = frame.f_back  # type: ignore[assignment]
            depth += 1

        logger.opt(depth=depth, exception=record.exc_info).log(level, record.getMessage())


def setup_logging(
    log_dir: str | None = None,
    log_level: str | None = None,
    console: bool | None = None,
) -> None:
    """
    Configure loguru sinks and bridge stdlib logging.

    Parameters take effect only when the corresponding environment variable
    is **not** set – env vars always win.

    Args:
        log_dir: Directory for log files. ``""`` disables file logging.
        log_level: Minimum log level (e.g. ``"DEBUG"``, ``"INFO"``).
        console: Whether to emit to stderr with colors.
    """
    # --- resolve effective values (env var > argument > default) ---
    effective_dir = os.environ.get("OFFICE4AI_LOG_DIR", log_dir if log_dir is not None else "./logs")
    effective_level = os.environ.get("OFFICE4AI_LOG_LEVEL", log_level if log_level is not None else "INFO").upper()
    console_raw = os.environ.get("OFFICE4AI_LOG_CONSOLE")
    if console_raw is not None:
        effective_console = console_raw.lower() in ("1", "true", "yes")
    else:
        effective_console = console if console is not None else True

    # --- reset loguru ---
    logger.remove()

    # --- console sink ---
    if effective_console:
        logger.add(
            sys.stderr,
            level=effective_level,
            format=_LOG_FORMAT,
            colorize=True,
        )

    # --- file sink ---
    if effective_dir:
        logger.add(
            os.path.join(effective_dir, "office4ai-mcp_{time:YYYY-MM-DD}.log"),
            level=effective_level,
            format=_LOG_FORMAT,
            rotation="00:00",
            retention="3 days",
            enqueue=True,
            colorize=False,
        )

    # --- bridge stdlib logging → loguru ---
    logging.basicConfig(handlers=[InterceptHandler()], level=0, force=True)

    for name in _INTERCEPTED_LOGGERS:
        lib_logger = logging.getLogger(name)
        lib_logger.handlers = [InterceptHandler()]
        lib_logger.propagate = False
