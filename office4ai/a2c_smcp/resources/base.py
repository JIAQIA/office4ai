# filename: base.py
# @Time    : 2025/12/18 16:07
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm

from __future__ import annotations

from abc import ABC, abstractmethod
from urllib.parse import parse_qs, urlparse

from loguru import logger


def parse_window_uri_params(
    uri: str,
    current_priority: int,
    current_fullscreen: bool,
    log_prefix: str = "Window resource",
) -> tuple[int, bool]:
    """Parse priority and fullscreen from a window:// URI query string.

    Returns (new_priority, new_fullscreen) with unchanged values for invalid/missing params.
    """
    parsed = urlparse(uri)
    params = parse_qs(parsed.query)

    priority = current_priority
    fullscreen = current_fullscreen

    if "priority" in params:
        try:
            val = int(params["priority"][0])
            if 0 <= val <= 100:
                if val != priority:
                    logger.debug(f"{log_prefix} priority: {priority} -> {val}")
                    priority = val
            else:
                logger.warning(f"Invalid priority value in URI: {val}, must be in [0, 100]")
        except (ValueError, IndexError) as e:
            logger.warning(f"Failed to parse priority from URI: {e}")

    if "fullscreen" in params:
        try:
            fs_str = params["fullscreen"][0].lower()
            new_fs: bool | None = None
            if fs_str in {"true", "1", "yes", "on"}:
                new_fs = True
            elif fs_str in {"false", "0", "no", "off"}:
                new_fs = False
            else:
                logger.warning(f"Invalid fullscreen value in URI: {fs_str}, ignoring")

            if new_fs is not None and new_fs != fullscreen:
                logger.debug(f"{log_prefix} fullscreen: {fullscreen} -> {new_fs}")
                fullscreen = new_fs
        except (IndexError, AttributeError) as e:
            logger.warning(f"Failed to parse fullscreen from URI: {e}")

    return priority, fullscreen


class BaseResource(ABC):
    @property
    @abstractmethod
    def uri(self) -> str:  # pragma: no cover
        raise NotImplementedError

    @property
    @abstractmethod
    def base_uri(self) -> str:  # pragma: no cover
        raise NotImplementedError

    @property
    @abstractmethod
    def name(self) -> str:  # pragma: no cover
        raise NotImplementedError

    @property
    @abstractmethod
    def description(self) -> str:  # pragma: no cover
        raise NotImplementedError

    @property
    @abstractmethod
    def mime_type(self) -> str:  # pragma: no cover
        raise NotImplementedError

    def update_from_uri(self, uri: str) -> None:
        return

    @abstractmethod
    async def read(self) -> str:  # pragma: no cover
        raise NotImplementedError
