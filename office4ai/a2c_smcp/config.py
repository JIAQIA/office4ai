# filename: config.py
# @Time    : 2025/12/18 16:07
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm

from __future__ import annotations

from typing import Literal

from confz import BaseConfig, CLArgSource, EnvSource
from confz.base_config import BaseConfigMetaclass
from pydantic import Field, field_validator


class MCPServerConfig(BaseConfig, metaclass=BaseConfigMetaclass):
    CONFIG_SOURCES = [
        EnvSource(
            allow_all=True,
            prefix="",
            remap={
                "TRANSPORT": "transport",
                "HOST": "host",
                "PORT": "port",
            },
        ),
        CLArgSource(
            prefix="",
            remap={
                "transport": "transport",
                "host": "host",
                "port": "port",
            },
        ),
    ]

    transport: Literal["stdio", "sse", "streamable-http"] = Field(default="stdio")
    host: str = Field(default="127.0.0.1")
    port: int = Field(default=8000)

    @field_validator("port")
    @classmethod
    def validate_port(cls, v: int) -> int:
        """验证端口范围 | Validate port range"""
        if not 1 <= v <= 65535:
            raise ValueError("Port must be between 1 and 65535")
        return v
