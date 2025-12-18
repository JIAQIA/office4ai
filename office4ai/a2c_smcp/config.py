# -*- coding: utf-8 -*-
# filename: config.py
# @Time    : 2025/12/18 16:07
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm

from __future__ import annotations

from typing import Literal

from confz import BaseConfig, CLArgSource, EnvSource
from confz.base_config import BaseConfigMetaclass
from pydantic import Field


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
