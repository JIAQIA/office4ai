# -*- coding: utf-8 -*-
# filename: base.py
# @Time    : 2025/12/18 16:07
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm

from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Any


class BaseTool(ABC):
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
    def input_schema(self) -> dict[str, Any]:  # pragma: no cover
        raise NotImplementedError

    @abstractmethod
    async def execute(self, arguments: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        raise NotImplementedError
