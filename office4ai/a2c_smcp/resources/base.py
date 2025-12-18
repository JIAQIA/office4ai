# -*- coding: utf-8 -*-
# filename: base.py
# @Time    : 2025/12/18 16:07
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm

from __future__ import annotations

from abc import ABC, abstractmethod


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
