"""Contract test fixtures for window resources."""

from __future__ import annotations

import pytest_asyncio

from office4ai.a2c_smcp.resources.ppt_window import PptWindowResource
from office4ai.a2c_smcp.resources.word_window import WordWindowResource
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


@pytest_asyncio.fixture
async def word_window_resource(workspace: OfficeWorkspace) -> WordWindowResource:
    return WordWindowResource(workspace, priority=50, fullscreen=True)


@pytest_asyncio.fixture
async def ppt_window_resource(workspace: OfficeWorkspace) -> PptWindowResource:
    return PptWindowResource(workspace, priority=50, fullscreen=True)
