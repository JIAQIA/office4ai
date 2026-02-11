"""
Test Socket.IO server lifecycle

测试 Socket.IO 服务器的生命周期。
"""

import pytest
from aiohttp import web

from office4ai.environment.workspace.socketio.server import create_app


@pytest.mark.asyncio
@pytest.mark.integration
async def test_create_app() -> None:
    """Test creating aiohttp app with Socket.IO"""
    app = create_app()

    assert isinstance(app, web.Application)
    # Health check is added by create_app, will be tested in test_health_check_endpoint


@pytest.mark.asyncio
@pytest.mark.integration
async def test_health_check_endpoint() -> None:
    """Test health check endpoint returns correct data"""
    from aiohttp.test_utils import TestClient, TestServer

    app = create_app()

    async with TestServer(app) as server:
        async with TestClient(server) as client:
            response = await client.get("/health")

            assert response.status == 200
            data = await response.json()

            assert data["status"] == "ok"
            assert "service" in data
            assert data["service"] == "office4ai-workspace-socketio"
            assert "connections" in data
            assert "documents" in data
