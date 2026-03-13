"""WordWindowResource 契约测试 — 真实 Socket.IO + MockAddInClient."""

from __future__ import annotations

import time

import pytest

from office4ai.a2c_smcp.resources.word_window import WordWindowResource
from tests.contract_tests.mock_addin.client import MockAddInClient


@pytest.mark.asyncio
@pytest.mark.contract
class TestWordWindowContract:
    async def test_read_fetches_stats_and_content(
        self,
        word_window_resource: WordWindowResource,
        word_factory,
    ) -> None:
        """完整链路：MockAddInClient 响应 stats + visibleContent → read() 渲染正确。"""
        doc_uri = "file:///tmp/contract_word_test.docx"

        client = MockAddInClient(
            server_url="http://127.0.0.1:3003",
            namespace="/word",
            client_id="contract_word_client_stats",
            document_uri=doc_uri,
        )

        client.register_response(
            "word:get:documentStats",
            lambda req: {
                "requestId": req["requestId"],
                "success": True,
                "data": {"pageCount": 5, "wordCount": 1200, "paragraphCount": 20},
                "timestamp": time.time(),
                "duration": 10,
            },
        )
        client.register_response(
            "word:get:visibleContent",
            lambda req: {
                "requestId": req["requestId"],
                "success": True,
                "data": {"text": "Hello World\nContract test paragraph"},
                "timestamp": time.time(),
                "duration": 10,
            },
        )

        await client.connect()
        try:
            # 设置激活文档
            word_window_resource.workspace.update_last_activity(doc_uri, "word_get_visible_content", {})

            content = await word_window_resource.read()

            assert "文档列表 (1)" in content
            assert "⭐" in content
            assert "总页数: 5" in content
            assert "1,200" in content
            assert "Hello World" in content
        finally:
            await client.disconnect()

    async def test_read_timeout_degradation(
        self,
        word_window_resource: WordWindowResource,
    ) -> None:
        """超时降级：5s 延迟响应 → 3s 超时 → 降级渲染。"""
        import asyncio

        doc_uri = "file:///tmp/contract_word_timeout.docx"

        client = MockAddInClient(
            server_url="http://127.0.0.1:3003",
            namespace="/word",
            client_id="contract_word_client_timeout",
            document_uri=doc_uri,
        )

        async def slow_response(req):
            await asyncio.sleep(5)
            return {
                "requestId": req["requestId"],
                "success": True,
                "data": {"pageCount": 1, "wordCount": 10, "paragraphCount": 1},
                "timestamp": time.time(),
                "duration": 5000,
            }

        client.register_response("word:get:documentStats", slow_response)
        client.register_response("word:get:visibleContent", slow_response)

        await client.connect()
        try:
            word_window_resource.workspace.update_last_activity(doc_uri, "word_get_visible_content", {})

            content = await word_window_resource.read()

            assert "文档列表 (1)" in content
            assert "不可用" in content or "超时" in content
        finally:
            await client.disconnect()

    async def test_read_no_active_document(
        self,
        word_window_resource: WordWindowResource,
    ) -> None:
        """无 last_activity → 列表正常, 无详情, 无 Socket.IO 请求。"""
        doc_uri = "file:///tmp/contract_word_noactive.docx"

        client = MockAddInClient(
            server_url="http://127.0.0.1:3003",
            namespace="/word",
            client_id="contract_word_client_noactive",
            document_uri=doc_uri,
        )

        await client.connect()
        try:
            content = await word_window_resource.read()

            assert "文档列表 (1)" in content
            assert "激活文档" not in content
            assert len(client.received_events) == 0
        finally:
            await client.disconnect()
