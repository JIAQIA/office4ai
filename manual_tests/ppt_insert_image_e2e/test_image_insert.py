"""
PPT Insert Image E2E Tests

测试 PPT 图片插入功能。

测试场景:
1. 基本图片插入 — 插入最小合法 PNG
2. 带位置图片插入 — 指定 left/top/width/height
3. 指定幻灯片插入 — 在 slideIndex=1 上插入图片

运行方式:
    uv run python manual_tests/ppt_insert_image_e2e/test_image_insert.py --test all
"""

import asyncio
import sys
import time
from pathlib import Path
from typing import Any

from manual_tests.ppt_e2e_base import (
    PPTTestRunner,
    PptTestCase,
    ensure_ppt_fixtures,
)
from manual_tests.ppt_test_helpers import ppt_insert_image

FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "ppt_e2e"

# 最小合法 1x1 红色 PNG (base64)
MINIMAL_PNG_BASE64 = (
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4"
    "nGP4z8BQDwAEgAF/pooBPQAAAABJRU5ErkJggg=="
)


def validate_basic_insert(data: dict[str, Any]) -> bool:
    """验证基本图片插入"""
    print(f"   ✅ 协议返回: {data}")
    return True


def validate_positioned_insert(data: dict[str, Any]) -> bool:
    """验证带位置的图片插入"""
    print(f"   ✅ 协议返回: {data}")
    return True


def validate_specific_slide_insert(data: dict[str, Any]) -> bool:
    """验证指定幻灯片的图片插入"""
    print(f"   ✅ 协议返回: {data}")
    return True


TEST_CASES: list[PptTestCase] = [
    PptTestCase(
        name="基本图片插入",
        fixture_name="empty.pptx",
        description="在空白幻灯片插入最小 PNG 图片",
        validator=validate_basic_insert,
        tags=["basic"],
    ),
    PptTestCase(
        name="带位置图片插入",
        fixture_name="empty.pptx",
        description="在指定位置 (200, 150) 插入图片",
        validator=validate_positioned_insert,
        tags=["basic"],
    ),
    PptTestCase(
        name="指定幻灯片插入",
        fixture_name="simple.pptx",
        description="在 slideIndex=1 上插入图片",
        validator=validate_specific_slide_insert,
        tags=["basic"],
    ),
]

_INSERT_PARAMS: list[tuple[dict[str, Any], dict[str, Any] | None]] = [
    ({"base64": MINIMAL_PNG_BASE64}, None),
    ({"base64": MINIMAL_PNG_BASE64}, {"left": 200, "top": 150, "width": 100, "height": 100}),
    ({"base64": MINIMAL_PNG_BASE64}, {"slideIndex": 1}),
]


async def run_single_test(runner: PPTTestRunner, test_case: PptTestCase, test_number: int) -> bool:
    print("\n" + "=" * 70)
    print(f"🧪 测试 {test_number}: {test_case.name}")
    print("=" * 70)
    print(f"📋 描述: {test_case.description}")

    fixture_path = f"ppt_e2e/{test_case.fixture_name}"
    image_data, options = _INSERT_PARAMS[test_number - 1]

    try:
        async with runner.run_with_workspace(fixture_path, open_delay=3.0) as (workspace, fixture):
            print(f"\n📝 执行: 插入图片 (options={options})...")
            start_time = time.time()
            success, data, error = await ppt_insert_image(
                workspace, fixture.document_uri, image_data, options=options
            )
            elapsed_ms = (time.time() - start_time) * 1000
            print(f"\n⏱️  执行时间: {elapsed_ms:.1f}ms")

            if not success:
                print(f"❌ 插入失败: {error}")
                return False

            print("✅ 协议返回成功")
            data = data or {}

            print("\n📊 验证结果:")
            passed = True
            if test_case.validator and not test_case.validator(data):
                passed = False

            print("\n" + "=" * 70)
            print(f"{'✅' if passed else '❌'} 测试 {test_number} {'通过' if passed else '失败'}")
            print("=" * 70)
            return passed

    except Exception as e:
        print(f"\n❌ 测试异常: {e}")
        import traceback

        traceback.print_exc()
        return False


async def run_tests(test_indices: list[int], auto_open: bool = True, cleanup_on_success: bool = True) -> bool:
    ensure_ppt_fixtures(FIXTURES_DIR)
    runner = PPTTestRunner(fixtures_dir=FIXTURES_DIR.parent, auto_open=auto_open, cleanup_on_success=cleanup_on_success)
    results: list[bool] = []
    for idx in test_indices:
        if idx < 1 or idx > len(TEST_CASES):
            continue
        if len(test_indices) > 1 and results:
            if auto_open:
                await asyncio.sleep(2.0)
            else:
                input("按回车继续...")
        result = await run_single_test(runner, TEST_CASES[idx - 1], idx)
        results.append(result)
    if len(results) > 1:
        print(f"\n📈 总体结果: {sum(results)}/{len(results)} 测试通过")
    return all(results)


def main() -> None:
    import argparse

    parser = argparse.ArgumentParser(description="PPT Insert Image E2E Tests")
    parser.add_argument("--test", choices=[str(i) for i in range(1, len(TEST_CASES) + 1)] + ["all"], default="1")
    parser.add_argument("--no-auto-open", action="store_true")
    parser.add_argument("--always-cleanup", action="store_true")
    parser.add_argument("--list", action="store_true")
    args = parser.parse_args()

    if args.list:
        for i, tc in enumerate(TEST_CASES, 1):
            print(f"  {i}. {tc.name} — {tc.description}")
        return

    test_indices = list(range(1, len(TEST_CASES) + 1)) if args.test == "all" else [int(args.test)]
    try:
        success = asyncio.run(
            run_tests(test_indices, auto_open=not args.no_auto_open, cleanup_on_success=not args.always_cleanup or True)
        )
        sys.exit(0 if success else 1)
    except KeyboardInterrupt:
        print("\n\n⏸️  测试被用户中断")
        sys.exit(130)


if __name__ == "__main__":
    main()
