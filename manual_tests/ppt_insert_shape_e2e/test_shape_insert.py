"""
PPT Insert Shape E2E Tests — Basic

测试基本的 PPT 形状插入功能。

测试场景:
1. 矩形 — 插入 Rectangle
2. 圆形 — 插入 Circle
3. 带文本 — 插入带 text 选项的形状
4. 带样式 — 插入带 fillColor/borderColor 的形状

运行方式:
    uv run python manual_tests/ppt_insert_shape_e2e/test_shape_insert.py --test all
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
from manual_tests.ppt_test_helpers import ppt_insert_shape

FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "ppt_e2e"


def validate_shape_insert(data: dict[str, Any]) -> bool:
    """通用形状插入验证"""
    print(f"   ✅ 协议返回: {data}")
    return True


TEST_CASES: list[PptTestCase] = [
    PptTestCase(
        name="插入矩形",
        fixture_name="empty.pptx",
        description="插入 Rectangle 形状",
        validator=validate_shape_insert,
        tags=["basic"],
    ),
    PptTestCase(
        name="插入圆形",
        fixture_name="empty.pptx",
        description="插入 Circle 形状",
        validator=validate_shape_insert,
        tags=["basic"],
    ),
    PptTestCase(
        name="带文本形状",
        fixture_name="empty.pptx",
        description="插入带 text='Hello Shape' 的矩形",
        validator=validate_shape_insert,
        tags=["basic"],
    ),
    PptTestCase(
        name="带样式形状",
        fixture_name="empty.pptx",
        description="插入带红色填充和蓝色边框的矩形",
        validator=validate_shape_insert,
        tags=["advanced"],
    ),
]

_SHAPE_PARAMS: list[tuple[str, dict[str, Any] | None]] = [
    ("Rectangle", {"left": 100, "top": 100, "width": 200, "height": 150}),
    ("Circle", {"left": 300, "top": 100, "width": 150, "height": 150}),
    ("Rectangle", {"text": "Hello Shape", "left": 100, "top": 300, "width": 250, "height": 100}),
    (
        "Rectangle",
        {
            "fillColor": "#FF0000",
            "borderColor": "#0000FF",
            "borderWidth": 3,
            "left": 100,
            "top": 100,
            "width": 200,
            "height": 150,
        },
    ),
]


async def run_single_test(runner: PPTTestRunner, test_case: PptTestCase, test_number: int) -> bool:
    print("\n" + "=" * 70)
    print(f"🧪 测试 {test_number}: {test_case.name}")
    print("=" * 70)
    print(f"📋 描述: {test_case.description}")

    fixture_path = f"ppt_e2e/{test_case.fixture_name}"
    shape_type, options = _SHAPE_PARAMS[test_number - 1]

    try:
        async with runner.run_with_workspace(fixture_path, open_delay=3.0) as (workspace, fixture):
            print(f"\n📝 执行: 插入形状 '{shape_type}' (options={options})...")
            start_time = time.time()
            success, data, error = await ppt_insert_shape(workspace, fixture.document_uri, shape_type, options=options)
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

    parser = argparse.ArgumentParser(description="PPT Insert Shape E2E Tests — Basic")
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
