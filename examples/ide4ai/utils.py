# filename: utils.py
# @Time    : 2024/5/8 17:24
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
import os
from typing import Literal

from loguru import logger

from ide4ai.dtos.text_documents import LSPRange
from ide4ai.environment.workspace.schema import Range


def render_symbols(symbols: list[dict], render_symbol_kind: list[int], indent: int = 0) -> str:
    """
    递归渲染LSP符号列表为人类可读的文本格式，并返回形成的字符串。

    Args:
        symbols: 符号列表，每个符号是一个包含name, kind, 可选children的字典。
        render_symbol_kind: 需要渲染的符号种类列表。
        indent: 当前缩进级别，用于格式化输出。

    返回:
        str: 返回形成的符号结构字符串。
    """
    # 用于缩进的空格
    indent_space = " " * 2 * indent
    # 符号种类的字典映射
    symbol_kinds = {
        1: "File",
        2: "Module",
        3: "Namespace",
        4: "Package",
        5: "Class",
        6: "Method",
        7: "Property",
        8: "Field",
        9: "Constructor",
        10: "Enum",
        11: "Interface",
        12: "Function",
        13: "Variable",
        14: "Constant",
        15: "String",
        16: "Number",
        17: "Boolean",
        18: "Array",
        19: "Object",
        20: "Key",
        21: "Null",
        22: "EnumMember",
        23: "Struct",
        24: "Event",
        25: "Operator",
        26: "TypeParameter",
    }

    lines = []  # 用于收集所有的输出行
    for symbol in symbols:
        if symbol["kind"] not in render_symbol_kind:
            continue
        # 获取符号的种类名称，如果找不到则默认为'Unknown Symbol'
        kind_name = symbol_kinds.get(symbol["kind"], "Unknown Symbol")

        # 构造当前符号的描述
        line = f"{indent_space}{kind_name}: {symbol['name']}"
        if lsp_range_dict := symbol.get("location", {}).get("range"):
            lsp_range = LSPRange.model_validate(lsp_range_dict)
            tf_ide_range = Range.from_lsp_range(lsp_range)
            line += (
                f" Range({tf_ide_range.start_position.line}:{tf_ide_range.start_position.character}-"
                f"{tf_ide_range.end_position.line}:{tf_ide_range.end_position.character})"
            )
        lines.append(line)

        # 如果有子符号，递归调用以渲染它们，并将结果添加到lines中
        if "children" in symbol:
            child_output = render_symbols(symbol["children"], render_symbol_kind, indent + 1)
            lines.append(child_output)

    return "\n".join(lines)  # 将所有行合并为一个单独的字符串并返回


def list_directory_tree(
    path: str,
    include_dirs: list[str] | set[str] | Literal["all"] | None = None,
    recursive: bool = True,
    indent: str = "",
) -> str:
    """
    递归或非递归列出目录树结构，返回一个目录树的字符串。

    Args:
        path (str): 要遍历的根目录路径。
        include_dirs (Iterable[str] | 'all'): 允许展开的目录列表或 'all' 表示展开所有目录。
        recursive (bool): 是否递归展开目录。
        indent (str): 当前递归层的缩进，用于格式化输出。

    Returns:
        str: 格式化的目录树字符串。
    """
    output = []
    if include_dirs is None:
        include_dirs = []
    with os.scandir(path) as entries:
        for entry in entries:
            entry_path = os.path.join(path, entry.name)
            if entry.is_dir(follow_symlinks=False):
                output.append(f"{indent}{entry.name}/")
                if include_dirs == "all" or entry.name in include_dirs:
                    new_indent = "  " + indent  # 增加缩进
                    output.append(
                        list_directory_tree(
                            entry_path,
                            "all" if recursive else include_dirs,
                            recursive,
                            new_indent,
                        ),
                    )
            elif entry.is_file(follow_symlinks=False):
                output.append(f"{indent}{entry.name}")

    return "\n".join(output)


def is_subdirectory(sub_dir: str, root_dir: str) -> bool:
    """
    Determines if a directory is a subdirectory of another directory.

    Args:
        sub_dir: A string representing the subdirectory path.
        root_dir: A string representing the root directory path.

    Returns:
        A boolean value indicating whether the subdirectory is a subdirectory of the root directory.

    Note:
        This method ensures that both paths are absolute paths by using os.path.realpath.
        It checks if both directories exist using os.path.exists.
        It compares the common path between the subdirectory and root directory using os.path.commonpath.
        If the common path is equal to the root directory, it returns True, indicating that the subdirectory is a subdirectory of the
        root directory. Otherwise, it returns False.

    Example:
        is_subdirectory('/path/to/subdir', '/path/to')
        # returns True
    """
    # 确保两个路径都是绝对路径
    sub_dir = os.path.realpath(sub_dir)
    root_dir = os.path.realpath(root_dir)

    # 检查两个目录是否存在
    if not os.path.exists(sub_dir) or not os.path.exists(root_dir):
        return False

    # 检查sub_dir是否是root_dir的子目录
    # 这里使用os.path.commonpath来比较两个路径的公共部分是否为root_dir
    common_path = os.path.commonpath([sub_dir, root_dir])
    if common_path == root_dir:
        return True
    else:
        return False


def get_minimal_expanded_tree(root_dir: str, target_file_path: str, indent: str = "") -> str:
    """
    生成最小化展开的目录树，仅展开到目标文件所在路径 | Generate minimally expanded directory tree to target file

    这个函数从项目根目录开始，仅展开包含目标文件的路径分支，其他分支只显示一级
    This function expands only the path branch containing the target file from root, showing only first level for others

    Args:
        root_dir (str): 项目根目录 | Project root directory
        target_file_path (str): 目标文件的绝对路径 | Absolute path of target file
        indent (str): 当前缩进 | Current indentation

    Returns:
        str: 格式化的目录树字符串 | Formatted directory tree string
    """
    output = []

    # 确保路径是绝对路径 | Ensure paths are absolute
    root_dir = os.path.realpath(root_dir)
    target_file_path = os.path.realpath(target_file_path)

    # 获取从根目录到目标文件的路径组件 | Get path components from root to target file
    try:
        rel_path = os.path.relpath(target_file_path, root_dir)
        path_parts = rel_path.split(os.sep)
    except ValueError:
        # 目标文件不在根目录下 | Target file not under root directory
        return list_directory_tree(root_dir, include_dirs=None, recursive=False, indent=indent)

    # 遍历当前目录 | Traverse current directory
    try:
        with os.scandir(root_dir) as entries:
            for entry in entries:
                entry_path = os.path.join(root_dir, entry.name)

                if entry.is_dir(follow_symlinks=False):
                    output.append(f"{indent}{entry.name}/")

                    # 检查这个目录是否在目标文件的路径上 | Check if this dir is on target file path
                    if path_parts and entry.name == path_parts[0]:
                        # 这个目录在目标路径上，递归展开 | This dir is on target path, expand recursively
                        new_indent = "  " + indent
                        next_target = os.path.join(root_dir, *path_parts)
                        output.append(
                            get_minimal_expanded_tree(entry_path, next_target, new_indent),
                        )
                    else:
                        # 不在目标路径上，只显示一级 | Not on target path, show only first level
                        new_indent = "  " + indent
                        output.append(
                            list_directory_tree(entry_path, include_dirs=None, recursive=False, indent=new_indent),
                        )

                elif entry.is_file(follow_symlinks=False):
                    # 如果是目标文件，标记它 | Mark target file
                    if entry_path == target_file_path:
                        output.append(f"{indent}{entry.name} ← 当前文件 | Current file")
                    else:
                        output.append(f"{indent}{entry.name}")
    except PermissionError:
        output.append(f"{indent}[权限不足 | Permission denied]")

    return "\n".join(output)


def detect_makefile_commands(root_dir: str) -> dict[str, list[str]] | None:
    """
    检测项目根目录下的Makefile并提取可用命令 | Detect Makefile and extract available commands

    检测规则 | Detection rules:
    1. 标准Makefile名称：Makefile, makefile, GNUmakefile (GNU Make标准约定)
    2. .mk扩展名文件：*.mk (模块化makefile片段)
    3. Makefile.*变体：Makefile.* (特定平台或配置的makefile)

    Args:
        root_dir (str): 项目根目录 | Project root directory

    Returns:
        dict[str, list[str]] | None: 命令字典，key为命令前缀(如"make")，value为命令列表 |
            Command dict, key is command prefix (e.g., "make"), value is command list
            如果没有Makefile则返回None | Returns None if no Makefile found
    """
    import re
    from pathlib import Path

    root_path = Path(root_dir)
    all_targets = set()  # 使用set去重 | Use set to deduplicate

    # 1. 检测标准Makefile名称 | Detect standard Makefile names
    # GNU Make按此优先级查找：GNUmakefile > makefile > Makefile
    # GNU Make searches in this priority: GNUmakefile > makefile > Makefile
    standard_names = ["GNUmakefile", "makefile", "Makefile"]

    # 2. 使用glob查找所有可能的makefile | Use glob to find all possible makefiles
    makefile_patterns = [
        "*.mk",  # 模块化makefile片段 | Modular makefile fragments
        "Makefile.*",  # 平台特定makefile | Platform-specific makefiles
    ]

    makefile_paths = []

    # 添加标准名称的makefile | Add standard named makefiles
    for name in standard_names:
        path = root_path / name
        if path.is_file():
            makefile_paths.append(path)
            break  # 找到第一个标准makefile就停止 | Stop at first standard makefile

    # 添加通过glob找到的makefile | Add makefiles found via glob
    for pattern in makefile_patterns:
        makefile_paths.extend(root_path.glob(pattern))

    if not makefile_paths:
        return None

    # 提取所有目标（targets）| Extract all targets
    # Makefile目标格式: target: dependencies | Makefile target format: target: dependencies
    # 匹配行首的目标定义，排除以.开头的特殊目标 |
    # Match target definitions at line start, exclude special targets starting with .
    pattern = r"^([a-zA-Z0-9_-]+):"

    for makefile_path in makefile_paths:
        try:
            with open(makefile_path, encoding="utf-8") as f:
                content = f.read()

            for line in content.split("\n"):
                line = line.strip()
                # 跳过注释和空行 | Skip comments and empty lines
                if line.startswith("#") or not line:
                    continue
                match = re.match(pattern, line)
                if match:
                    target = match.group(1)
                    # 排除常见的内部目标 | Exclude common internal targets
                    if not target.startswith(".") and target not in ["PHONY"]:
                        all_targets.add(target)

        except Exception as e:
            # 读取失败，忽略该文件 | Read failed, ignore this file
            logger.error(f"处理Makefile {makefile_path} 时发生异常: {e}")

    if all_targets:
        # 返回排序后的目标列表 | Return sorted target list
        return {"make": sorted(all_targets)}

    return None
