# filename: utils.py
# @Time    : 2025/10/29 10:35
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm
import ast
import json
import os
from typing import Any, Literal

from loguru import logger


def _extract_module_info(file_path: str) -> dict[str, str | list[Any]]:
    """
    提取Python模块的__all__定义和docstring | Extract __all__ definition and docstring from Python module

    Args:
        file_path (str): Python文件路径 | Python file path

    Returns:
        dict: 包含'docstring'和'__all__'的字典 | Dict containing 'docstring' and '__all__'
    """
    result: dict[str, str | list[Any]] = {"docstring": "", "__all__": []}

    try:
        with open(file_path, encoding="utf-8") as f:
            content = f.read()

        tree = ast.parse(content)

        # 提取模块级docstring | Extract module-level docstring
        if isinstance(tree, ast.Module) and tree.body:
            first_node = tree.body[0]
            if isinstance(first_node, ast.Expr) and isinstance(first_node.value, ast.Constant):
                if isinstance(first_node.value.value, str):
                    result["docstring"] = first_node.value.value.strip()

        # 提取__all__定义 | Extract __all__ definition
        for node in ast.walk(tree):
            if isinstance(node, ast.Assign):
                for target in node.targets:
                    if isinstance(target, ast.Name) and target.id == "__all__":
                        if isinstance(node.value, (ast.List, ast.Tuple)):
                            all_items = []
                            for elt in node.value.elts:
                                if isinstance(elt, ast.Constant) and isinstance(elt.value, str):
                                    all_items.append(elt.value)
                            result["__all__"] = all_items
                        break

    except SyntaxError as e:
        logger.error(f"Syntax error in {file_path}: {e}")

    except FileNotFoundError:
        logger.error(f"File not found: {file_path}")

    except Exception as e:
        logger.exception(f"Unexpected error processing {file_path}: {e}")

    return result


def _collect_package_info(entry_path: str, entry_name: str, descriptions: dict) -> None:
    """
    收集Python包的描述信息 | Collect Python package description info

    Args:
        entry_path (str): 目录路径 | Directory path
        entry_name (str): 目录名称 | Directory name
        descriptions (dict): 描述信息字典 | Descriptions dictionary
    """
    init_file = os.path.join(entry_path, "__init__.py")
    if os.path.isfile(init_file):
        info = _extract_module_info(init_file)
        if info["docstring"] or info["__all__"]:
            descriptions[f"{entry_name}/"] = info


def _format_descriptions(descriptions: dict) -> str:
    """
    格式化描述信息为Markdown | Format descriptions to Markdown

    Args:
        descriptions (dict): 描述信息字典 | Descriptions dictionary

    Returns:
        str: 格式化的描述字符串 | Formatted description string
    """
    if not descriptions:
        return ""

    desc_lines = ["\n---\n"]
    for name, info in sorted(descriptions.items()):
        desc_lines.append(f"**{name}**")

        # 在尾部显示__all__和docstring | Show __all__ and docstring in footer
        if info["__all__"]:
            desc_lines.append(f"  __all__: {json.dumps(info['__all__'], ensure_ascii=False)}")

        if info["docstring"]:
            # 缩进docstring | Indent docstring
            docstring_lines = info["docstring"].split("\n")
            for line in docstring_lines:
                desc_lines.append(f"  {line}" if line.strip() else "")

        desc_lines.append("")  # 空行分隔 | Empty line separator

    return "\n".join(desc_lines)


def list_directory_tree_with_desc(
    path: str,
    include_dirs: list[str] | set[str] | Literal["all"] | None = None,
    recursive: bool = True,
    indent: str = "",
    _is_root: bool = True,
    _descriptions: dict | None = None,
    _is_last: bool = True,
) -> str:
    """
    递归或非递归列出目录树结构，并附带Python包/模块的描述信息 | List directory tree with Python package/module descriptions

    Args:
        path (str): 要遍历的根目录路径 | Root directory path to traverse
        include_dirs (list[str] | set[str] | 'all' | None): 允许展开的目录列表或'all'表示展开所有目录 | Directories to expand
            or 'all' for all
        recursive (bool): 是否递归展开目录 | Whether to expand directories recursively
        indent (str): 当前递归层的缩进，用于格式化输出 | Current indentation for formatting
        _is_root (bool): 内部参数，标识是否为根调用 | Internal param, indicates if root call
        _descriptions (dict | None): 内部参数，收集描述信息 | Internal param, collects descriptions
        _is_last (bool): 内部参数，标识是否为最后一项 | Internal param, indicates if last item

    Returns:
        str: 格式化的目录树字符串，包含描述信息 | Formatted directory tree with descriptions
    """
    output = []
    if include_dirs is None:
        include_dirs = []

    # 初始化描述字典 | Initialize descriptions dict
    if _descriptions is None:
        _descriptions = {}

    with os.scandir(path) as entries:
        entries_list = sorted(entries, key=lambda e: (not e.is_dir(), e.name))
        total = len(entries_list)

        for idx, entry in enumerate(entries_list):
            is_last = idx == total - 1
            entry_path = os.path.join(path, entry.name)

            # 确定前缀符号 | Determine prefix symbol
            if _is_root:
                prefix = ""
                connector = ""
            else:
                prefix = "└── " if is_last else "├── "
                connector = "    " if is_last else "│   "

            if entry.is_dir(follow_symlinks=False):
                # 检查是否为Python包（包含__init__.py）| Check if it's a Python package
                _collect_package_info(entry_path, entry.name, _descriptions)

                # 构建目录行，如果有__all__则添加为行内注释 | Build dir line with inline __all__ comment
                dir_key = f"{entry.name}/"
                dir_line = f"{indent}{prefix}{dir_key}"
                if dir_key in _descriptions and _descriptions[dir_key]["__all__"]:
                    all_str = json.dumps(_descriptions[dir_key]["__all__"], ensure_ascii=False)
                    dir_line += f"  # __all__: {all_str}"
                output.append(dir_line)

                if include_dirs == "all" or entry.name in include_dirs:
                    new_indent = indent + connector
                    output.append(
                        list_directory_tree_with_desc(
                            entry_path,
                            "all" if recursive else include_dirs,
                            recursive,
                            new_indent,
                            _is_root=False,
                            _descriptions=_descriptions,
                            _is_last=is_last,
                        ),
                    )

            elif entry.is_file(follow_symlinks=False):
                # 如果是Python文件，提取信息 | If it's a Python file, extract info
                if entry.name.endswith(".py"):
                    info = _extract_module_info(entry_path)
                    if info["docstring"] or info["__all__"]:
                        _descriptions[entry.name] = info

                # 构建文件行，如果有__all__则添加为行内注释 | Build file line with inline __all__ comment
                file_line = f"{indent}{prefix}{entry.name}"
                if entry.name in _descriptions and _descriptions[entry.name]["__all__"]:
                    all_str = json.dumps(_descriptions[entry.name]["__all__"], ensure_ascii=False)
                    file_line += f"  # __all__: {all_str}"
                output.append(file_line)

    tree_output = "\n".join(output)

    # 只在根调用时添加描述部分 | Add descriptions section only at root call
    if _is_root:
        tree_output += _format_descriptions(_descriptions)

    return tree_output


def get_minimal_expanded_tree_with_desc(
    root_dir: str,
    target_file_path: str,
    indent: str = "",
    _is_root: bool = True,
    _descriptions: dict | None = None,
) -> str:
    """
    生成最小化展开的目录树，仅展开到目标文件所在路径，并附带Python包/模块的描述信息
    Generate minimally expanded directory tree to target file with Python package/module descriptions

    这个函数从项目根目录开始，仅展开包含目标文件的路径分支，其他分支只显示一级
    This function expands only the path branch containing the target file from root, showing only first level for others

    Args:
        root_dir (str): 项目根目录 | Project root directory
        target_file_path (str): 目标文件的绝对路径 | Absolute path of target file
        indent (str): 当前缩进 | Current indentation
        _is_root (bool): 内部参数，标识是否为根调用 | Internal param, indicates if root call
        _descriptions (dict | None): 内部参数，收集描述信息 | Internal param, collects descriptions

    Returns:
        str: 格式化的目录树字符串，包含描述信息 | Formatted directory tree with descriptions
    """
    output = []

    # 初始化描述字典 | Initialize descriptions dict
    if _descriptions is None:
        _descriptions = {}

    # 确保路径是绝对路径 | Ensure paths are absolute
    root_dir = os.path.realpath(root_dir)
    target_file_path = os.path.realpath(target_file_path)

    # 获取从根目录到目标文件的路径组件 | Get path components from root to target file
    try:
        rel_path = os.path.relpath(target_file_path, root_dir)
        path_parts = rel_path.split(os.sep)
    except ValueError:
        # 目标文件不在根目录下 | Target file not under root directory
        return list_directory_tree_with_desc(root_dir, include_dirs=None, recursive=False, indent=indent)

    # 遍历当前目录 | Traverse current directory
    try:
        with os.scandir(root_dir) as entries:
            entries_list = sorted(entries, key=lambda e: (not e.is_dir(), e.name))
            total = len(entries_list)

            for idx, entry in enumerate(entries_list):
                is_last = idx == total - 1
                entry_path = os.path.join(root_dir, entry.name)

                # 确定前缀符号 | Determine prefix symbol
                if _is_root:
                    prefix = ""
                    connector = ""
                else:
                    prefix = "└── " if is_last else "├── "
                    connector = "    " if is_last else "│   "

                if entry.is_dir(follow_symlinks=False):
                    # 检查是否为Python包 | Check if it's a Python package
                    _collect_package_info(entry_path, entry.name, _descriptions)

                    # 构建目录行，如果有__all__则添加为行内注释 | Build dir line with inline __all__ comment
                    dir_key = f"{entry.name}/"
                    dir_line = f"{indent}{prefix}{dir_key}"
                    if dir_key in _descriptions and _descriptions[dir_key]["__all__"]:
                        all_str = json.dumps(_descriptions[dir_key]["__all__"], ensure_ascii=False)
                        dir_line += f"  # __all__: {all_str}"
                    output.append(dir_line)

                    # 检查这个目录是否在目标文件的路径上 | Check if this dir is on target file path
                    if path_parts and entry.name == path_parts[0]:
                        # 这个目录在目标路径上，递归展开 | This dir is on target path, expand recursively
                        new_indent = indent + connector
                        next_target = os.path.join(root_dir, *path_parts)
                        output.append(
                            get_minimal_expanded_tree_with_desc(
                                entry_path,
                                next_target,
                                new_indent,
                                _is_root=False,
                                _descriptions=_descriptions,
                            ),
                        )
                    else:
                        # 不在目标路径上，只显示一级 | Not on target path, show only first level
                        new_indent = indent + connector
                        output.append(
                            list_directory_tree_with_desc(
                                entry_path,
                                include_dirs=None,
                                recursive=False,
                                indent=new_indent,
                                _is_root=False,
                                _descriptions=_descriptions,
                            ),
                        )

                elif entry.is_file(follow_symlinks=False):
                    # 如果是Python文件，提取信息 | If it's a Python file, extract info
                    if entry.name.endswith(".py"):
                        info = _extract_module_info(entry_path)
                        if info["docstring"] or info["__all__"]:
                            _descriptions[entry.name] = info

                    # 构建文件行，如果有__all__则添加为行内注释 | Build file line with inline __all__ comment
                    file_line = f"{indent}{prefix}{entry.name}"

                    # 如果是目标文件，标记它 | Mark target file
                    if entry_path == target_file_path:
                        file_line += " ← 当前文件 | Current file"
                    elif entry.name in _descriptions and _descriptions[entry.name]["__all__"]:
                        all_str = json.dumps(_descriptions[entry.name]["__all__"], ensure_ascii=False)
                        file_line += f"  # __all__: {all_str}"

                    output.append(file_line)

    except PermissionError:
        output.append(f"{indent}[权限不足 | Permission denied]")

    tree_output = "\n".join(output)

    # 只在根调用时添加描述部分 | Add descriptions section only at root call
    if _is_root:
        tree_output += _format_descriptions(_descriptions)

    return tree_output
