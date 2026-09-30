# -*- coding: utf-8 -*-
"""
origin-bridge — 用 Python 通过 COM 自动化控制 OriginLab Origin
================================================================

一行代码连上 Origin，读写数据、算统计、画图、存项目::

    from origin_bridge import OriginBridge

    with OriginBridge() as ob:
        ob.new_workbook("Demo")
        ob.set_column("A", [1, 2, 3, 4, 5])
        ob.set_column("B", [2, 4, 6, 8, 10])
        print(ob.get_column("A"))     # [1.0, 2.0, 3.0, 4.0, 5.0]
        print(ob.stats("B"))          # 列 B: count=5 total=30 ...
        ob.plot("A", "B", kind="line_symbol")
        ob.save(r"D:\\data\\demo.opju")

还附带界面语言切换工具（Origin 没有图形化语言开关）::

    from origin_bridge import switch_language

    switch_language("中文")          # 切到中文界面
    switch_language("en", restart=True)   # 切到英文并重启

命令行用法::

    origin-cli info
    origin-cli lang zh
    origin-cli import data.csv --book MyData
    origin-cli stats A
    origin-cli plot A B --type scatter
    origin-cli save result.opju
"""

from .bridge import (
    PROG_ID,
    PLOT_TYPES,
    ColumnStats,
    OriginBridge,
    OriginCommandError,
    OriginConnectionError,
    OriginError,
    OriginInfo,
    OriginNotFoundError,
    find_origin_exe,
    find_origin_install,
    find_origin_versions,
    index_to_letters,
    is_origin_running,
    iter_origin_reg_keys,
)

from .language import (
    LANGUAGE_NAMES,
    LanguageError,
    LanguageState,
    close_origin,
    describe as describe_language,
    get_language,
    launch_origin,
    list_installations,
    list_processes,
    normalize_language,
    set_language,
    switch_language,
)

__version__ = "0.1.0"

__all__ = [
    # 核心
    "OriginBridge",
    "OriginInfo",
    "ColumnStats",
    "PROG_ID",
    "PLOT_TYPES",
    # 异常
    "OriginError",
    "OriginNotFoundError",
    "OriginConnectionError",
    "OriginCommandError",
    "LanguageError",
    # 环境探测
    "find_origin_install",
    "find_origin_exe",
    "find_origin_versions",
    "index_to_letters",
    "is_origin_running",
    "iter_origin_reg_keys",
    # 语言切换
    "switch_language",
    "set_language",
    "get_language",
    "describe_language",
    "list_installations",
    "LanguageState",
    "LANGUAGE_NAMES",
    "normalize_language",
    # 进程管理
    "list_processes",
    "close_origin",
    "launch_origin",
    # 元信息
    "__version__",
]
