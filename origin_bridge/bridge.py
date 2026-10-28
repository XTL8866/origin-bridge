# -*- coding: utf-8 -*-
"""
origin_bridge.bridge — Origin COM 自动化桥接核心

通过 Windows COM 接口控制 OriginLab Origin，实现数据读写、统计计算、
绘图与项目保存等自动化操作。

支持版本：Origin 2018 ~ 2026（凡注册了 Origin.ApplicationSI 均可）

技术路线说明
------------
本库刻意**不使用** OriginLab 官方的 ``originpro`` PyPI 包，原因是在
实测中该包的 ``attach()`` 可以连上实例，但 ``get_lt_str()`` 等取值接口
一律返回空字符串，无法读取真实数据。COM 直连（win32com）稳定可靠。

踩坑记录（均为实测结论，供后来者参考）
--------------------------------------
1. **ProgID 用 ``Origin.ApplicationSI``**。
   ``Origin.ApplicationSI64`` 在多数安装中**并未注册**，直接 Dispatch 会失败。

2. **读值只用 ``Evaluate``，不要用 ``LTVar``**。
   ``app.LTVar("col(B)[1]")`` 恒返回 ``0.0``（不报错，静默错误，极具迷惑性）；
   ``app.LTStr("$(col(B)[1])")`` 返回空串。
   正确做法：``app.Evaluate("col(B)[1]")``，或统计量 ``app.Evaluate("total(col(B))")``。

3. **``LTStr`` 只用于字符串表达式**，例如 ``page.name$``、``wks.col1.name$``。

4. **批量读取用 ``GetWorksheet(name)``**，一次拿回整表，
   返回 row-major 元组 ``((A1,B1), (A2,B2), ...)``；
   第二个参数是**起始行偏移**（0-based），默认 0。
   逐行 ``Evaluate`` 虽然也可用，但行数多时明显更慢。

5. **批量写入用 ``Execute`` 拼接命令**，不要用 ``PutWorksheet``
   （实测写入位置错乱，会写出 ``-1.23456789e-300`` 这类垃圾值）。

6. **保存项目用 ``app.Save(绝对路径)``**。
   LabTalk 的 ``save -i "路径"`` 执行不报错，但**不生成文件**。

7. **``app.Path`` 是方法不是属性**，直接访问会抛「类型不匹配」异常。
   安装路径请从注册表读取，见 :func:`find_origin_install`。

8. **首次 Dispatch 约需 3~7 秒**（进程启动 + 类型库加载），此后单次调用
   通常在毫秒级。不要在循环里反复建立连接。

9. **``Execute`` 可能长时间阻塞**（例如打开了模态对话框时）。
   在自动化脚本里建议后台运行并落盘日志，避免超时被杀。
"""

from __future__ import annotations

import os
import subprocess
import time
import warnings
import winreg
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Sequence

try:
    import win32com.client as _win32
except ImportError as exc:  # pragma: no cover
    raise ImportError(
        "缺少 pywin32 依赖，请先安装：pip install pywin32"
    ) from exc


# --------------------------------------------------------------------------- #
# 常量
# --------------------------------------------------------------------------- #

#: Origin 的 COM ProgID。注意：ApplicationSI64 通常未注册，不要使用。
PROG_ID = "Origin.ApplicationSI"

#: Origin 在注册表中的根路径（HKEY_CURRENT_USER 下）
REG_ROOT = r"Software\OriginLab"

#: Origin 可执行文件名（按优先级）
ORIGIN_EXES = ("Origin64.exe", "Origin.exe")

#: LabTalk 绘图类型编号 → 名称
PLOT_TYPES: Dict[str, int] = {
    "line": 200,          # 线图
    "scatter": 201,       # 散点图
    "line_symbol": 202,   # 点线图
    "column": 203,        # 柱状图
    "bar": 204,           # 条形图
    "area": 205,          # 面积图
}

#: 图片扩展名 → ``expGraph`` 的 ``type`` 参数值。
#:
#: 注意：``type`` 接受的是**字符串**而非序号。序号 1 对应 BMP 而不是 PNG，
#: 传错会静默导出成别的格式。
EXPORT_TYPES: Dict[str, str] = {
    ".png": "png",
    ".jpg": "jpg",
    ".jpeg": "jpg",
    ".tif": "tif",
    ".tiff": "tif",
    ".bmp": "bmp",
    ".gif": "gif",
    ".eps": "eps",
    ".pdf": "pdf",
    ".svg": "svg",
    ".emf": "emf",
    ".wmf": "wmf",
    ".pcx": "pcx",
    ".tga": "tga",
    ".psd": "psd",
}

#: Origin 内部标识「缺失值」(NANUM) 的浮点哨兵值。
#:
#: COM 接口不会把缺失值转成 Python 的 ``None``，而是返回这个魔数。
#: 任何读取数值的接口都要过滤它，否则会把哨兵当成真实数据。
NANUM_SENTINEL: float = -1.23456789e-300

#: LabTalk 中可用的描述性统计函数名。
#:
#: .. warning::
#:    标准差函数是 ``stddev()``。``sd()`` / ``std()`` / ``stdev()`` / ``SD()``
#:    这些看似合理的名字在 LabTalk 中**都不存在**，且调用时**不报错**，
#:    只会静默返回哨兵值 —— 是最容易踩的坑之一。
STAT_FUNCTIONS: Dict[str, str] = {
    "count": "count",     # 有效数据个数
    "total": "total",     # 求和
    "mean": "mean",       # 算术平均
    "minimum": "min",     # 最小值
    "maximum": "max",     # 最大值
    "sd": "stddev",       # 样本标准差（注意不是 sd）
    "median": "median",   # 中位数
}

#: 注册表中存放安装路径与构建号的子键（**不在版本键本身，而在其下的 International 层**）
_INTERNATIONAL_SUBKEY = "International"


# --------------------------------------------------------------------------- #
# 异常
# --------------------------------------------------------------------------- #

class OriginError(Exception):
    """Origin 相关错误的基类。"""


class OriginNotFoundError(OriginError):
    """系统中未找到 Origin 安装。"""


class OriginConnectionError(OriginError):
    """无法建立 COM 连接。"""


class OriginCommandError(OriginError):
    """LabTalk 命令执行失败。"""


# --------------------------------------------------------------------------- #
# 数据结构
# --------------------------------------------------------------------------- #

@dataclass
class OriginInfo:
    """Origin 实例的基本信息。"""

    name: str = ""                        # 例如 "OriginPro 2025b"
    prog_id: str = PROG_ID
    install_path: Optional[str] = None    # 安装目录
    build: Optional[str] = None           # 内部版本号

    def __str__(self) -> str:
        parts = [self.name or "(未知)"]
        if self.build:
            parts.append("build=%s" % self.build)
        if self.install_path:
            parts.append("path=%s" % self.install_path)
        return "  ".join(parts)


@dataclass
class ColumnStats:
    """单列的描述性统计量。取不到的项为 None。"""

    name: str
    count: Optional[float] = None
    total: Optional[float] = None
    mean: Optional[float] = None
    minimum: Optional[float] = None
    maximum: Optional[float] = None
    sd: Optional[float] = None
    median: Optional[float] = None

    def as_dict(self) -> Dict[str, Optional[float]]:
        return {
            "count": self.count,
            "total": self.total,
            "mean": self.mean,
            "min": self.minimum,
            "max": self.maximum,
            "sd": self.sd,
            "median": self.median,
        }

    def __str__(self) -> str:
        d = self.as_dict()
        items = ["%s=%s" % (k, ("%.6g" % v if isinstance(v, float) else v))
                 for k, v in d.items()]
        return "列 %s: %s" % (self.name, "  ".join(items))


# --------------------------------------------------------------------------- #
# 环境探测辅助函数
# --------------------------------------------------------------------------- #

def iter_origin_reg_keys():
    """遍历 HKCU\\Software\\OriginLab 下的所有子键名（如 'Origin 10.2b'）。"""
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_ROOT) as root:
            n_sub, _, _ = winreg.QueryInfoKey(root)
            for i in range(n_sub):
                try:
                    yield winreg.EnumKey(root, i)
                except OSError:
                    continue
    except FileNotFoundError:
        return


def index_to_letters(idx: int) -> str:
    """
    0-based 列下标 → Excel 风格列字母。

    >>> index_to_letters(0)
    'A'
    >>> index_to_letters(25)
    'Z'
    >>> index_to_letters(26)
    'AA'
    """
    letters = ""
    idx += 1
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(ord("A") + rem) + letters
    return letters


def _read_version_value(key_name: str, value_name: str) -> Any:
    """
    从指定版本键中读取某个注册表值。

    注册表布局是**分层**的，这点容易踩坑：

    * ``Language``  位于版本键本身 —— ``...\\Origin 10.2b``
    * ``Executable Path`` / ``Build`` / ``Path`` 位于其下的 ``International`` 子键

    所以两个位置都要查，顺序为「先本层，后 International」。
    """
    for sub in ("", _INTERNATIONAL_SUBKEY):
        path = "%s\\%s" % (REG_ROOT, key_name)
        if sub:
            path = "%s\\%s" % (path, sub)
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as k:
                try:
                    value, _ = winreg.QueryValueEx(k, value_name)
                    return value
                except FileNotFoundError:
                    continue
        except OSError:
            continue
    return None


def find_origin_install() -> Optional[str]:
    """
    在注册表中查找 Origin 的安装目录。

    之所以要读注册表：``app.Path`` 是方法而非属性，调用会抛「类型不匹配」，
    因此无法从 COM 对象直接拿到安装路径。

    Returns:
        安装目录绝对路径；找不到返回 ``None``。
    """
    for key_name in iter_origin_reg_keys():
        if not key_name.lower().startswith("origin"):
            continue
        path = _read_version_value(key_name, "Executable Path")
        if not path:
            continue
        path = os.path.normpath(str(path))
        if os.path.isdir(path):
            return path
    return None


def find_origin_versions() -> List[str]:
    """返回注册表中记录的所有 Origin 版本键名，如 ``['Origin 10.2b']``。"""
    return [k for k in iter_origin_reg_keys() if k.lower().startswith("origin")]


def find_origin_exe() -> Optional[str]:
    """在安装目录中定位 Origin 主程序，返回其绝对路径。"""
    install = find_origin_install()
    if not install:
        return None
    for exe in ORIGIN_EXES:
        candidate = os.path.join(install, exe)
        if os.path.isfile(candidate):
            return candidate
    return None


def is_origin_running() -> bool:
    """检查是否已有 Origin 进程在运行（不含判断是否可自动化）。"""
    try:
        out = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq Origin64.exe", "/FO", "CSV", "/NH"],
            capture_output=True, text=True, timeout=10,
            encoding="gbk", errors="replace",
        )
        return "Origin64.exe" in (out.stdout or "")
    except Exception:
        return False


def is_nanum(value: Any) -> bool:
    """判断一个值是否为 Origin 的缺失值(NANUM)哨兵。"""
    try:
        return float(value) == NANUM_SENTINEL
    except (TypeError, ValueError):
        return False


def _to_float(value: Any) -> Optional[float]:
    """
    尽量把 COM 返回值转成 float。

    Origin 用 ``-1.23456789e-300`` 这个哨兵值表示缺失值(NANUM)，此处统一
    转换成 Python 的 ``None``，避免把哨兵当成真实数据参与计算。

    :return: 浮点数；无法解析或为缺失值时返回 ``None``。
    """
    if value is None:
        return None
    if isinstance(value, bool):
        return float(value)
    if not isinstance(value, (int, float)):
        text = str(value).strip()
        if not text:
            return None
        try:
            value = float(text)
        except (TypeError, ValueError):
            return None
    result = float(value)
    if result != result:              # NaN
        return None
    if result == NANUM_SENTINEL:      # Origin 缺失值哨兵
        return None
    return result


# --------------------------------------------------------------------------- #
# 核心桥接类
# --------------------------------------------------------------------------- #

class OriginBridge:
    """
    Origin COM 自动化桥接。

    典型用法::

        from origin_bridge import OriginBridge

        with OriginBridge() as ob:
            ob.new_workbook("MyData")
            ob.set_column("A", [1, 2, 3, 4, 5])
            ob.set_column("B", [2, 4, 6, 8, 10])
            print(ob.get_column("A"))
            print(ob.stats("B"))
            ob.save(r"D:\\data\\out.opju")

    :param prog_id: COM ProgID，默认 ``Origin.ApplicationSI``。
    :param visible: 是否让 Origin 窗口可见（对已有实例无效）。
    :param timeout: 连接超时秒数（仅作提示，COM 本身不支持超时）。
    """

    def __init__(self, prog_id: str = PROG_ID, visible: bool = True,
                 timeout: float = 60.0):
        self.prog_id = prog_id
        self.visible = visible
        self.timeout = timeout
        self._app: Optional[Any] = None

    # ------------------------------------------------------------------ #
    # 连接管理
    # ------------------------------------------------------------------ #

    @property
    def app(self) -> Any:
        """底层 COM 对象。未连接时抛 :class:`OriginConnectionError`。"""
        if self._app is None:
            raise OriginConnectionError(
                "尚未连接 Origin，请先调用 connect() 或使用 with 语句。"
            )
        return self._app

    @property
    def connected(self) -> bool:
        return self._app is not None

    def connect(self) -> "OriginBridge":
        """
        建立与 Origin 的连接。

        若已有 Origin 实例在运行，通常会自动附着；否则由 COM 启动一个新实例。
        首次连接约需 3~7 秒。

        :raises OriginConnectionError: 连接失败。
        """
        if self._app is not None:
            return self
        try:
            self._app = _win32.Dispatch(self.prog_id)
        except Exception as exc:
            hint = ""
            if not is_origin_running():
                exe = find_origin_exe()
                hint = ("\n  未检测到运行中的 Origin。请先手动启动 Origin"
                        + ("（%s）" % exe if exe else "，或确认已正确安装。"))
            raise OriginConnectionError(
                "无法连接 Origin（ProgID=%s）：%s%s" % (self.prog_id, exc, hint)
            ) from exc
        return self

    def disconnect(self) -> None:
        """断开引用（不会关闭 Origin 进程）。"""
        self._app = None

    def close(self) -> None:
        """请求 Origin 退出。谨慎使用，未保存的工作会丢失。"""
        if self._app is not None:
            try:
                self._app.Exit()
            except Exception:
                pass
        self._app = None

    def __enter__(self) -> "OriginBridge":
        return self.connect()

    def __exit__(self, exc_type, exc, tb) -> None:
        self.disconnect()

    def __repr__(self) -> str:
        state = "已连接" if self.connected else "未连接"
        return "<OriginBridge %s prog_id=%r>" % (state, self.prog_id)

    # ------------------------------------------------------------------ #
    # 元信息
    # ------------------------------------------------------------------ #

    def info(self) -> OriginInfo:
        """获取 Origin 名称、内部版本号与安装路径。"""
        data = OriginInfo(prog_id=self.prog_id)
        try:
            data.name = self.app.Name or ""
        except Exception:
            pass
        # 构建号与安装路径只能从注册表取：
        # app.Path 是方法而非属性，调用会抛「类型不匹配」。
        for key_name in find_origin_versions():
            build = _read_version_value(key_name, "Build")
            if build is not None:
                data.build = str(build)
                break
        data.install_path = find_origin_install()
        return data

    # ------------------------------------------------------------------ #
    # 底层执行
    # ------------------------------------------------------------------ #

    def execute(self, command: str, wait: float = 0.3) -> None:
        """
        执行 LabTalk 命令。

        :param command: LabTalk 命令，可分号分隔多条。
        :param wait: 命令执行后等待秒数，给 Origin 留出处理时间。
        """
        try:
            self.app.Execute(command)
        except Exception as exc:
            raise OriginCommandError(
                "LabTalk 命令执行失败：%r\n  命令：%s" % (exc, command[:200])
            ) from exc
        if wait:
            time.sleep(wait)

    def evaluate(self, expression: str) -> Optional[float]:
        """
        求值一个 LabTalk 数值表达式。

        **这是读取数值/统计量的唯一正确方式**。不要用 ``LTVar``（恒返回 0.0）
        或 ``LTStr("$(...)")``（返回空串）。

        :param expression: 例如 ``"col(B)[1]"``、``"total(col(B))"``。
        """
        try:
            return _to_float(self.app.Evaluate(expression))
        except Exception:
            return None

    def get_text(self, expression: str) -> str:
        """
        读取字符串表达式。

        **仅适用于字符串**，例如 ``"page.name$"``、``"wks.col1.name$"``。
        数值请改用 :meth:`evaluate`。
        """
        try:
            return self.app.LTStr(expression) or ""
        except Exception:
            return ""

    # ------------------------------------------------------------------ #
    # 工作表定位
    # ------------------------------------------------------------------ #

    def active_book(self) -> str:
        """返回当前活动页面的名称。"""
        return self.get_text("page.name$")

    def list_books(self) -> List[str]:
        """列出项目中的所有页面名称。"""
        names: List[str] = []
        try:
            pages = self.app.Pages
            for i in range(pages.Count):
                try:
                    names.append(pages.Item(i).Name)
                except Exception:
                    continue
        except Exception:
            pass
        if not names:
            cur = self.active_book()
            if cur:
                names.append(cur)
        return names

    def _resolve_book(self, book: Optional[str]) -> str:
        """把 book 参数解析成实际名称，None 表示当前活动页。"""
        if book:
            return book
        cur = self.active_book()
        if not cur:
            raise OriginCommandError("无法确定当前工作表，请显式指定 book 参数。")
        return cur

    # ------------------------------------------------------------------ #
    # 读取数据
    # ------------------------------------------------------------------ #

    def get_data(self, book: Optional[str] = None,
                 start_row: int = 0) -> List[List[float]]:
        """
        一次性读回整张工作表，返回**按行组织**的二维列表。

        :param book: 工作表/工作簿名称，``None`` 表示当前活动页。
        :param start_row: 起始行偏移（0 表示从第一行开始）。
        :return: 形如 ``[[A1, B1], [A2, B2], ...]``。

        实现说明：``GetWorksheet(name, offset)`` 直接返回 row-major 元组，
        比逐行 ``Evaluate`` 快得多。
        """
        target = self._resolve_book(book)
        try:
            raw = self.app.GetWorksheet(target, start_row)
        except Exception as exc:
            raise OriginCommandError(
                "读取工作表 %r 失败：%r" % (target, exc)
            ) from exc
        if raw is None:
            return []
        data: List[List[float]] = []
        for row in raw:
            if isinstance(row, (tuple, list)):
                data.append([_to_float(v) for v in row])
            else:
                data.append([_to_float(row)])
        return data

    def get_column(self, column: Any = 1, book: Optional[str] = None,
                   rows: Optional[int] = None) -> List[float]:
        """
        读回一列数据。

        :param column: 列序号（1-based，int）或列字母（str，如 ``"A"``）。
        :param book: 工作表名称，``None`` 表示当前活动页。
        :param rows: 只取前 N 行；``None`` 表示全部。
        :return: 数值列表（缺失值以 ``None`` 占位）。
        """
        idx = self._column_index(column, book)
        data = self.get_data(book)
        out = [row[idx] if idx < len(row) else None for row in data]
        return out[:rows] if rows else out

    def _column_index(self, column: Any, book: Optional[str] = None) -> int:
        """把列字母或序号统一转成 0-based 下标。"""
        if column is None:
            raise ValueError("列标识不能为空")
        if isinstance(column, int):
            return max(0, column - 1)
        text = str(column).strip().upper()
        if not text:
            raise ValueError("列标识不能为空")
        # 列字母 → 0-based（A=0, B=1, ..., Z=25, AA=26）
        if text.isalpha():
            idx = 0
            for ch in text:
                idx = idx * 26 + (ord(ch) - ord("A") + 1)
            return idx - 1
        try:
            return max(0, int(text) - 1)
        except ValueError:
            raise ValueError("无法识别的列标识：%r" % column)

    # ------------------------------------------------------------------ #
    # 写入数据
    # ------------------------------------------------------------------ #

    def set_column(self, column: Any, values: Sequence[Any],
                   start_row: int = 1, batch: int = 200,
                   wait: float = 0.5) -> int:
        """
        把一组数值写入指定列。

        :param column: 列序号（1-based）或列字母。
        :param values: 要写入的数值序列。
        :param start_row: 起始行号（1-based）。
        :param batch: 每条 LabTalk 命令包含的最大赋值数，防止命令过长。
        :param wait: 全部写入完成后等待秒数。
        :return: 实际写入的元素个数。

        实现说明：用 ``Execute`` 拼接 ``col(X)[i]=v;`` 批量赋值。
        实测 100 个值仅需约 1 秒。**不要用 PutWorksheet**，其写入位置不可控。
        """
        col = self._normalize_column_name(column)
        values = list(values)
        total = 0
        for i in range(0, len(values), batch):
            chunk = values[i:i + batch]
            cmds = " ".join(
                "col(%s)[%d]=%s;" % (col, start_row + i + j, self._fmt(v))
                for j, v in enumerate(chunk)
            )
            self.execute(cmds, wait=0)
            total += len(chunk)
        if wait:
            time.sleep(wait)
        return total

    @staticmethod
    def _fmt(value: Any) -> str:
        """把 Python 值格式化成 LabTalk 能接受的字面量。"""
        if value is None:
            return "NANUM"          # Origin 的缺失值常量
        if isinstance(value, bool):
            return "1" if value else "0"
        if isinstance(value, (int, float)):
            if isinstance(value, float) and (value != value):   # NaN
                return "NANUM"
            return repr(value)
        text = str(value).strip()
        # 非数值内容按字符串处理（写入文本列）
        return '"%s"' % text.replace('"', '\\"')

    @staticmethod
    def _normalize_column_name(column: Any) -> str:
        """把列标识统一成 LabTalk 的列名（如 'A'、'B'）。"""
        if isinstance(column, int):
            if column < 1:
                raise ValueError("列序号必须 >= 1")
            return OriginBridge._index_to_letters(column - 1)
        text = str(column).strip().upper()
        if not text:
            raise ValueError("列标识不能为空")
        return text

    @staticmethod
    def _index_to_letters(idx: int) -> str:
        """0-based 下标 → Excel 风格列字母（保留作向后兼容，公开接口请用
        :func:`index_to_letters`）。"""
        return index_to_letters(idx)

    # ------------------------------------------------------------------ #
    # 工作簿操作
    # ------------------------------------------------------------------ #

    def new_workbook(self, name: str, columns: int = 2,
                     wait: float = 1.5) -> str:
        """
        新建工作簿。

        .. note::
            Origin 对工作簿名称有**长度上限**（实测约 13 个字符），
            超长部分会被**静默截断** —— 传入 ``"SignalAnalysis"``
            实际得到 ``"SignalAnalysi"``。本方法会在检测到截断时
            发出 ``UserWarning``。

        :return: **实际**创建的工作簿名称（可能与传入的 ``name`` 不同，
                 建议后续操作使用本返回值）。
        """
        self.execute('newbook name:="%s" sheet:=1 option:=1;' % name, wait=wait)
        actual = self.active_book()
        if actual and name and actual != name:
            warnings.warn(
                "工作簿名称被 Origin 截断：请求 %r，实际 %r"
                "（名称长度上限约 13 字符）" % (name, actual),
                UserWarning,
                stacklevel=2,
            )
        return actual

    def set_column_label(self, column: Any, label: str,
                         kind: str = "long") -> None:
        """
        设置列标签。

        :param kind: ``"long"`` 长名称 / ``"unit"`` 单位 / ``"comment"`` 注释。
        """
        idx = self._column_index(column)
        prop = {"long": "lname", "unit": "units", "comment": "comment"}.get(
            kind.lower(), "lname"
        )
        self.execute('wks.col%d.%s$="%s";' % (idx + 1, prop, label), wait=0)

    def clear_book(self, book: Optional[str] = None) -> None:
        """清空指定工作表的全部数据。"""
        target = self._resolve_book(book)
        self.execute('win -a %s;' % target, wait=0.3)
        self.execute('del -r;', wait=0.5)

    # ------------------------------------------------------------------ #
    # 统计
    # ------------------------------------------------------------------ #

    def stats(self, column: Any, book: Optional[str] = None) -> ColumnStats:
        """
        计算一列的描述性统计量。

        实现说明（两个必须避开的坑）：

        * 必须用 ``Evaluate("total(col(B))")`` 形式；``LTVar`` 恒返回 0.0，
          ``LTStr("$(...)")`` 返回空串。另外 ``sum()`` 并非求和函数
          （实测 ``sum(col(A))`` 返回的是别的东西），求和请用 ``total()``。
        * 标准差函数是 ``stddev()`` 而非 ``sd()``。``sd()`` / ``std()`` /
          ``stdev()`` 在 LabTalk 中都不存在，调用时**不报错**，只会静默
          返回缺失值哨兵，极易误判为真实结果。
        """
        col = self._normalize_column_name(column)
        result = ColumnStats(name=col)
        for attr, func in STAT_FUNCTIONS.items():
            if not hasattr(result, attr):
                continue
            setattr(result, attr, self.evaluate("%s(col(%s))" % (func, col)))
        return result

    # ------------------------------------------------------------------ #
    # 绘图
    # ------------------------------------------------------------------ #

    def plot(self, x_column: Any, y_column: Any,
             kind: str = "line", wait: float = 2.0) -> None:
        """
        用两列数据绘制图形。

        :param x_column: X 列（序号或字母）。
        :param y_column: Y 列（序号或字母）。
        :param kind: ``line`` / ``scatter`` / ``line_symbol`` / ``column`` /
                     ``bar`` / ``area``。
        """
        xi = self._column_index(x_column) + 1
        yi = self._column_index(y_column) + 1
        plot_id = PLOT_TYPES.get(kind.lower())
        if plot_id is None:
            raise ValueError(
                "不支持的图形类型 %r，可选：%s" % (kind, ", ".join(sorted(PLOT_TYPES)))
            )
        self.execute(
            "plotxy iy:=(%d,%d) plot:=%d ogl:=[<new>];" % (xi, yi, plot_id),
            wait=wait,
        )

    # ------------------------------------------------------------------ #
    # 保存与导出
    # ------------------------------------------------------------------ #

    def save(self, path: str, wait: float = 2.5) -> bool:
        """
        保存项目为 ``.opju`` / ``.opj`` 文件。

        :param path: 目标绝对路径。
        :return: 文件是否成功生成。

        实现说明：必须用 COM 的 ``Save`` 方法；LabTalk 的 ``save -i`` 无效。
        """
        path = os.path.abspath(path)
        parent = os.path.dirname(path)
        if parent and not os.path.isdir(parent):
            os.makedirs(parent, exist_ok=True)
        try:
            self.app.Save(path)
        except Exception as exc:
            raise OriginCommandError("保存项目失败：%r" % (exc,)) from exc
        if wait:
            time.sleep(wait)
        return os.path.exists(path)

    def export_ascii(self, path: str, book: Optional[str] = None,
                     wait: float = 2.0) -> bool:
        """
        把当前工作表的全部数据导出为 ASCII（制表符分隔）文本。

        :return: 文件是否成功生成。

        实现说明：LabTalk 的 ``expASC`` 在本环境中未能生成文件，
        因此这里改用 ``GetWorksheet`` 读数据 + Python 写文件的方式，
        行为更可控（附带好处：可指定列数与数值格式）。
        """
        path = os.path.abspath(path)
        if not path.lower().endswith(".txt"):
            path += ".txt"
        parent = os.path.dirname(path)
        if parent and not os.path.isdir(parent):
            os.makedirs(parent, exist_ok=True)
        data = self.get_data(book)
        with open(path, "w", encoding="utf-8", newline="") as fh:
            for row in data:
                fh.write("\t".join(
                    "" if v is None else ("%.10g" % v) for v in row
                ) + "\n")
        return os.path.exists(path)

    def export_image(self, path: str, page: Optional[str] = None,
                     width: Optional[float] = None, unit: int = 2,
                     dpi: Optional[float] = None,
                     wait: float = 2.5) -> bool:
        """
        把图形窗口导出为图片。

        :param path: 目标路径，扩展名决定格式（``.png`` / ``.jpg`` / ``.tif``…
                     ，受 :data:`EXPORT_TYPES` 支持）。
        :param page: 要导出的图形窗口名；``None`` 表示当前活动窗口。
        :param width: 导出宽度；``None`` 表示沿用 Origin 当前设置。
        :param unit: ``width`` 的单位。``0`` 英寸 / ``1`` 厘米 / ``2`` 像素 / ``3`` 页面比例。
        :param dpi: 位图分辨率（仅对 png/jpg/tif 等栅格格式有效）。
        :return: 文件是否成功生成。

        .. warning::
            ``expGraph`` 的 ``type`` 是**字符串**（``png`` / ``tif``），不是序号；
            且 ``path`` 必须是**目录**、文件名由 ``filename`` 单独给出。
            旧实现把完整文件路径塞进 ``path`` 并传 ``type:=1``，结果既不报错
            也不生成文件（``type:=1`` 实际是 BMP）。此处已修正。
        """
        path = os.path.abspath(path)
        folder, filename = os.path.split(path)
        stem, ext = os.path.splitext(filename)
        img_type = EXPORT_TYPES.get(ext.lower())
        if img_type is None:
            raise ValueError(
                "不支持的图片格式 %r，可选：%s"
                % (ext or "(无扩展名)", ", ".join(sorted(EXPORT_TYPES)))
            )
        if folder and not os.path.isdir(folder):
            os.makedirs(folder, exist_ok=True)
        # 先删目标文件：若已存在，expGraph 可能弹出覆盖确认框而阻塞
        if os.path.exists(path):
            try:
                os.remove(path)
            except OSError:
                pass

        if page:
            self.execute("win -a %s;" % page, wait=0.8)

        cmd = 'expGraph type:=%s path:="%s" filename:="%s"' % (
            img_type, folder, stem
        )
        if width:
            cmd += " tr1.unit:=%d tr1.width:=%g" % (unit, width)
        if dpi:
            cmd += " tr2.%s.dotsperinch:=%g" % (img_type, dpi)
        self.execute(cmd + ";", wait=wait)
        return os.path.exists(path)

