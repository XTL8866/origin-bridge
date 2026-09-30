# -*- coding: utf-8 -*-
"""
origin-bridge 冒烟测试 / 集成测试

**前置条件**：本机已安装 Origin，且 Origin 正在运行（或允许由 COM 启动）。

运行方式::

    # 直接运行（不依赖 pytest）
    python tests/test_smoke.py

    # 或用 pytest
    pytest tests/ -v

测试分为两组：
* **纯函数测试** —— 不需要 Origin，任何环境都能跑。
* **集成测试** —— 需要真实 Origin 实例，缺环境时自动跳过。
"""

from __future__ import annotations

import os
import statistics
import sys
import tempfile
import time

# 允许直接以脚本方式运行（把项目根目录加入 sys.path）
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from origin_bridge import (  # noqa: E402
    OriginBridge,
    OriginError,
    find_origin_exe,
    find_origin_install,
    find_origin_versions,
    list_installations,
    get_language,
    normalize_language,
    describe_language,
)
from origin_bridge.bridge import (  # noqa: E402
    NANUM_SENTINEL,
    OriginBridge as _OB,
    _to_float,
    is_nanum,
)


# --------------------------------------------------------------------------- #
# 测试框架（极简，兼容 pytest）
# --------------------------------------------------------------------------- #

_PASSED = []
_FAILED = []
_SKIPPED = []


def _run(name, fn):
    try:
        fn()
    except AssertionError as exc:
        _FAILED.append((name, "断言失败: %s" % exc))
        print("  ✗ %s —— %s" % (name, exc))
    except Exception as exc:  # noqa: BLE001
        _FAILED.append((name, "异常: %r" % exc))
        print("  ✗ %s —— %r" % (name, exc))
    else:
        _PASSED.append(name)
        print("  ✓ %s" % name)


def _skip(name, reason):
    _SKIPPED.append((name, reason))
    print("  - %s（跳过：%s）" % (name, reason))


# --------------------------------------------------------------------------- #
# 组一：纯函数测试（无需 Origin）
# --------------------------------------------------------------------------- #

def test_column_letter_conversion():
    """列序号与列字母的互转。"""
    conv = _OB._index_to_letters
    assert conv(0) == "A", conv(0)
    assert conv(1) == "B", conv(1)
    assert conv(25) == "Z", conv(25)
    assert conv(26) == "AA", conv(26)
    assert conv(27) == "AB", conv(27)
    assert conv(51) == "AZ", conv(51)
    assert conv(52) == "BA", conv(52)
    assert conv(701) == "ZZ", conv(701)
    assert conv(702) == "AAA", conv(702)


def test_column_index_parsing():
    """列标识解析（字母 / 序号 / 混合大小写）。"""
    ob = _OB()
    assert ob._column_index("A") == 0
    assert ob._column_index("a") == 0
    assert ob._column_index("B") == 1
    assert ob._column_index("AA") == 26
    assert ob._column_index(1) == 0
    assert ob._column_index(3) == 2
    assert ob._column_index("5") == 4
    # 非法输入应报错
    for bad in ("", None):
        try:
            ob._column_index(bad)
        except (ValueError, AttributeError):
            pass
        else:
            raise AssertionError("应对 %r 抛出异常" % (bad,))


def test_column_name_normalize():
    """列名标准化。"""
    norm = _OB._normalize_column_name
    assert norm(1) == "A"
    assert norm(2) == "B"
    assert norm(27) == "AA"
    assert norm("a") == "A"
    assert norm("  b  ") == "B"
    try:
        norm(0)
    except ValueError:
        pass
    else:
        raise AssertionError("列序号 0 应报错")


def test_value_formatting():
    """写入值格式化。"""
    fmt = _OB._fmt
    assert fmt(1) == "1"
    assert fmt(1.5) == "1.5"
    assert fmt(None) == "NANUM"
    assert fmt(True) == "1"
    assert fmt(False) == "0"
    assert fmt(float("nan")) == "NANUM"
    assert fmt("abc") == '"abc"'


def test_nanum_filtering():
    """缺失值(NANUM)哨兵应被过滤成 None，而非当成真实数据。"""
    # Origin 用 -1.23456789e-300 表示缺失值
    assert _to_float(NANUM_SENTINEL) is None, "哨兵值未被过滤"
    assert is_nanum(NANUM_SENTINEL)
    assert not is_nanum(1.0)
    assert not is_nanum(None)
    # 其他无效输入
    assert _to_float(None) is None
    assert _to_float("") is None
    assert _to_float("abc") is None
    assert _to_float(float("nan")) is None
    # 正常值不受影响
    assert _to_float(1.5) == 1.5
    assert _to_float("2.5") == 2.5
    assert _to_float(0) == 0.0
    assert _to_float(-3) == -3.0
    # 极小但非哨兵的值不应被误伤
    assert _to_float(-1.2e-300) is not None


def test_language_normalize():
    """语言别名归一化。"""
    for text in ("zh", "ZH", "中文", "简体", "C", "c", "cn", "chinese"):
        assert normalize_language(text) == "C", text
    for text in ("en", "EN", "英文", "英语", "E", "e", "english", "eng"):
        assert normalize_language(text) == "E", text
    try:
        normalize_language("klingon")
    except Exception:
        pass
    else:
        raise AssertionError("非法语言应报错")


def test_environment_probe():
    """环境探测函数应能在装有 Origin 的机器上给出结果。"""
    versions = find_origin_versions()
    print("      [信息] 注册表版本键: %s" % (versions or "（无）"))
    install = find_origin_install()
    print("      [信息] 安装目录: %s" % (install or "（未找到）"))
    exe = find_origin_exe()
    print("      [信息] 主程序: %s" % (exe or "（未找到）"))
    print("      [信息] 语言: %s" % describe_language())
    # 这些函数不应抛异常；若装了 Origin 则应能定位到
    if versions:
        assert install is not None, "有注册表记录却找不到安装目录"
        assert os.path.isdir(install), install


# --------------------------------------------------------------------------- #
# 组二：集成测试（需要 Origin）
# --------------------------------------------------------------------------- #

def _origin_available() -> bool:
    """探测能否连上 Origin。"""
    try:
        with OriginBridge() as ob:
            return bool(ob.info().name)
    except Exception:
        return False


def test_bridge_roundtrip():
    """完整往返：建表 → 写 → 读 → 统计 → 保存。"""
    # 注意：Origin 工作簿名上限约 13 字符，这里保持简短以避免截断警告
    book = "Smoke%d" % int(time.time() % 100000)
    vals_a = [float(i) for i in range(1, 11)]          # 1..10
    vals_b = [float(i * 3 + 1) for i in range(1, 11)]  # 4,7,...,31

    with OriginBridge() as ob:
        assert ob.connected, "连接后状态应为已连接"

        # 元信息
        info = ob.info()
        assert info.name, "应能取得实例名称"
        print("      [信息] %s" % info)

        # 建表 + 写数据
        actual_book = ob.new_workbook(book)
        assert actual_book, "应返回工作簿名称"
        print("      [信息] 新建工作簿: %s" % actual_book)

        ob.set_column("A", vals_a)
        ob.set_column("B", vals_b)

        # 读回
        got_a = ob.get_column("A", rows=10)
        got_b = ob.get_column("B", rows=10)
        assert got_a == vals_a, "A 列往返不一致: %r" % (got_a,)
        assert got_b == vals_b, "B 列往返不一致: %r" % (got_b,)

        # 统计
        st = ob.stats("B")
        assert st.count == len(vals_b), st.count
        assert abs((st.total or 0) - sum(vals_b)) < 1e-6, st
        assert abs((st.mean or 0) - sum(vals_b) / len(vals_b)) < 1e-6, st
        assert abs((st.maximum or 0) - max(vals_b)) < 1e-6, st
        assert abs((st.minimum or 0) - min(vals_b)) < 1e-6, st
        # 样本标准差：回归验证 —— 必须用 stddev()，无效的 sd() 会返回哨兵值
        expected_sd = statistics.stdev(vals_b)
        assert st.sd is not None, \
            "sd 为 None —— 哨兵值未被过滤（可能用了无效的函数名）"
        assert abs(st.sd - expected_sd) < 1e-6, \
            "sd=%r 期望 %r" % (st.sd, expected_sd)
        print("      [信息] %s" % st)

        # 二维读取
        data = ob.get_data()
        assert isinstance(data, list) and data, "get_data 应返回非空列表"
        assert len(data[0]) >= 2, "每行应至少包含 2 列: %r" % (data[0],)
        print("      [信息] get_data 返回 %d 行 × %d 列"
              % (len(data), len(data[0])))

        # 索引/字母两种写法应等价
        assert ob.get_column(1, rows=5) == ob.get_column("A", rows=5)

        # 保存（输出到系统临时目录，避免污染项目目录）
        out = os.path.join(tempfile.gettempdir(), "_origin_bridge_smoke.opju")
        assert ob.save(out), "保存应生成文件"
        assert os.path.isfile(out), out
        size = os.path.getsize(out)
        print("      [信息] 已保存 %s（%.1f KB）" % (out, size / 1024))
        assert size > 1024, "保存的文件过小，可能损坏"

        # 清理
        try:
            os.remove(out)
        except OSError:
            pass


def test_stats_dataclass():
    """ColumnStats 的格式化输出。"""
    with OriginBridge() as ob:
        book = "Stats%d" % int(time.time() % 100000)
        ob.new_workbook(book)
        ob.set_column("A", [2.0, 4.0, 6.0, 8.0, 10.0])
        st = ob.stats("A")
        text = str(st)
        assert "A" in text
        d = st.as_dict()
        assert set(d) == {"count", "total", "mean", "min", "max", "sd", "median"}
        assert abs(d["total"] - 30.0) < 1e-6, d
        assert abs(d["mean"] - 6.0) < 1e-6, d
        assert abs(d["min"] - 2.0) < 1e-6, d
        assert abs(d["max"] - 10.0) < 1e-6, d
        # stddev 是样本标准差：[2,4,6,8,10] → sqrt(40/4) ≈ 3.16228
        assert d["sd"] is not None, "sd 为 None——哨兵未被过滤"
        assert abs(d["sd"] - statistics.stdev([2, 4, 6, 8, 10])) < 1e-6, d["sd"]
        print("      [信息] %s" % text)


def test_connection_error_handling():
    """未连接时访问 app 属性应抛出明确异常。"""
    ob = OriginBridge()
    assert not ob.connected
    try:
        _ = ob.app
    except OriginError:
        pass
    else:
        raise AssertionError("未连接访问 app 应抛 OriginError")


# --------------------------------------------------------------------------- #
# 主流程
# --------------------------------------------------------------------------- #

def main() -> int:
    print()
    print("=" * 64)
    print("  origin-bridge 冒烟测试")
    print("=" * 64)

    print("\n[组一] 纯函数测试（无需 Origin）")
    for name, fn in [
        ("列字母互转", test_column_letter_conversion),
        ("列标识解析", test_column_index_parsing),
        ("列名标准化", test_column_name_normalize),
        ("数值格式化", test_value_formatting),
        ("缺失值过滤", test_nanum_filtering),
        ("语言别名归一化", test_language_normalize),
        ("环境探测", test_environment_probe),
        ("连接异常处理", test_connection_error_handling),
    ]:
        _run(name, fn)

    print("\n[组二] 集成测试（需要真实 Origin 实例）")
    if not _origin_available():
        for name in ("数据往返", "统计格式化"):
            _skip(name, "无法连接 Origin")
    else:
        print("  （已连上 Origin）")
        _run("数据往返（建表/写/读/统计/保存）", test_bridge_roundtrip)
        _run("统计格式化", test_stats_dataclass)

    print()
    print("=" * 64)
    print("  通过 %d 项   失败 %d 项   跳过 %d 项"
          % (len(_PASSED), len(_FAILED), len(_SKIPPED)))
    if _FAILED:
        print()
        print("  失败明细：")
        for name, why in _FAILED:
            print("    · %s —— %s" % (name, why))
    print("=" * 64)
    print()
    return 1 if _FAILED else 0


# pytest 兼容：把每个 test_* 暴露出去
def test_all():  # pragma: no cover
    assert main() == 0


if __name__ == "__main__":
    sys.exit(main())
