# origin-bridge

> 用 Python 通过 COM 自动化控制 **OriginLab Origin** —— 数据读写、统计计算、绘图、项目保存，以及界面语言切换。

[![Python](https://img.shields.io/badge/python-3.8%2B-blue.svg)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)]()
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

<p align="center">
  <b>简体中文</b> · <a href="#english">English</a>
</p>

---

## 这是什么

`origin-bridge` 把 Origin 的命令行接口（LabTalk）与 COM 自动化封装成一套清爽的 Python API，
让你能用脚本完成原本需要在图形界面里一步步点的工作：

```python
from origin_bridge import OriginBridge

with OriginBridge() as ob:
    ob.new_workbook("MyData")
    ob.set_column("A", [1, 2, 3, 4, 5])
    ob.set_column("B", [2, 4, 6, 8, 10])
    print(ob.stats("B"))          # 列 B: count=5 total=30 mean=6 ...
    ob.save(r"D:\data\demo.opju")
```

顺带解决一个长期痛点：**Origin 没有图形化的语言切换菜单**，
本项目提供了命令行与批处理两种切换方式。

---

## 为什么不用官方的 `originpro` 包

OriginLab 在 PyPI 上发布过官方包 `originpro`。但实测（Origin 2025b + Python 3.13）中：

| 接口 | 官方包的表现 |
| --- | --- |
| `op.attach()` | 能连上实例，返回 `None`（无报错） |
| `op.get_lt_str("system.version")` | **返回空字符串** |
| `op.ltstr(...)` | **返回空字符串** |

也就是说，连接"看似成功"但**取不到任何真实数据**，而且不报错，很难排查。

本项目改用 COM 直连（`win32com`），所有接口都经过真机验证，行为确定。
当然，如果你在别的环境里官方包工作正常，那就用官方的 —— 本项目只是提供了一个
经过充分验证的替代方案，并把踩过的坑记录在案。

---

## 特性

- **数据读写** — 按列或整表读写，自动处理 Origin 的缺失值（NANUM）
- **描述性统计** — 计数、求和、均值、极值、样本标准差、中位数，一次取全
- **绘图** — 线图 / 散点图 / 点线图 / 柱状图等，一行代码完成
- **项目保存与导出** — 保存 `.opju`、导出 ASCII 数据、导出图片
- **界面语言切换** — 中英文一键切换，附命令行与批处理两种入口
- **环境自检** — 自动定位安装路径、构建号、运行状态、语言设置
- **零配置** — 不需要装 Origin 的 Python 支持包，不需要管理员权限

---

## 环境要求

| 项目 | 要求 |
| --- | --- |
| 操作系统 | **Windows**（COM 技术仅 Windows 可用） |
| Origin | 2018 ~ 2026 之间任意版本（需注册 `Origin.ApplicationSI`） |
| Python | 3.8 及以上 |
| 依赖 | [`pywin32`](https://pypi.org/project/pywin32/) |

> **关于版本兼容性**：本项目通过 ProgID 查找 Origin，不依赖具体版本号，
> 因此在所有注册了 `Origin.ApplicationSI` 的版本上都能工作。
> 已实测版本：Origin 2025b。其他版本欢迎反馈。

---

## 安装

```bash
# 从源码安装（推荐，便于修改）
git clone https://github.com/XTL8866/origin-bridge.git
cd origin-bridge
pip install -e .
```

只装依赖也是可以的：

```bash
pip install pywin32
```

安装完成后会获得 `origin-cli` 命令。

---

## 快速开始

### 1. 检查环境

```bash
origin-cli info
```

输出示例：

```
============================================================
  Origin 环境信息
============================================================
  注册表中共有 1 个安装记录：
    · Origin 10.2b: 中文（简体）  build=212  D:\Program Files\OriginLab\Origin2025b

  ✓ Origin 正在运行（PID: 16984）

  当前语言：Origin 10.2b  →  中文（简体）

  ✓ COM 连接成功
  实例名称：OriginPro 2025b
  当前工作表：Book1
```

### 2. 跑一遍示例

```bash
python examples/01_quickstart.py
```

### 3. 导入数据

```bash
origin-cli import data.csv --book MyData
origin-cli stats A B
origin-cli plot A B --type scatter
origin-cli save result.opju
```

---

## API 文档

### `OriginBridge`

核心类。支持上下文管理器，退出时自动断开引用（**不会**关闭 Origin 进程）。

```python
from origin_bridge import OriginBridge

with OriginBridge() as ob:
    ...
```

#### 连接与环境

| 方法 | 说明 |
| --- | --- |
| `connect()` | 建立连接。首次约需 3~7 秒 |
| `disconnect()` | 断开引用（不关闭 Origin） |
| `close()` | 请求 Origin 退出（**未保存的工作会丢失**） |
| `info()` | 返回 `OriginInfo`：名称、构建号、安装路径 |
| `connected` | 属性，当前是否已连接 |
| `app` | 底层 COM 对象 |

#### 工作表操作

| 方法 | 说明 |
| --- | --- |
| `new_workbook(name, columns=2)` | 新建工作簿，返回其名称 |
| `active_book()` | 当前活动页名称 |
| `list_books()` | 列出所有页面名称 |
| `set_column_label(col, label, kind)` | 设置列标签（`long`/`unit`/`comment`） |
| `clear_book(book=None)` | 清空工作表数据 |

#### 数据读写

| 方法 | 说明 |
| --- | --- |
| `set_column(col, values, start_row=1)` | 写入一列。`col` 可用 `"A"` 或 `1` |
| `get_column(col, book=None, rows=None)` | 读回一列，返回 `list[float]` |
| `get_data(book=None)` | 读回整表，返回按行组织的 `list[list[float]]` |
| `export_ascii(path, book=None)` | 导出为制表符分隔文本 |

#### 统计、绘图、保存

| 方法 | 说明 |
| --- | --- |
| `stats(col)` | 返回 `ColumnStats`（count/total/mean/min/max/sd/median） |
| `plot(x_col, y_col, kind="line")` | 绘图。`kind` 见下表 |
| `save(path)` | 保存项目为 `.opju` / `.opj` |
| `export_image(path, page=None, width=None, unit=2, dpi=None)` | 导出图窗口为图片，扩展名决定格式 |

支持的图形类型：`line`、`scatter`、`line_symbol`、`column`、`bar`、`area`。

`plot()` 只画单条 Y 列。**多曲线**要用列区间语法（X 为第 1 列，Y 取第 2~4 列）：

```python
ob.execute('plotxy iy:=(1,2:4) plot:=200 ogl:=[<new>];')
```

导出图片时，`path` 的扩展名决定格式（`png` / `tif` / `jpg` / `pdf` / `eps` 等，
见 `EXPORT_TYPES`）；`width` + `unit`（0 英寸 / 1 厘米 / 2 像素）控制尺寸，
`dpi` 控制栅格分辨率：

```python
ob.export_image(r"D:\out\fig.png", width=1600, unit=2, dpi=300)
```


#### 执行原生 LabTalk

需要用到本库未封装的功能时，可以直接下命令：

```python
ob.execute('newbook name:="X";')          # 执行命令
value = ob.evaluate("total(col(A))")       # 求数值表达式的值
text  = ob.get_text("page.name$")          # 取字符串表达式
```

### 语言切换

```python
from origin_bridge import switch_language, get_language, list_installations

print(list_installations())        # 列出所有安装及语言状态
print(get_language())              # 'C' 或 'E'

switch_language("中文")             # 写入中文设置（需重启生效）
switch_language("en", restart=True) # 切英文并自动重启 Origin
```

> ⚠️ `switch_language()` 默认会**关闭正在运行的 Origin**（`close_first=True`），
> 未保存的工作会丢失。若只想改注册表，传 `close_first=False`。

---

## 命令行工具

```
origin-cli <命令> [参数]

命令：
  info      查看 Origin 安装与运行状态
  lang      查看或切换界面语言
  import    导入数据文件到 Origin
  export    导出当前工作表数据到文件
  stats     查看列统计量
  plot      用两列数据绘图
  save      保存当前项目
```

常用示例：

```bash
origin-cli info                            # 环境自检
origin-cli info --quiet                    # 只查注册表，不连 COM
origin-cli lang                            # 查看当前语言
origin-cli lang zh                         # 切换到中文
origin-cli lang en --restart               # 切换到英文并重启
origin-cli lang --versions                 # 列出所有安装版本
origin-cli import data.csv --book MyData   # 导入 CSV
origin-cli stats A B C                     # 统计多列
origin-cli plot A B --type line_symbol     # 画点线图
origin-cli save result.opju                # 保存项目
```

---

## 界面语言切换

Origin **没有**图形化的语言开关。界面语言由一个注册表字符串值决定：

```
HKEY_CURRENT_USER\Software\OriginLab\<版本键>     →     Language = "C" / "E"
```

| 取值 | 含义 |
| --- | --- |
| `C` | 中文（简体） |
| `E` | 英文 |

**注意「版本键」是内部版本号，不是市场名**，这点最容易搞错：

| 市场名 | 内部版本键 |
| --- | --- |
| Origin 2025b | `Origin 10.2b` |
| Origin 2024b | `Origin 10.1b` |
| Origin 2023b | `Origin 10.0b` |

所以 2025b 对应的键名看起来像旧版本号，**这是正常的**，不要改成 `Origin 2025b`。

### 三种切换方式

**方式一：命令行（推荐）**

```bash
origin-cli lang zh --restart
```

**方式二：双击批处理**

`scripts/` 目录下有两个文件，双击即可（会自动定位版本键）：

- `切换中文.bat`
- `切换英文.bat`

**方式三：Python**

```python
from origin_bridge import switch_language
switch_language("中文", restart=True)
```

三种方式都是修改同一个注册表值，效果完全一致。
`HKCU` 分支**不需要管理员权限**。

> **中文语言包并非 100% 覆盖** —— 部分高级分析工具和第三方 App 的菜单
> 会保持英文，这是正常现象，不影响功能。

---

## 踩坑记录

以下每一条都是在本项目开发过程中**实际踩到并验证过**的。
记在这里是因为它们大多**静默失败、不报错**，排查成本很高。

### 一、连接与识别

| 坑 | 现象 | 正确做法 |
| --- | --- | --- |
| ProgID 用错 | `Origin.ApplicationSI64` 多数安装**未注册**，Dispatch 直接失败 | 用 `Origin.ApplicationSI` |
| 误判调用耗时 | 以为每次调用要 1~3 秒 | 那是**首次建立连接**的开销；之后单次调用在毫秒级 |
| `app.Path` | 报「类型不匹配」 | 它是**方法**不是属性；安装路径改从注册表读 |
| 注册表布局 | 在版本键下找不到 `Executable Path` | 它在 **`<版本键>\International`** 子键下；而 `Language` 在版本键本身 |

### 二、读取数据（重灾区）

| 坑 | 现象 | 正确做法 |
| --- | --- | --- |
| `LTVar()` | 恒返回 `0.0`，**不报错** | 改用 `Evaluate("col(B)[1]")` |
| `LTStr("$(...)")` | 返回空字符串 | 只用于字符串表达式，如 `page.name$` |
| 缺失值 | 返回 `-1.23456789e-300` 这个魔数 | 识别并转成 `None`，否则会当真实数据参与计算 |
| `sum()` | 不是求和函数（实测结果很怪） | 求和用 `total()` |

### 三、统计函数（最隐蔽的坑）

**标准差函数是 `stddev()`，不是 `sd()`。**

```python
app.Evaluate("sd(col(A))")        # -1.23456789e-300  ← 哨兵值，函数不存在
app.Evaluate("std(col(A))")       # -1.23456789e-300  ← 同上
app.Evaluate("stdev(col(A))")     # -1.23456789e-300  ← 同上
app.Evaluate("stddev(col(A))")    # 2.138089935299395 ← 正确！
```

`sd()` / `std()` / `stdev()` / `SD()` / `sdev()` 这些"看起来完全合理"的名字
在 LabTalk 中**都不存在**，调用时**不抛异常**，只返回缺失值哨兵。
如果你只看输出数字，很容易把 `-1.234e-300` 当成某个极端小的结果。

已实测可用的统计函数：`total()`、`mean()`、`min()`、`max()`、`count()`、
`stddev()`、`median()`。

### 四、写入数据

| 坑 | 现象 | 正确做法 |
| --- | --- | --- |
| `PutWorksheet` | 写入位置错乱，写出垃圾值 | 用 `Execute` 拼接 `col(X)[i]=v;` 批量赋值 |
| 逐条 `Execute` 太慢 | 每个值一次 COM 调用 | 拼成一条命令，实测 100 个值约 1 秒 |
| 工作簿名被截断 | 传 `"SignalAnalysis"` 实际得到 `"SignalAnalysi"` | 名称长度上限**约 13 字符**，超出会**静默截断**；务必用返回值而非传入值 |

### 五、保存文件

| 坑 | 现象 | 正确做法 |
| --- | --- | --- |
| LabTalk `save -i` | 执行**不报错**，但**不生成文件** | 用 COM 的 `app.Save(绝对路径)` |
| LabTalk `expASC` | 本次环境中未能生成文件 | 改用 `GetWorksheet` 读数据 + Python 写文件 |
| LabTalk `expGraph` | `type:=1` 实际导出 **BMP** 而非 PNG；`path:=` 若传完整文件路径，则**既不报错也不生成文件** | `type` 要传**字符串**（`png` / `tif` / `jpg`）；`path` 只传**目录**，文件名用 `filename` 单独给 |

### 六、批量读取

| 方式 | 实测 |
| --- | --- |
| `GetWorksheet(name)` | ✅ **一次拿回整表**，返回 row-major 元组 `((A1,B1), (A2,B2), ...)` |
| `Evaluate` 逐行 | 可用，但行数多时明显更慢 |
| LabTalk 循环拼字符串 | ❌ 字符串拼接不生效，返回的是表达式原文 |

`GetWorksheet(name, offset)` 的第二个参数是**起始行偏移**（0-based）。

### 七、Windows 环境相关

| 坑 | 说明 |
| --- | --- |
| `.bat` 中文乱码 | 含中文的批处理必须存为 **GBK** 编码 + CRLF 换行；UTF-8 在 cmd 下会乱码 |
| GUI 进程被回收 | 从受管会话直接 `Popen` 启动的 GUI 程序会被连坐杀掉；改用 `explorer.exe` 启动 |
| 沙箱/安全软件拦截 | `reg.exe` 可能被安全策略禁用；本项目改用 `winreg` 直接操作注册表，更可靠 |

---

## 示例

`examples/` 目录下的脚本都可以直接运行：

| 文件 | 内容 |
| --- | --- |
| `01_quickstart.py` | 连接 → 建表 → 写数据 → 读回 → 统计 → 保存 |
| `02_import_csv.py` | 生成示例 CSV → 导入 Origin → 统计分析 → 导出 |
| `03_plot.py` | 绘制多种图形并导出图片 |
| `04_language.py` | 查看 / 切换界面语言（默认只读模式） |
| `05_batch_export.py` | 批量导出所有工作表数据并生成汇总报告 |

运行测试：

```bash
python tests/test_smoke.py     # 或 pytest tests/ -v
```

测试分两组：纯函数测试（任何环境可跑）和集成测试（需要真实 Origin）。
缺环境时会自动跳过，不会报错。

---

## 常见问题

**Q：提示「无法连接 Origin」怎么办？**

先确认 Origin 已安装且至少手动运行过一次，然后：

```bash
origin-cli info --quiet     # 检查注册表里有没有安装记录
```

如果注册表读得到但依然连不上，试试手动启动 Origin 后重试。

**Q：第一次调用为什么很慢？**

首次 `Dispatch` 需要启动进程并加载类型库，约 3~7 秒，属正常。
之后单次调用在毫秒级 —— 不要在循环里反复建立连接。

**Q：切换语言后没生效？**

必须**完全重启 Origin**。如果 Origin 还开着，读到的仍是旧设置。
用 `--restart` 参数可以自动完成重启。

**Q：中文界面下有些菜单还是英文？**

正常现象。Origin 的中文语言包不是 100% 覆盖，部分高级分析工具与
第三方 App 会保持英文。

**Q：能同时操作多个 Origin 实例吗？**

COM 的 `Dispatch` 会附着到已有的实例或新建一个，无法精确指定目标实例。
如需操作多个，请考虑用不同用户会话或虚拟机隔离。

**Q：脚本卡住不动了？**

`Execute` 在 Origin 弹出模态对话框时会一直阻塞。检查 Origin 窗口是否
有等待你点击的弹窗。自动化脚本建议后台运行并落盘日志，避免超时被杀。

**Q：支持 macOS / Linux 吗？**

不支持。COM 是 Windows 专有技术。
macOS 下可以走 AppleScript，但接口与本文完全不同，需要另写一套。

---

## 项目结构

```
origin-bridge/
├── origin_bridge/
│   ├── __init__.py        包入口，导出公开 API
│   ├── bridge.py          COM 桥接核心
│   ├── language.py        界面语言切换
│   ├── cli.py             命令行工具
│   └── __main__.py        支持 python -m origin_bridge
├── examples/              可直接运行的示例
├── scripts/               语言切换批处理（双击即用）
├── tests/                 冒烟测试 / 集成测试
├── docs/                  补充文档
├── pyproject.toml
├── LICENSE
└── README.md
```

---

## 贡献

欢迎提 Issue 反馈你遇到的坑，特别是：

- 在其他 Origin 版本上的兼容性情况
- 本 README「踩坑记录」里没提到但你也踩到的坑
- 中文语言包的覆盖情况

如果某个 LabTalk 命令在本项目里用不了但你有更好的写法，也非常欢迎 PR。

---

## 许可

[MIT](LICENSE)

本项目与 OriginLab 公司无隶属关系。Origin 是 OriginLab Corporation 的注册商标。

---

<h2 id="english">English</h2>

**origin-bridge** — Automate OriginLab Origin from Python via COM.

Read/write worksheet data, compute descriptive statistics, create plots,
save projects, and switch the UI language (Chinese / English) — all from
clean Python APIs or a CLI.

```python
from origin_bridge import OriginBridge

with OriginBridge() as ob:
    ob.new_workbook("MyData")
    ob.set_column("A", [1, 2, 3, 4, 5])
    print(ob.stats("A"))
    ob.save(r"D:\data\demo.opju")
```

**Why not the official `originpro` package?** In our testing it connects but
every value getter returns an empty string (silently). This project uses
`win32com` directly, with every API verified against a real Origin instance.

**Requirements:** Windows, Python 3.8+, `pywin32`, Origin 2018–2026.

**Install:** `pip install -e .` then run `origin-cli info`.

See the [pitfalls section](#踩坑记录) above — it documents a number of silent
failures (wrong ProgID, `LTVar` always returning `0.0`, `sd()` not existing in
LabTalk, NANUM sentinel values, LabTalk `save -i` not writing files) that are
worth knowing regardless of which library you use.
