# 故障排查指南

本文档汇总 `origin-bridge` 使用中的常见问题与诊断方法。
README 的「踩坑记录」讲的是**原理**，这里讲**怎么查**。

---

## 快速诊断

先跑这三步，能定位绝大多数问题：

```bash
# 1. 只查注册表，不连 COM —— 验证 Origin 是否装好
origin-cli info --quiet

# 2. 完整自检，含 COM 连接 —— 验证能否自动化
origin-cli info

# 3. 跑测试 —— 纯函数测试不需要 Origin，集成测试需要
python tests/test_smoke.py
```

第 1 步输出示例：

```
============================================================
  Origin 环境信息
============================================================
  注册表中共有 1 个安装记录：
    · Origin 10.2b: 中文（简体）  build=212  D:\Program Files\OriginLab\Origin2025b

  ✓ Origin 正在运行（PID: 16984）

  当前语言：Origin 10.2b  →  中文（简体）
```

---

## 一、连不上 Origin

### 症状

```
OriginConnectionError: 无法连接 Origin（ProgID=Origin.ApplicationSI）
```

### 排查顺序

**1. 注册表里有没有安装记录？**

```bash
origin-cli info --quiet
```

- **没有任何记录** → Origin 未安装，或安装后从未运行过。
  Origin 是在**首次运行时**才写入注册表的，请手动启动一次。

- **有记录但路径为空** → 安装已损坏或被移动，建议重装。

**2. ProgID 是否注册？**

在「注册表编辑器」中确认存在：

```
HKEY_CLASSES_ROOT\Origin.ApplicationSI
```

如果只有 `Origin.Application` 而没有 `Origin.ApplicationSI`，
说明安装不完整。

> ⚠️ **不要用 `Origin.ApplicationSI64`** —— 它在绝大多数安装中**并未注册**，
> 这是个常见的误导性猜测。

**3. Origin 是否在运行？**

```bash
tasklist | findstr Origin64
```

若未运行，先手动启动一次再试。首次 `Dispatch` 需要 3~7 秒，请耐心等待。

---

## 二、读到的数据不对

### 症状 A：数值全是 `0.0`

**原因**：用了 `LTVar()`。这个方法**恒返回 0.0**，而且不报错。

```python
# ❌ 错误
ob.app.LTVar("col(B)[1]")        # 永远是 0.0

# ✅ 正确
ob.evaluate("col(B)[1]")         # 真实值
```

### 症状 B：数值全是空字符串

**原因**：用了 `LTStr("$(...)")`。`LTStr` **只适用于字符串表达式**。

```python
# ❌ 错误
ob.app.LTStr("$(col(B)[1])")     # 空字符串

# ✅ 字符串用它
ob.get_text("page.name$")        # "Book1"

# ✅ 数值用 evaluate
ob.evaluate("col(B)[1]")
```

### 症状 C：出现 `-1.23456789e-300` 这种数

**这是 Origin 的缺失值哨兵（NANUM）**，不是真实数据。

来源通常有两种：

1. **单元格本身是空的** — 正常，`origin-bridge` 会自动转成 `None`
2. **调用了不存在的函数** — 见下一条

### 症状 D：转成 `None` 了

`origin-bridge` 会把 NANUM 哨兵统一转成 Python 的 `None`。
如果你的数据里真的出现了 `None`：

- 检查原始单元格是否为空
- 检查是不是调用了不存在的 LabTalk 函数（会返回哨兵 → 被转成 `None`）

---

## 三、统计量不对（最隐蔽的坑）

### 症状：标准差是个天文数字或极小值

```
sd = -1.23457e-300
```

**原因**：用了 `sd()`。这个函数在 LabTalk 中**不存在**。

```python
# ❌ 这些函数全都不存在，调用不报错，只返回哨兵值
"sd(col(A))"       # -1.23456789e-300
"std(col(A))"      # -1.23456789e-300
"stdev(col(A))"    # -1.23456789e-300
"SD(col(A))"       # -1.23456789e-300

# ✅ 正确的是 stddev
"stddev(col(A))"   # 2.138089935299395
```

`origin-bridge` 的 `stats()` 已使用正确函数，此问题只在你手写
`evaluate()` 表达式时才会遇到。

### 已验证可用的统计函数

| 功能 | 函数 |
| --- | --- |
| 计数 | `count()` |
| 求和 | `total()` ← 不是 `sum()` |
| 均值 | `mean()` |
| 最小值 | `min()` |
| 最大值 | `max()` |
| 样本标准差 | `stddev()` ← **不是 `sd()`** |
| 中位数 | `median()` |

> `sum()` 在 LabTalk 中存在但**不是求和函数**，实测 `sum(col(A))` 返回的结果
> 与求和无关。求和请一律用 `total()`。

---

## 四、写入数据异常

### 症状：数据写到了奇怪的位置，出现垃圾值

**原因**：用了 `PutWorksheet`。

```python
# ❌ 实测写入位置错乱
ob.app.PutWorksheet("Book1", data)

# ✅ 用 Execute 拼接（本库的 set_column 已封装）
ob.set_column("A", [1, 2, 3])
```

### 症状：写入很慢

**原因**：每个值发一次 COM 调用。

```python
# ❌ 慢：100 个值 = 100 次调用
for i, v in enumerate(values):
    ob.execute("col(A)[%d]=%s;" % (i + 1, v))

# ✅ 快：拼成一条命令，100 个值约 1 秒
ob.set_column("A", values)
```

`set_column` 的 `batch` 参数控制每条命令包含多少个赋值（默认 200），
命令过长时可调小。

### 症状：工作簿名称不是我指定的

```
UserWarning: 工作簿名称被 Origin 截断：请求 'SignalAnalysis'，实际 'SignalAnalysi'
```

**原因**：Origin 对工作簿名称有长度上限，实测**约 13 个字符**，超出部分被静默截断。

```python
# ❌ 传入 14 字符，实际只保留 13 个
book = ob.new_workbook("SignalAnalysis")     # 实际得到 "SignalAnalysi"

# ✅ 用返回值，不要用传入值
book = ob.new_workbook("SignalAnalysis")
ob.set_column("A", data)                      # 后续操作基于当前活动页，没问题
ob.get_data(book)                             # 若按名字查，必须用实际名
```

`new_workbook()` 会在检测到截断时发出 `UserWarning`，建议保持警告可见，
或改用简短名称（≤ 13 字符）。

---

## 五、保存失败

### 症状：`save` 调用没报错，但文件不存在

**原因**：用了 LabTalk 的 `save -i`。

```python
# ❌ 执行成功，但不生成文件
ob.execute('save -i "D:\\out.opju";')

# ✅ 用 COM 的 Save 方法（本库已封装）
ob.save(r"D:\out.opju")
```

### 症状：`export_ascii` 没生成文件

LabTalk 的 `expASC` 在部分环境中不工作。
`origin-bridge` 的 `export_ascii()` **不走 expASC**，
而是用 `GetWorksheet` 读数据再由 Python 写文件，行为更可控。

如果是自己写的 expASC 命令失败，改用：

```python
ob.export_ascii(r"D:\out.txt")
```

---

## 六、性能问题

### 症状：第一次调用特别慢

**这是正常的。** 首次 `Dispatch` 需要启动进程 + 加载类型库，约 3~7 秒。

关键点：**之后的调用是毫秒级的**。如果你觉得每次调用都慢，
多半是重复建立了连接。

```python
# ❌ 慢：每次循环都连接一次
for name in names:
    with OriginBridge() as ob:
        ob.new_workbook(name)

# ✅ 快：连接一次，复用到结束
with OriginBridge() as ob:
    for name in names:
        ob.new_workbook(name)
```

### 症状：读取大表很慢

用 `get_data()` 一次读回整表，而不是逐单元格读。

```python
# ❌ 慢：1000 行 × 5 列 = 5000 次调用
for r in range(1000):
    for c in range(5):
        ob.evaluate("col(%d)[%d]" % (c + 1, r + 1))

# ✅ 快：一次调用拿回全部
data = ob.get_data()
```

---

## 七、脚本卡死

### 症状：`Execute` 一直不返回

**原因**：Origin 弹出了模态对话框在等你操作，`Execute` 会一直阻塞。

**处理**：

1. 切到 Origin 窗口，看是否有等待点击的弹窗
2. 关掉弹窗后脚本会自动继续

**预防**：自动化脚本建议后台运行并落盘日志，避免被外层超时杀掉：

```python
import subprocess, sys

with open("run.log", "w") as fh:
    subprocess.Popen([sys.executable, "my_script.py"],
                     stdout=fh, stderr=subprocess.STDOUT)
```

---

## 八、语言切换问题

### 症状：改了注册表但界面没变

**原因**：**没有重启 Origin**。语言设置在启动时读取一次，之后不会重新读取。

```python
from origin_bridge import switch_language

switch_language("中文", restart=True)   # 自动关闭并重启
```

### 症状：找不到注册表键

`Language` 值位于：

```
HKEY_CURRENT_USER\Software\OriginLab\<内部版本键>
```

**内部版本键 ≠ 市场名**，这是最容易搞错的地方：

| 市场名 | 内部版本键 |
| --- | --- |
| Origin 2025b | `Origin 10.2b` |
| Origin 2024b | `Origin 10.1b` |
| Origin 2023b | `Origin 10.0b` |

```bash
origin-cli lang --versions     # 列出实际存在的版本键
```

### 症状：部分菜单仍是英文

**正常现象。** Origin 的中文语言包不是 100% 覆盖，
部分高级分析工具和第三方 App 会保持英文，不影响功能。

### 症状：写入注册表报权限错误

`HKCU` 分支**不需要管理员权限**。若仍报错，通常是：

- 安全软件拦截了注册表写入
- 系统组策略限制了注册表访问

可尝试以管理员身份运行，或临时关闭安全软件验证。

---

## 九、环境特定问题

### 在 CI / 受管环境中运行

| 限制 | 表现 | 应对 |
| --- | --- | --- |
| `reg.exe` 被禁用 | 命令被安全策略拦截 | 本项目用 `winreg` 操作注册表，不受影响 |
| GUI 进程被回收 | 启动的 Origin 立刻消失 | 用 `explorer.exe` 启动，让进程归属桌面 shell |
| 无 `tasklist` | 进程检测失败 | `list_processes()` 会返回空列表，不影响其他功能 |

### 批处理中文乱码

`.bat` 含中文时必须存为 **GBK 编码 + CRLF 换行**。
UTF-8 在 cmd 默认代码页（936）下会显示成乱码。

`scripts/` 下的两个批处理已按此要求生成。

### 多个 Origin 版本共存

`Dispatch` 会附着到已有实例或新建一个，**无法指定目标实例**。
若注册表中有多个版本键：

```python
from origin_bridge import set_language

# 显式指定要改哪个版本
set_language("C", version_key="Origin 10.2b")
```

---

## 十、还是没解决？

收集以下信息后提 Issue：

1. `origin-cli info --quiet` 的完整输出
2. `python tests/test_smoke.py` 的完整输出
3. Origin 版本（`帮助 → 关于 Origin` 里的信息）
4. Python 版本与 `pywin32` 版本：

```bash
python -c "import sys, win32; print(sys.version); print(win32.__file__)"
```

5. 报错的完整堆栈
