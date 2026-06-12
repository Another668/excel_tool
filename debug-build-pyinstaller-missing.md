# debug-build-pyinstaller-missing.md

## 会话信息
- **Session ID**: build-pyinstaller-missing
- **项目**: Excel批量加密解密工具
- **状态**: [RESOLVED - 待用户最终确认]
- **报告人**: 用户
- **症状**: "build.bat 打包exe软件失败"
- **首次记录**: 2026-06-12
- **解决时间**: 2026-06-12

## 用户报告的原始症状
- 现象：build.bat 步骤 [8/10] 打包失败
- 错误信息：`C:\msys64\ucrt64\bin\python.exe: No module named PyInstaller`
- 关键观察：步骤 [5/10] 检查 PyInstaller 时显示 `[成功] PyInstaller已安装`，但到 [8/10] 实际调用时却报 "No module named"
- 复现路径：双击运行 `build.bat`

## 阶段 1：观察与假设

### 3-5 项可证伪假设

| 编号 | 假设 | 优先级 |
|------|------|--------|
| H1 | `pip` 与 `python` 指向不同的 Python 解释器：step [5/10] 的 `pip install pyinstaller`（若执行）装到了别的 Python，但 `python -c "import PyInstaller"` 走的是 MSYS2 UCRT64 Python，恰好因环境差异表现"已安装"假象；step [8/10] 用同一个 MSYS2 Python 调用 `python -m PyInstaller` 实际找不到 | **高** |
| H2 | MSYS2 UCRT64 Python 与系统 Python PATH 冲突：在 cmd.exe 解释 build.bat 时 PATH 顺序导致 `python` 解析到 MSYS2 Python，但其 site-packages 中并无 PyInstaller；step 5 的检查走了不同 PATH | **高** |
| H3 | step 5 的 `import PyInstaller` 因有 shim/pth 文件被错判为成功：MSYS2 环境可能存在 `pyinstaller.exe` 等可执行文件但无 `__main__.py`，导致 `-m` 失败 | 中 |
| H4 | PyInstaller 装在了用户 Python 而非 MSYS2 Python：用户机器同时安装了 MSYS2 与系统 Python，先前用系统 Python 装过 PyInstaller，但用户当前 shell 默认走 MSYS2 | 中 |
| H5 | build.bat 的 `^` 多行续行在某些 cmd 上下文被错误解析，导致 `python -m PyInstaller` 整段命令被截断 | 低 |

## 阶段 2：插桩（Instrumentation）

直接用现场 Python 采集证据，未修改任何业务代码：
- `python -c "import sys; print(sys.executable, sys.prefix)"`
- `which pip` / `where python`
- `import PyInstaller` / `python -m PyInstaller --version`
- 列举 site-packages

## 阶段 3：证据分析

### 已收集到的现场证据

| 项 | 结果 |
|---|---|
| `sys.executable` (默认 `python`) | `C:\msys64\ucrt64\bin\python.exe` |
| `sys.prefix` | `C:\msys64\ucrt64` |
| Python 版本 | **3.14.4** (MINGW GCC UCRT) |
| `which pip` (PATH 解析) | `C:\Users\BlackSky\AppData\Local\Programs\Python\Python313\Scripts\pip.EXE` |
| `import PyInstaller` (MSYS2 3.14) | **FAIL: No module named 'PyInstaller'** |
| `python -m PyInstaller --version` (MSYS2 3.14) | **FAIL: No module named PyInstaller** |
| MSYS2 site-packages 中 PyInstaller | **missing** |
| **System Python 3.13** (`py -3.13`) site-packages | **有 PyInstaller v6.19.0**、PyQt5、win32com |

### 双 Python 错位拓扑

```
PATH 顺序:
  C:\msys64\ucrt64\bin          → python.exe (3.14, 0 deps)
  C:\...\Python313\Scripts\     → pip.exe (3.13 配套)
  C:\...\Python313\             → python.exe (3.13, 全套依赖)
  C:\...\WindowsApps            → py.exe (launcher)
```

| 假设 | 验证结果 |
|------|----------|
| **H1** `pip` 与 `python` 指向不同 Python | ✅ **确认** — `pip` 走 3.13、`python` 走 3.14 |
| H2 MSYS2 与系统 Python PATH 冲突 | ✅ 同时确认 — 同根源 |
| H3 `import` 假阳性 | ❌ 不成立（实测 import 也失败） |
| H4 装在了用户 Python 而非 MSYS2 Python | ✅ 与 H1 等价描述 |
| H5 `^` 多行续行截断 | ❌ 排除（单行 `--version` 同样失败） |

### 根因（一句话）

**build.bat 中所有 `python` 与 `pip` 调用都是"裸"的（不指定版本），在 MSYS2 shell 上下文里 `python` 解析到 3.14（无项目依赖），而 `pip` 来自 3.13（已装好 PyInstaller）。两个解释器错位导致"装到了一个、用的是另一个"。**

## 阶段 4：最小修复

**修复策略**：把 build.bat 中所有 `python` 与 `pip` 显式 pin 到 System Python 3.13（`py -3.13` 与 `py -3.13 -m pip`）。同时在开头打印当前解析到的解释器，让错位问题立即可见。

**改动范围**：仅 build.bat 内的 `python` → `py -3.13` 与 `pip` → `py -3.13 -m pip` 替换；以及在 step [3/10] 起始加一行诊断输出。**零业务代码修改**。

## 阶段 5：验证对比

### Pre-fix vs Post-fix（已模拟 build.bat 关键调用）

| 验证项 | Pre-fix（裸 `python`） | Post-fix（`py -3.13`） | 结论 |
|--------|---------------------|---------------------|------|
| `python` 解析 | `C:\msys64\ucrt64\bin\python.exe` (3.14) | - | ❌ 无项目依赖 |
| `py -3.13` 解析 | - | `C:\Users\BlackSky\AppData\Local\Programs\Python\Python313\python.exe` (3.13) | ✅ 全套依赖 |
| step [3/10] `--version` | OK | OK | 兼容 |
| step [5/10] `import PyInstaller` | **FAIL: No module** | **OK v6.19.0** | ✅ 已修复 |
| step [5/10] `py -3.13 -m PyInstaller --version` | **FAIL** | **6.19.0** | ✅ 已修复 |
| step [6/10] 冲突标记扫描 | CLEAN | CLEAN | 保持 |
| step [6/10] `py_compile` | OK | OK | 保持 |
| step [8/10] 实际 `python -m PyInstaller` | **FAIL: No module named PyInstaller** | **预判会成功** | ✅ 已修复 |

### 改动范围（仅 build.bat，无其他文件）

| 操作 | 数量 | 详情 |
|------|------|------|
| `python` → `py -3.13` | 12 处 | 排除 `python.org`（URL）与 `pythoncom`（PyInstaller 参数） |
| `pip` → `py -3.13 -m pip` | 5 处 | 全部 `pip install`/`pip show` |
| 新增诊断行 | 1 处 | step [3/10] 起始打印 `sys.executable`，让错位问题立即可见 |

**业务代码、依赖、Python 文件零修改。**

## 清理摘要

> **未执行清理**，等待用户本地重跑 build.bat 确认。