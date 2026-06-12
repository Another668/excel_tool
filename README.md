# Excel批量加密解密工具

<div align="center">

**一个高效的Windows桌面应用，用于批量加密和解密Excel文件**

[![Version](https://img.shields.io/badge/version-3.6.6-blue.svg)](https://github.com/excel-protector/releases)
[![Python](https://img.shields.io/badge/python-3.6+-green.svg)](https://www.python.org/)
[![PyQt5](https://img.shields.io/badge/PyQt5-5.15+-orange.svg)](https://www.riverbankcomputing.com/software/pyqt/)
[![License](https://img.shields.io/badge/license-MIT-lightgrey.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)](https://www.microsoft.com/windows)

</div>

---

## 📋 版本迭代规范

### 版本号格式
采用语义化版本控制体系，版本号格式为：**X.Y.Z**，其中：
- **X** 表示主版本号 (Major Version)
- **Y** 表示次版本号 (Minor Version)
- **Z** 表示修订号 (Patch Version)

### 版本号变更规则

#### 1. 主版本号 (X)
- **变更条件**：发生以下任意情况时，必须递增主版本号：
  - 系统架构发生重大调整或重构
  - 引入不兼容旧版本的 API 变更
  - 核心功能模块进行颠覆性重构
  - 产品进行重大改版或换代升级
- **变更示例**：1.2.3 → 2.0.0

#### 2. 次版本号 (Y)
- **变更条件**：发生以下任意情况时，必须递增次版本号，同时将修订号重置为 0：
  - 新增重要功能模块
  - 现有功能进行较大幅度增强或扩展
  - 引入新的 API 但保持向后兼容
  - 性能或用户体验有显著提升
- **变更示例**：1.1.5 → 1.2.0

#### 3. 修订号 (Z)
- **变更条件**：发生以下任意情况时，必须递增修订号：
  - 修复已知 Bug 或缺陷
  - 进行代码优化或重构（不影响外部接口）
  - 文档更新或注释完善
  - UI/UX 细节调整
  - 性能微调或小范围改进
- **变更示例**：2.3.4 → 2.3.5

### 版本号变更原则
- 修复 Bug 或小优化时，递增修订号 (Z)
- 新增功能或大模块更新时，递增次版本号 (Y) 并重置修订号为 0
- 架构大改或不兼容更新时，递增主版本号 (X) 并重置次版本号和修订号为 0
- 所有版本号变更必须在本更新日志中详细记录变更内容、影响范围及迁移指南（如适用）

---

## 📝 更新日志

### [v3.6.6] - 2026-06-12

#### 🐛 根本性修复：按钮与文件列表行视觉重叠
- **问题根因**
  - 文件组 `MinimumHeight(280)` + 表格 `MinimumHeight(360)` + 其他固定区域 > 小窗口可用高度
  - 表格 31 行时 `sizeHint ≈ 1024px`，远超文件组容量
  - 表格行被画到按钮行上方，造成视觉重叠
- **根本方案：QScrollArea 接管主内容**
  - 用 `QScrollArea` 包裹整个主布局（`init_ui`）
  - 内部 widget 始终获得自然高度，不再被窗口挤压
  - 窗口高度不足时自动出现垂直滚动条
  - 表格与按钮的相对位置永远正确（按钮在表格下方）
  - 彻底解决布局溢出问题
- **配套优化**
  - 表格 sizePolicy: `Expanding, Preferred`（不强行撑大）
  - 表格 maxHeight 限制放宽到 800px，minHeight 收紧到 150px
  - 文件组 minHeight 收紧到 200px
  - `_on_window_resize` 简化为只设上限
  - 滚动条美化：10px 窄条、圆角 handle、hover 高亮

#### ✅ 多分辨率验证
- 1024×680: 内部 widget 934px，滚动条出现，按钮在表格下方 ✓
- 1280×800: 内部 widget 934px，滚动条出现，按钮在表格下方 ✓
- 1440×900: 内部 widget 934px，滚动条出现，按钮在表格下方 ✓
- 1920×1080: 内部 widget 934px，滚动条消失，全内容可见 ✓

#### 📦 影响范围
- 文件清单：Excel批量解密工具与密码管理.py、README.md
- 不影响打包脚本（build.bat 自动读取新 __version__ 生成 v3.6.6.exe）

### [v3.6.5] - 2026-06-12

#### 🐛 按钮重叠修复
- **文件列表区底部按钮与表格重叠问题**
  - 原因：`QTableWidget` 设置 `sizePolicy(Expanding, Expanding)` 且 `MinimumHeight=380`，当文件数>25 时表格撑出 `QGroupBox` 边界
  - 修复：内部 `QVBoxLayout` 使用 `stretch(表格=1, 按钮=0)`，表格吸收多余空间、按钮保持固定高度
  - 表格按钮 `setFixedHeight(34)` + 最小宽度，按钮行永远可见不再重叠
  - 日志框从 `MaximumHeight=180` 改为 `MinimumHeight=100` + Preferred，按需伸缩
  - 文件组 `MinimumHeight` 从 420 降为 280，由内部 stretch 自然填充

#### 📐 响应式布局系统
- **屏幕档位识别**
  - tiny (<1366宽) / small (1366-1680) / medium (1680-1920) / large (≥1920)
  - 基于 `QScreen.availableGeometry()` + `devicePixelRatio()` 检测
  - 紧凑屏用更小间距(8px) 与边距(10px)，平衡屏用标准间距(12px)
- **智能窗口尺寸**
  - 默认窗口：`min(屏幕宽×0.85, 1400)` × `min(屏幕高×0.88, 1000)`
  - 最小窗口：1024×680，避免小屏内容溢出
  - 窗口居中显示
  - `setMinimumSize()` 防止用户拖到过小
- **动态控件尺寸**
  - 表格行高：tiny=26 / small=28 / medium=32 / large=32 px
  - 按钮高度：tiny=30 / small=32 / medium=34 / large=36 px
  - 字号：tiny/small=12pt，medium/large=13pt
  - `_apply_responsive_styles()` 在 `apply_styles()` 末尾自动调用
- **窗口大小变化自适应**
  - 覆写 `resizeEvent` → `_on_window_resize()`
  - 窗口拉大时文件表格高度按 `win_h // 2` 自动扩展
  - 多文件时滚动条接管，按钮行始终在表格正下方
  - 不会因内容增多导致按钮被遮挡

#### 📦 影响范围
- 文件清单：Excel批量解密工具与密码管理.py、README.md
- 不影响打包脚本（build.bat 自动读取新 __version__ 生成 v3.6.5.exe）

### [v3.6.4] - 2026-06-12

#### 🎨 主界面布局重构
- **功能设置 + 密码设置并排布局**
  - 两个模块由"上下堆叠"改为"左右并排"，处于同一水平行
  - 左侧"功能设置"压缩宽度（最大 420px），保持紧凑
  - 右侧"密码设置"占主要宽度（最小 500px，stretch=1 自动扩展）
  - 两模块在视觉上平级、风格统一，QGroupBox 边框/标题样式一致
- **"统一密码"标签加粗**
  - 字号加大到 14px，font-weight: bold
  - 颜色 #1976D2（主蓝），与按钮色调呼应
  - 强化视觉强调，便于用户快速定位核心操作
- **文件列表区域扩展**
  - 占据原密码设置模块的垂直空间（absorb 高度约 420px）
  - 默认即可看到 10-12 行文件，文件夹扫描结果一目了然
  - 表格 sizePolicy.Expanding + 主布局 stretch=3，窗口拉大时优先扩展文件列表
- **滚动性能优化**
  - QTableWidget 滚动模式改为 ScrollPerPixel，平滑且开销低
  - setWordWrap(False) 减少逐行重绘
  - 主布局 spacing 从 15px 收紧到 12px，节省空间
- **窗口高度自适应**
  - 初始窗口高度 850 → 900px，容纳更长的文件列表
  - 文件列表区最小高度 420px、表格最小高度 380px

#### 📦 影响范围
- 文件清单：Excel批量解密工具与密码管理.py、README.md
- 不影响打包脚本（build.bat 自动读取新 __version__ 生成 v3.6.4.exe）

### [v3.6.3] - 2026-06-12

#### ✨ 打包脚本优化
- **EXE 文件名自动附带版本号**
  - 每次打包自动从 `__version__` 提取版本号，构造 `Excel批量加密解密工具_v3.6.3.exe` 命名
  - 解决多版本并存时文件名冲突、混淆等问题
  - 验证/UPX 步骤同步使用带版本号的文件名
  - build.bat 中通过 `for /f` 解析 Python 源文件的 `__version__`，无需外部参数
- **构建流程稳健性**
  - 若版本号解析失败自动 fallback 为 `unknown`，不会中断打包

#### 🎨 GUI 全面优化
- **高 DPI 适配**
  - 启动前启用 `Qt.AA_EnableHighDpiScaling` 与 `AA_UseHighDpiPixmaps`
  - 设置 `HighDpiScaleFactorRoundingPolicy.PassThrough` 减少缩放模糊
  - 解决高分辨率屏下字体过小、控件错位问题
- **字体后备链**
  - 应用全局默认字体设为 `Microsoft YaHei UI` (9pt)
  - 通过 `QFont.setFamilies` 设置后备链：Microsoft YaHei UI → Microsoft YaHei → Segoe UI
  - 样式表同步加 `font-family` 链，含 `微软雅黑` 兜底
  - 解决部分中文字符显示为方框、字形不完整问题
- **控件尺寸自适应**
  - 全局 `font-size: 13px`（原 12px），中英文混排更清晰
  - QLabel `min-height: 20px`，给中文字符留足垂直空间
  - 表格数据行高 32px、表头行高 32px（防中文截断）
  - QPushButton/QLineEdit/QRadioButton 等均设 `min-height`，避免压缩
- **菜单/状态栏同步加字号**
  - QMenuBar、QMenu 字号 13px，菜单层级与正文一致

#### 🖼️ 图标显示修复
- **任务栏图标修复**
  - `SetCurrentProcessExplicitAppUserModelID` 从 `_load_and_set_icon`（窗口创建后）前移到 `main()`（窗口创建前）
  - 解决 Windows 7+ 任务栏把 EXE 归为"通用"图标的问题
  - AppUserModelID 包含版本号（`com.excelprotector.exceltool.3.6.3`），多版本并存也能正确分组
- **窗口/应用图标双绑**
  - QApplication 与 QMainWindow 各自调用 `setWindowIcon`
  - 所有 QDialog（关于/使用说明/完成提示等）继承应用图标
- **图标资源规范化**
  - 图标.ico 已含 16/32/48/64/128/256 六种尺寸，跨场景显示清晰
  - PyInstaller 通过 `--icon=图标.ico` 嵌入 EXE 资源，文件系统图标正常显示

#### 📦 影响范围
- 文件清单：build.bat、Excel批量解密工具与密码管理.py、README.md
- 兼容性：保持与之前版本的配置文件、密码本格式完全兼容
- 用户感知：升级后首次启动即可看到字体/图标/任务栏全方位改善

---

### [v3.6.2] - 2026-06-12

#### 🐛 打包脚本修复
- **修复 pip/python 解释器错位问题**
  - build.bat 中所有裸 `python` 与 `pip` 调用统一 pin 到 `py -3.13`（12 处）与 `py -3.13 -m pip`（5 处）
  - 解决 `C:\msys64\ucrt64\bin\python.exe: No module named PyInstaller` 报错——根因为 MSYS2 Python 3.14（无项目依赖）与 System Python 3.13（已装好 PyInstaller）共存导致 `pip` 与 `python` 解析到不同解释器
  - 排除 `python.org`（URL）与 `pythoncom`（PyInstaller 参数）误改
- **新增启动诊断行**
  - step [3/10] 起始打印 `sys.executable`，让解释器错位问题立即可见
  - 日志示例：`[信息] 打包将使用: C:\Users\BlackSky\AppData\Local\Programs\Python\Python313\python.exe`
#### 📦 影响范围
- 打包流程：仅在 `py -3.13` 解析失败时才会报错（之前会静默选错解释器）
- 用户机器要求：必须安装 Python 3.13（已有则无需变动）

---

### [v3.6.1] - 2026-06-12

#### 🐛 关键Bug修复
- **修复未解决的Git合并冲突**
  - 解决了主源文件、convert_icon.py、requirements.txt 中残留的Git合并冲突标记
  - 恢复了Python解释器对所有源文件的正常解析（`ast.parse` 与 `py_compile` 均通过）
  - 解决"批量解密模块无法使用"的根因——冲突标记导致 `SyntaxError`，程序根本无法启动
  - 保留 HEAD 分支（v3.6.0），剔除 v3.0 旧版分支
- **build.bat 适配修复后版本**
  - 版本号 v3.5 → v3.6.0
  - 升级为 10 步构建流程
  - 新增步骤 [6/10]：源代码健全性检查（Git 冲突标记扫描 + Python 语法校验）
  - 防止类似冲突再次进入打包流程
- **依赖清单修复**
  - requirements.txt 恢复为版本化固定（`PyQt5==5.15.11`、`pywin32==308`）

#### 📦 影响范围
- 主程序入口：原 `SyntaxError` 已消除，所有模块（含批量解密）恢复正常
- 打包流程：新增健全性检查环节，冲突未解决时拒绝打包
- 文件瘦身：主源文件从 2997 行缩减至 1507 行（剔除 1490 行旧版分支）

---

### [v3.6.0] - 2026-05-15

#### 🐛 打包脚本修复
- **修复Git合并冲突**
  - 解决了build.bat文件中的Git合并冲突标记
  - 恢复了完整的9步打包流程
  - 确保脚本正确读取和处理文件路径
- **优化打包验证逻辑**
  - 修复了EXE文件验证时显示的错误信息
  - 确保打包成功时正确显示状态信息
  - 验证文件存在性后立即更新显示
- **版本迭代规范引入**
  - 引入语义化版本控制体系（X.Y.Z）
  - 添加完整的版本迭代规范文档
  - 更新版本号从3.6升级为3.6.0

---

### [v3.5.0] - 2026-05-15

#### 🧪 版本同步验证
- **版本同步机制验证**
  - 验证了从README.md到Python程序的版本同步功能
  - 确保打包流程中版本同步正确执行
  - 完成了两次不同版本的打包测试

---

### [v3.4.0] - 2026-05-15

#### 🔧 版本同步机制
- **新增版本同步工具（sync_version.py）**
  - 自动从README.md读取最新版本号
  - 智能匹配：优先查找版本标题 `### [v3.4]`，备用方案查找徽章
  - 自动更新Python主程序中的 `__version__` 变量
  - 自动更新文档注释中的版本信息
  - 内置版本验证功能，确保同步成功
- **build.bat打包流程升级**
  - 升级到9步构建流程
  - 新增步骤[6/9]：版本同步
  - 自动调用sync_version.py进行版本同步
  - 错误容错：同步失败仍可继续打包
- **README文档结构优化**
  - 更新日志移至文档顶部，方便用户第一时间查看
  - 保持完整的项目文档结构
  - 更新徽章显示当前版本

#### 📦 项目文件管理
- **新增项目文件**
  - sync_version.py：版本同步专用工具
  - 支持独立运行，也可集成到打包流程
- **优化现有文件**
  - build.bat：集成版本同步功能
  - README.md：结构重排，版本信息更醒目

---

### [v3.3.0] - 2026-05-15

#### 🚀 性能优化
- **启动速度大幅提升**
  - 延迟加载win32com和pythoncom模块
  - 仅在实际需要Excel处理时才加载COM组件
  - 减少初始内存占用，提升启动响应速度
- **代码精简优化**
  - 移除未使用的导入（traceback、pythoncom）
  - 清理冗余属性（self.excel_app、self._icon）
  - 删除注释掉的调试代码
  - 优化异常处理逻辑

#### 🔧 构建系统升级
- **打包脚本优化（build.bat v3.3）**
  - 升级到8步构建流程
  - 添加依赖排除机制：--exclude-module
    - 排除tkinter、matplotlib、pandas、numpy、pytest
    - 显著减小EXE文件体积
  - 添加精确的PyQt5模块导入声明
  - 改进文件大小显示（同时显示MB和KB）
- **UPX压缩集成**
  - 自动检测D:`tools`upx.exe
  - 使用--best --lzma最佳压缩模式
  - 显示压缩前后对比和节省空间百分比
  - 优雅降级：无UPX时仍可正常打包

#### 💻 代码质量改进
- **资源路径优化**
  - 改进get_resource_path()函数
  - 使用try-except替代hasattr检测
  - 更好地兼容开发环境和打包环境
  - 使用__file__获取脚本目录
- **异常处理优化**
  - 移除冗余的traceback导入
  - 简化异常捕获逻辑
  - 保持错误信息的完整性
- **版本管理统一**
  - 更新版本号至v3.3
  - 统一代码和文档中的版本信息

#### 📦 依赖管理
- **排除不必要模块**
  - tkinter：不使用的GUI库
  - matplotlib：不使用的绘图库
  - pandas/numpy：不使用的数据处理库
  - pytest：仅开发阶段使用的测试库
- **优化打包体积**
  - 预期可减少30-50%的文件大小
  - 保持所有功能完整性
  - 提升启动和运行速度

---

## 📖 项目简介

Excel批量加密解密工具是一款专为Windows用户设计的桌面应用程序，提供安全、高效的Excel文件批量处理能力。通过友好的图形界面，您可以轻松为多个Excel文件设置密码保护，或移除现有密码保护，而无需手动逐个操作。

### ✨ 核心特性

- 🔒 **批量加密** - 为多个Excel文件同时设置密码保护
- 🔓 **批量解密** - 批量移除Excel文件的密码保护
- 📚 **密码本管理** - 支持CSV格式密码本，自动匹配文件名和密码
- 🎯 **智能处理** - 完全后台运行，不显示任何Excel窗口或对话框
- 🛡️ **文件安全** - 原文件不受影响，处理后生成新文件
- 📊 **实时反馈** - 进度条、彩色日志、详细的状态追踪
- 🎨 **现代界面** - Material Design风格，操作直观友好

---

## 🚀 快速开始

### 方式一：直接使用EXE文件（推荐）

1. 从 [Releases](https://github.com/excel-protector/releases) 下载最新版本
2. 解压后双击 `Excel批量加密解密工具.exe` 即可运行
3. 开始使用！（需要系统已安装 Microsoft Excel）

### 方式二：从源代码运行

**环境要求**
- Python 3.6+
- Windows 操作系统
- Microsoft Excel 已安装

**安装步骤**

```bash
# 1. 克隆或下载项目
git clone <repository-url>
cd Excel_tool

# 2. 安装依赖
pip install -r requirements.txt

# 3. 运行程序
python "Excel批量解密工具与密码管理.py"
```

---

## 📖 使用说明

### 基本流程

```
选择文件夹 → 扫描文件 → 设置密码 → 预览 → 开始处理 → 查看结果
```

**详细步骤**：

1. **选择输入文件夹** - 包含要处理的Excel文件
2. **选择输出文件夹** - 处理后的文件保存位置
3. **选择模式** - 批量加密 或 批量解密
4. **设置密码** - 统一密码 或 从密码本加载
5. **点击"开始执行"** - 确认后即可开始处理

### 密码本格式

CSV文件格式，示例：

```csv
文件名,密码,备注
文件1.xlsx,password123,示例1
文件2.xlsx,abc@2024,示例2
财务表.xlsx,YS2026-XM,重要文件
# 以#开头的行会被忽略
```

**支持**：
- UTF-8、GBK等多种编码
- 逗号或制表符分隔
- 注释行（以#开头）

### 快捷键

| 快捷键 | 功能 |
|--------|------|
| `Ctrl+I` | 打开输入文件夹 |
| `Ctrl+O` | 打开输出文件夹 |
| `Ctrl+E` | 导出配置 |
| `Ctrl+Q` | 退出程序 |
| `Ctrl+A` | 全选文件 |
| `Ctrl+D` | 取消全选 |
| `Ctrl+L` | 清空日志 |
| `F1` | 查看使用说明 |

---

## 📸 功能特性

### 🔐 加密功能
- ✅ 为Excel文件设置读写密码
- ✅ 支持统一密码和独立密码
- ✅ 自定义文件名后缀
- ✅ 保持原文件不变

### 🔓 解密功能
- ✅ 移除Excel文件的密码保护
- ✅ 支持从密码本自动匹配密码
- ✅ 支持无密码文件解密
- ✅ 处理完成后文件状态标记

### 📋 密码管理
- ✅ CSV密码本导入导出
- ✅ 自动编码检测
- ✅ 批量更新密码
- ✅ 密码显示/隐藏切换

### 📊 界面功能
- ✅ 实时进度条显示
- ✅ 彩色日志（成功/失败/警告/信息）
- ✅ 文件状态追踪
- ✅ 日志导出功能
- ✅ 配置自动保存

---

## 🔧 打包说明

### 一键打包

双击运行 `build.bat` 脚本，自动完成以下流程：

1. 清理旧文件
2. 转换图标为ICO格式
3. 检查运行环境
4. 检查项目依赖
5. 检查打包工具
6. 读取并同步版本号
7. 开始打包为EXE文件
8. 验证打包结果和版本一致性
9. UPX压缩（可选）

### 手动打包

```bash
# 安装PyInstaller
pip install pyinstaller

# 执行打包
pyinstaller --onefile --windowed --name "Excel批量加密解密工具" --icon=图标.ico --add-data "requirements.txt;." --exclude-module tkinter --exclude-module matplotlib --exclude-module pandas --exclude-module numpy --exclude-module pytest --hidden-import=win32com --hidden-import=win32com.client --hidden-import=pythoncom --hidden-import=PyQt5 --hidden-import=PyQt5.QtCore --hidden-import=PyQt5.QtGui --hidden-import=PyQt5.QtWidgets --noconfirm "Excel批量解密工具与密码管理.py"
```

生成的EXE文件位于 `dist/` 目录下。

---

## 📁 项目结构

```
Excel_tool/
├── Excel批量解密工具与密码管理.py    # 主程序代码
├── build.bat                        # 一键打包脚本
├── convert_icon.py                  # 图标转换脚本
├── sync_version.py                  # 版本同步工具
├── requirements.txt                 # Python依赖清单
├── README.md                        # 项目文档
├── tool_config.json                 # 配置文件（自动生成）
├── 图标.png                         # 应用程序图标
├── 图标.ico                         # ICO格式图标
└── dist/
    └── Excel批量加密解密工具.exe      # 打包生成的可执行文件
```

---

## 🤝 贡献指南

欢迎贡献代码、报告问题或提出建议！

### 提交问题
- 使用 GitHub Issues 报告Bug或提出功能请求
- 提供详细的复现步骤和环境信息

### 提交代码
1. Fork 本仓库
2. 创建特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 提交 Pull Request

---

## 📦 依赖项

| 依赖 | 版本 | 用途 |
|------|------|------|
| PyQt5 | 5.15+ | GUI界面框架 |
| pywin32 | 308+ | Windows COM组件调用 |

详见 [requirements.txt](requirements.txt)

---

## ⚠️ 注意事项

- 原文件不会被修改，所有处理都在新文件中进行
- 加密时必须设置密码
- 解密时需要提供正确的密码
- 目标计算机需要安装 Microsoft Excel
- 不支持加密后已被修改过的文件

---

## 📄 许可证

本项目仅供学习和个人使用。

---

## 📮 联系方式

如有问题或建议，请通过 GitHub Issues 联系。
