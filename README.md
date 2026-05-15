# Excel批量加密解密工具

<div align="center">

**一个高效的Windows桌面应用，用于批量加密和解密Excel文件**

[![Version](https://img.shields.io/badge/version-3.5-blue.svg)](https://github.com/excel-protector/releases)
[![Python](https://img.shields.io/badge/python-3.6+-green.svg)](https://www.python.org/)
[![PyQt5](https://img.shields.io/badge/PyQt5-5.15+-orange.svg)](https://www.riverbankcomputing.com/software/pyqt/)
[![License](https://img.shields.io/badge/license-MIT-lightgrey.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)](https://www.microsoft.com/windows)

</div>

---

## 📝 更新日志

### [v3.4.1] - 2026-05-15

#### 🧪 版本同步验证
- **版本同步机制验证
  - 验证了从README.md到Python程序的版本同步功能
  - 确保打包流程中版本同步正确执行
  - 完成了两次不同版本的打包测试

---

### [v3.4] - 2026-05-15

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

### [v3.3] - 2026-05-15

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
  - 自动检测D:\tools\upx.exe
  - 使用--best --lzma最佳压缩模式
  - 显示压缩前后对比和节省空间百分比
  - 优雅降级：无UPX时仍能正常打包

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

### [v3.2] - 2026-04-28

#### 🎨 图标配置
- 添加自定义应用程序图标（图标.png）
  - Excel文件与锁的组合设计，直观展示软件功能
  - 绿色背景配蓝色锁图标，视觉识别度高
- EXE文件图标嵌入
  - 可执行文件显示自定义图标
  - 任务栏最小化时显示图标
  - 窗口标题栏显示图标
- 图标格式转换
  - 新增convert_icon.py脚本
  - 支持PNG自动转换为ICO格式
  - 多尺寸图标生成（256x256到16x16）

#### 🔧 打包脚本优化
- build.bat升级到v3.2版本
- 打包流程从6步扩展到7步
- 新增图标转换步骤（步骤[2/7]）
  - 自动检测图标.png文件
  - 自动安装Pillow库进行格式转换
  - 智能回退机制（转换失败时使用默认图标）
- 动态图标嵌入配置
  - 使用%ICON_CONFIG%变量动态指定图标
  - 兼容无图标情况下的正常打包
- 构建流程优化
  - 清理步骤移到最前（步骤[1/7]）
  - 避免清理刚生成的图标文件
  - 提高打包成功率

#### 💻 代码改进
- 新增get_resource_path()函数
  - 支持PyInstaller打包后的资源路径访问
  - 兼容开发环境和打包环境
  - 使用sys._MEIPASS检测打包状态
- 修改main()函数图标加载逻辑
  - 优先加载自定义图标.ico
  - 降级处理：使用系统默认图标
  - 完整的异常捕获机制
- 增强窗口图标设置
  - 应用程序级别图标（QApplication）
  - 窗口级别图标（QWidget）
  - 确保所有场景下图标正确显示

---

### [v3.1] - 2026-04-28

#### 🎨 新增功能
- 全新现代化GUI界面，采用Material Design风格配色方案
- 完整菜单栏（文件、编辑、帮助）及快捷键支持
  - Ctrl+I/O：快速打开输入/输出文件夹
  - Ctrl+A/D：全选/取消全选文件
  - Ctrl+L：清空日志
  - Ctrl+Q：退出程序
  - F1：查看使用说明
- 配置自动保存与加载功能（tool_config.json）
  - 记忆文件夹路径、操作模式、后缀设置
  - 自动加载上次使用的密码本路径
- 按钮增强，添加emoji图标提升辨识度
- 使用说明对话框（富文本格式）
- 关于对话框

#### ✨ 界面优化
- 圆角卡片布局，白色背景配浅灰色主界面
- 按钮悬停和按下的视觉反馈效果
- 进度条样式美化，支持百分比显示
- 表格交替行颜色，提升数据可读性
- 智能列宽分配（固定列+自动拉伸列）
- 输入框焦点高亮效果
- 统一按钮高度，视觉更加整齐

#### 🔧 技术改进
- 添加requirements.txt依赖管理文件
- 优化UI组件结构，模块化设计
  - create_menu_bar()：创建菜单栏
  - create_folder_section()：文件夹设置区域
  - create_function_section()：功能设置区域
  - create_password_section()：密码设置区域
  - create_files_section()：文件列表区域
  - create_progress_section()：进度和日志区域
  - create_button_section()：按钮区域
  - apply_styles()：全局样式应用
- 新增load_config()和save_config()方法
- 关闭程序时自动保存配置

---

### [v3.0] - 2024-01-01

#### 🎨 初始版本功能
- 批量加密Excel文件，支持设置密码
- 批量解密Excel文件，移除密码保护
- 密码本支持（CSV格式）
  - 自动编码检测（UTF-8、GBK等）
  - 自动匹配文件名和密码
  - 支持注释行（#开头）
- 统一密码设置功能
- 自定义文件名后缀
- 文件列表可视化展示
  - 文件选择（全选/反选）
  - 状态显示（待处理/成功/失败）
  - 密码和备注编辑
  - 新文件名预览
- 实时进度条和操作日志
- 日志导出功能
- 密码本模板导出
- 多线程处理，支持取消操作
- 静默处理（Excel后台运行）
- 原文件保护（不修改原文件）

#### 🔧 技术实现
- PyQt5 GUI框架
- win32com调用Excel COM组件
- csv模块处理密码本（替代pandas，减小体积）
- QThread异步处理，避免界面卡顿
- 完整的错误处理和异常捕获

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

**详细步骤：**

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

**支持：**
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
