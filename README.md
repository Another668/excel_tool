# Excel批量加密解密工具

<div align="center">

**一个高效的Windows桌面应用，用于批量加密和解密Excel文件**

[![Version](https://img.shields.io/badge/version-3.6.0-blue.svg)](https://github.com/excel-protector/releases)
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
  - 自动检测D:\tools\upx.exe
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
