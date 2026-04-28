@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion

title Excel批量加密解密工具 - 一键打包脚本

echo ==========================================
echo   Excel批量加密解密工具 - 打包脚本
echo   版本: v3.2
echo   日期: 2026-04-28
echo ==========================================
echo.

REM ==================== 清理旧文件 ====================
echo [1/7] 清理旧的打包文件...

if exist "dist" (
    rmdir /s /q "dist"
    echo [成功] 已删除dist目录
)

if exist "build" (
    rmdir /s /q "build"
    echo [成功] 已删除build目录
)

if exist "Excel批量加密解密工具.spec" (
    del /f /q "Excel批量加密解密工具.spec"
    echo [成功] 已删除旧spec文件
)

echo.

REM ==================== 图标转换 ====================
echo [2/7] 转换图标为ICO格式...

if not exist "图标.png" (
    echo [警告] 未找到图标.png文件，将使用默认图标
    set ICON_CONFIG=--icon=NONE
    set HAS_ICON=0
) else (
    python -c "import PIL" >nul 2>&1
    if errorlevel 1 (
        echo [提示] 未安装Pillow，正在安装...
        pip install Pillow
        if errorlevel 1 (
            echo [警告] Pillow安装失败，将使用默认图标
            set ICON_CONFIG=--icon=NONE
            set HAS_ICON=0
            goto skip_icon_convert
        )
    )
    
    python convert_icon.py
    if errorlevel 1 (
        echo [警告] 图标转换失败，将使用默认图标
        set ICON_CONFIG=--icon=NONE
        set HAS_ICON=0
    ) else (
        echo [成功] 图标转换完成
        set ICON_CONFIG=--icon=图标.ico
        set HAS_ICON=1
    )
)

:skip_icon_convert
echo.

REM ==================== 环境检查 ====================
echo [3/7] 检查运行环境...

python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到Python！请先安装Python 3.8+。
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

for /f "tokens=2" %%i in ('python --version 2^>^&1') do set PYTHON_VERSION=%%i
echo [成功] Python版本: %PYTHON_VERSION%

python -c "import struct; print('  架构:', '64位' if struct.calcsize('P')*8 == 64 else '32位')" 2>nul

echo.

REM ==================== 依赖检查 ====================
echo [3/7] 检查项目依赖...

python -c "import PyQt5" >nul 2>&1
if errorlevel 1 (
    echo [提示] 未安装PyQt5，正在安装依赖...
    pip install -r requirements.txt
    if errorlevel 1 (
        echo [错误] 依赖安装失败！
        pause
        exit /b 1
    )
) else (
    echo [成功] PyQt5已安装
)

python -c "import win32com.client" >nul 2>&1
if errorlevel 1 (
    echo [提示] 未安装pywin32，正在安装...
    pip install pywin32
    if errorlevel 1 (
        echo [错误] pywin32安装失败！
        pause
        exit /b 1
    )
) else (
    echo [成功] pywin32已安装
)

echo.

REM ==================== PyInstaller检查 ====================
echo [4/7] 检查打包工具PyInstaller...

python -c "import PyInstaller" >nul 2>&1
if errorlevel 1 (
    echo [提示] 未安装PyInstaller，正在安装...
    pip install pyinstaller
    if errorlevel 1 (
        echo [错误] PyInstaller安装失败！
        pause
        exit /b 1
    )
) else (
    echo [成功] PyInstaller已安装
)

for /f "tokens=3" %%i in ('pip show pyinstaller 2^>nul ^| findstr Version') do set PYINSTALLER_VERSION=%%i
echo [信息] PyInstaller版本: %PYINSTALLER_VERSION%

echo.

REM ==================== 开始打包 ====================
echo [5/7] 开始打包为EXE文件...
echo.

if "%HAS_ICON%"=="1" (
    echo [信息] 正在嵌入图标...
    set ADD_DATA_ICON=--add-data "图标.ico;."
) else (
    echo [信息] 未找到图标，跳过嵌入
    set ADD_DATA_ICON=
)

python -m PyInstaller ^
    --onefile ^
    --windowed ^
    --name "Excel批量加密解密工具" ^
    %ICON_CONFIG% ^
    --add-data "requirements.txt;." ^
    %ADD_DATA_ICON% ^
    --hidden-import=win32com ^
    --hidden-import=win32com.client ^
    --hidden-import=pythoncom ^
    --hidden-import=PyQt5 ^
    --noconfirm ^
    "Excel批量解密工具与密码管理.py"

if errorlevel 1 (
    echo.
    echo [错误] 打包失败！请检查错误信息。
    pause
    exit /b 1
)

echo.

REM ==================== 验证结果 ====================
echo [6/7] 验证打包结果...

if exist "dist\Excel批量加密解密工具.exe" (
    for %%i in ("dist\Excel批量加密解密工具.exe") do set EXE_SIZE=%%~zi
    set /a EXE_SIZE_MB=!EXE_SIZE! / 1048576
    echo [成功] EXE文件已生成！
    echo   文件位置: dist\Excel批量加密解密工具.exe
    echo   文件大小: !EXE_SIZE_MB! MB
    echo.
    echo ==========================================
    echo   打包完成！
    echo   请将dist目录下的EXE文件复制到目标计算机使用
    echo   目标计算机需要安装Microsoft Excel
    echo ==========================================
) else (
    echo [错误] 未找到生成的EXE文件！
    echo 请检查打包日志
    pause
    exit /b 1
)

echo.

REM 清理临时文件
echo [清理] 删除临时构建文件...
if exist "build" (
    rmdir /s /q "build"
    echo [成功] 已清理build目录
)

echo.
echo 按任意键退出...
pause >nul
