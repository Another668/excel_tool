@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion

title Excel批量加密解密工具 - 一键打包脚本

REM ==================== 根目录路径检测 ====================
set "SCRIPT_DIR=%~dp0"
set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"
cd /d "%SCRIPT_DIR%"

echo ==========================================
echo   Excel批量加密解密工具 - 打包脚本
echo   版本: v3.5
echo   日期: 2026-05-15
echo   工作目录: %SCRIPT_DIR%
echo ==========================================
echo.

REM ==================== 清理旧文件 ====================
echo [1/9] 清理旧的打包文件...

if exist "%SCRIPT_DIR%\dist" (
    rmdir /s /q "%SCRIPT_DIR%\dist"
    echo [成功] 已删除dist目录
)

if exist "%SCRIPT_DIR%\build" (
    rmdir /s /q "%SCRIPT_DIR%\build"
    echo [成功] 已删除build目录
)

if exist "%SCRIPT_DIR%\Excel批量加密解密工具.spec" (
    del /f /q "%SCRIPT_DIR%\Excel批量加密解密工具.spec"
    echo [成功] 已删除旧spec文件
)

echo.

REM ==================== 图标处理 ====================
echo [2/9] 处理图标资源...

set "ICON_PNG=%SCRIPT_DIR%\图标.png"
set "ICON_ICO=%SCRIPT_DIR%\图标.ico"
set "ICON_CONFIG=--icon=NONE"
set "HAS_ICON=0"

if exist "%ICON_ICO%" (
    echo [信息] 找到图标文件: %ICON_ICO%
    set "ICON_CONFIG=--icon=%ICON_ICO%"
    set "HAS_ICON=1"
    echo [成功] 将使用现有ICO图标
) else if exist "%ICON_PNG%" (
    echo [信息] 找到PNG图标，正在转换为ICO格式...
    python -c "import PIL" >nul 2>&1
    if errorlevel 1 (
        echo [提示] 未安装Pillow，正在安装...
        pip install Pillow
        if errorlevel 1 (
            echo [警告] Pillow安装失败，将使用默认图标
            set "ICON_CONFIG=--icon=NONE"
            set "HAS_ICON=0"
            goto skip_icon_convert
        )
    )
    
    python "%SCRIPT_DIR%\convert_icon.py"
    if errorlevel 1 (
        echo [警告] 图标转换失败，将使用默认图标
        set "ICON_CONFIG=--icon=NONE"
        set "HAS_ICON=0"
    ) else (
        echo [成功] 图标转换完成
        if exist "%ICON_ICO%" (
            set "ICON_CONFIG=--icon=%ICON_ICO%"
            set "HAS_ICON=1"
        )
    )
) else (
    echo [警告] 未找到任何图标文件
    echo   尝试查找: %ICON_PNG%
    echo   尝试查找: %ICON_ICO%
    echo [提示] 将使用默认图标
    set "ICON_CONFIG=--icon=NONE"
    set "HAS_ICON=0"
)

:skip_icon_convert
echo.

REM ==================== 环境检查 ====================
echo [3/9] 检查运行环境...

python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到Python！请先安装Python 3.6+。
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

for /f "tokens=2" %%i in ('python --version 2^>^&1') do set PYTHON_VERSION=%%i
echo [成功] Python版本: %PYTHON_VERSION%

python -c "import struct; print('   架构:', '64位' if struct.calcsize('P')*8 == 64 else '32位')" 2>nul

echo.

REM ==================== 依赖检查 ====================
echo [4/9] 检查项目依赖...

python -c "import PyQt5" >nul 2>&1
if errorlevel 1 (
    echo [提示] 未安装PyQt5，正在安装依赖...
    pip install -r "%SCRIPT_DIR%\requirements.txt"
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
echo [5/9] 检查打包工具PyInstaller...

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

for /f "tokens=3" %%i in ('pip show pyinstaller 2^>^&1 ^| findstr Version') do set PYINSTALLER_VERSION=%%i
echo [信息] PyInstaller版本: %PYINSTALLER_VERSION%

echo.

REM ==================== 版本同步 ====================
echo [6/9] 同步版本号...

python "%SCRIPT_DIR%\sync_version.py"
if errorlevel 1 (
    echo [警告] 版本同步出现问题，但将继续打包...
)

echo.

REM ==================== 开始打包 ====================
echo [7/9] 开始打包为EXE文件...
echo.

python -m PyInstaller ^
    --onefile ^
    --windowed ^
    --name "Excel批量加密解密工具" ^
    %ICON_CONFIG% ^
    --add-data "%SCRIPT_DIR%\requirements.txt;." ^
    --exclude-module tkinter ^
    --exclude-module matplotlib ^
    --exclude-module pandas ^
    --exclude-module numpy ^
    --exclude-module pytest ^
    --hidden-import=win32com ^
    --hidden-import=win32com.client ^
    --hidden-import=pythoncom ^
    --hidden-import=PyQt5 ^
    --hidden-import=PyQt5.QtCore ^
    --hidden-import=PyQt5.QtGui ^
    --hidden-import=PyQt5.QtWidgets ^
    --noconfirm ^
    "%SCRIPT_DIR%\Excel批量解密工具与密码管理.py"

if errorlevel 1 (
    echo.
    echo [错误] 打包失败！请检查错误信息。
    pause
    exit /b 1
)

echo.

REM ==================== 验证结果 ====================
echo [8/9] 验证打包结果...

set "EXE_PATH=%SCRIPT_DIR%\dist\Excel批量加密解密工具.exe"
if exist "%EXE_PATH%" (
    for %%i in ("%EXE_PATH%") do set EXE_SIZE=%%~zi
    set /a EXE_SIZE_MB=!EXE_SIZE! / 1048576
    set /a EXE_SIZE_KB=!EXE_SIZE! / 1024
    echo [成功] EXE文件已生成！
    echo   文件位置: %EXE_PATH%
    echo   原始大小: !EXE_SIZE_MB! MB (!EXE_SIZE_KB! KB)
) else (
    echo [错误] 未找到生成的EXE文件！
    echo   期望位置: %EXE_PATH%
    echo 请检查打包日志
    pause
    exit /b 1
)

echo.

REM ==================== UPX压缩 ====================
echo [9/9] 使用UPX压缩EXE文件...

set "UPX_PATH=D:\tools\upx.exe"
if exist "%UPX_PATH%" (
    echo [信息] 找到UPX工具: %UPX_PATH%
    echo [信息] 正在压缩，这可能需要几分钟...
    
    "%UPX_PATH%" --best --lzma "%EXE_PATH%"
    
    if errorlevel 1 (
        echo [警告] UPX压缩失败，使用原始文件
    ) else (
        for %%i in ("%EXE_PATH%") do set COMPRESSED_SIZE=%%~zi
        set /a COMPRESSED_SIZE_MB=!COMPRESSED_SIZE! / 1048576
        set /a COMPRESSED_SIZE_KB=!COMPRESSED_SIZE! / 1024
        set /a SAVED=!EXE_SIZE! - !COMPRESSED_SIZE!
        set /a SAVED_PERCENT=(!SAVED! * 100) / !EXE_SIZE!
        echo [成功] 压缩完成！
        echo   压缩后大小: !COMPRESSED_SIZE_MB! MB (!COMPRESSED_SIZE_KB! KB)
        echo   节省空间: !SAVED_PERCENT!%%
    )
) else (
    echo [提示] 未找到UPX工具: %UPX_PATH%
    echo [提示] 跳过压缩步骤，使用原始文件
)

echo.
echo ==========================================
echo   打包完成！
echo   请将dist目录下的EXE文件复制到目标计算机使用
echo   目标计算机需要安装Microsoft Excel
echo ==========================================

echo.

REM 清理临时文件
echo [清理] 删除临时构建文件...
if exist "%SCRIPT_DIR%\build" (
    rmdir /s /q "%SCRIPT_DIR%\build"
    echo [成功] 已清理build目录
)

echo.
echo 按任意键退出...
pause >nul
