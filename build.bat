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
echo   版本: v3.6.0
echo   日期: 2026-06-12
echo   工作目录: %SCRIPT_DIR%
echo ==========================================
echo.

REM ==================== 清理旧文件 ====================
echo [1/10] 清理旧的打包文件...

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
echo [2/10] 处理图标资源...

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
    py -3.13 -c "import PIL" >nul 2>&1
    if errorlevel 1 (
        echo [提示] 未安装Pillow，正在安装...
        py -3.13 -m pip install Pillow
        if errorlevel 1 (
            echo [警告] Pillow安装失败，将使用默认图标
            set "ICON_CONFIG=--icon=NONE"
            set "HAS_ICON=0"
            goto skip_icon_convert
        )
    )

    py -3.13 "%SCRIPT_DIR%\convert_icon.py"
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
echo [3/10] 检查运行环境...

REM 诊断：先打印 py -3.13 实际指向哪个解释器，让 pip/python 错位问题立即可见
for /f "delims=" %%P in ('py -3.13 -c "import sys; print(sys.executable)" 2^>nul') do set PY_EXE=%%P
echo [信息] 打包将使用: %PY_EXE%

py -3.13 --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到Python！请先安装Python 3.6+。
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

for /f "tokens=2" %%i in ('py -3.13 --version 2^>^&1') do set PYTHON_VERSION=%%i
echo [成功] Python版本: %PYTHON_VERSION%

py -3.13 -c "import struct; print('   架构:', '64位' if struct.calcsize('P')*8 == 64 else '32位')" 2>nul

echo.

REM ==================== 依赖检查 ====================
echo [4/10] 检查项目依赖...

py -3.13 -c "import PyQt5" >nul 2>&1
if errorlevel 1 (
    echo [提示] 未安装PyQt5，正在安装依赖...
    py -3.13 -m pip install -r "%SCRIPT_DIR%\requirements.txt"
    if errorlevel 1 (
        echo [错误] 依赖安装失败！
        pause
        exit /b 1
    )
) else (
    echo [成功] PyQt5已安装
)

py -3.13 -c "import win32com.client" >nul 2>&1
if errorlevel 1 (
    echo [提示] 未安装pywin32，正在安装...
    py -3.13 -m pip install pywin32
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
echo [5/10] 检查打包工具PyInstaller...

py -3.13 -c "import PyInstaller" >nul 2>&1
if errorlevel 1 (
    echo [提示] 未安装PyInstaller，正在安装...
    py -3.13 -m pip install pyinstaller
    if errorlevel 1 (
        echo [错误] PyInstaller安装失败！
        pause
        exit /b 1
    )
) else (
    echo [成功] PyInstaller已安装
)

for /f "tokens=3" %%i in ('py -3.13 -m pip show pyinstaller 2^>^&1 ^| findstr Version') do set PYINSTALLER_VERSION=%%i
echo [信息] PyInstaller版本: %PYINSTALLER_VERSION%

echo.

REM ==================== 源代码健全性检查 ====================
echo [6/10] 检查源代码健全性（Git冲突标记 + Python语法）...

set "SOURCE_FILE=%SCRIPT_DIR%\Excel批量解密工具与密码管理.py"
set "SOURCE_AUX=convert_icon.py requirements.txt"
set "CHECK_FAILED=0"

REM 检查1：主源文件 Git 冲突标记
py -3.13 -c "import re,sys;src=open(r'%SOURCE_FILE%','rb').read().decode('utf-8',errors='replace');m=re.findall(r'^(<<<<<<<|=======|>>>>>>>)',src,re.M);print('found',len(m),'conflict marker(s)') if m else None;sys.exit(1 if m else 0)" >nul 2>&1
if errorlevel 1 (
    echo [错误] 主源文件存在未解决的Git合并冲突标记！
    echo   文件: %SOURCE_FILE%
    echo   请先解决冲突再打包，否则程序无法启动。
    set "CHECK_FAILED=1"
) else (
    echo [成功] 主源文件无Git冲突标记
)

REM 检查2：主源文件 Python 语法
py -3.13 -m py_compile "%SOURCE_FILE%" >nul 2>&1
if errorlevel 1 (
    echo [错误] 主源文件Python语法错误！
    echo   文件: %SOURCE_FILE%
    set "CHECK_FAILED=1"
) else (
    echo [成功] 主源文件Python语法检查通过
)

REM 检查3：辅助脚本与配置文件冲突扫描
for %%F in (%SOURCE_AUX%) do (
    if exist "%SCRIPT_DIR%\%%F" (
        findstr /b /c:"<<<<<<<" /c:">>>>>>>" "%SCRIPT_DIR%\%%F" >nul 2>&1
        if not errorlevel 1 (
            echo [错误] %%F 存在未解决的Git合并冲突标记！
            set "CHECK_FAILED=1"
        ) else (
            echo [成功] %%F 无Git冲突标记
        )
    )
)

if "%CHECK_FAILED%"=="1" (
    echo.
    echo [失败] 源代码健全性检查未通过！请修复后重试。
    pause
    exit /b 1
)

echo.

REM ==================== 版本同步 ====================
echo [7/10] 同步版本号...

py -3.13 "%SCRIPT_DIR%\sync_version.py"
if errorlevel 1 (
    echo [警告] 版本同步出现问题，但将继续打包...
)

echo.

REM ==================== 构造带版本号的 EXE 名称 ====================
REM 从 Python 源文件读取 __version__，构造文件名（例：Excel批量加密解密工具_v3.6.6.exe）
REM 使用辅助脚本 get_version.py 避免 cmd 处理中文路径+复杂转义的问题
set "PKG_VERSION="
for /f "delims=" %%V in ('py -3.13 "%SCRIPT_DIR%\get_version.py" "%SCRIPT_DIR%\Excel批量解密工具与密码管理.py" 2^>nul') do (
    if not defined PKG_VERSION set "PKG_VERSION=%%V"
)
if "%PKG_VERSION%"=="" set "PKG_VERSION=unknown"
if "%PKG_VERSION:~0,3%"=="ERR" set "PKG_VERSION=unknown"
set "EXE_BASE_NAME=Excel批量加密解密工具"
set "EXE_NAME=%EXE_BASE_NAME%_v%PKG_VERSION%"
echo [信息] 本次打包将生成: %EXE_NAME%.exe
echo.

REM ==================== 预先检测 UPX 路径 ====================
REM 提前探测 UPX（仅在 [10/10] 手动调用以获最高压缩率 --best --lzma）
REM 不传 --upx-dir 给 PyInstaller，避免它先用默认参数压缩后无法再深度压缩
set "UPX_PATH="
if exist "D:\UPX\upx.exe" (
    set "UPX_PATH=D:\UPX\upx.exe"
) else if exist "D:\tools\upx.exe" (
    set "UPX_PATH=D:\tools\upx.exe"
) else (
    where upx >nul 2>&1
    if not errorlevel 1 (
        for /f "delims=" %%U in ('where upx') do (
            if not defined UPX_PATH set "UPX_PATH=%%U"
        )
    )
)
if defined UPX_PATH (
    echo [信息] 已检测到UPX: !UPX_PATH!（将在 [10/10] 用 --best --lzma 深度压缩）
) else (
    echo [提示] 未检测到UPX，跳过压缩
)
echo.

REM ==================== 开始打包 ====================
echo [8/10] 开始打包为EXE文件...
echo.

py -3.13 -m PyInstaller ^
    --onefile ^
    --windowed ^
    --name "%EXE_NAME%" ^
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
echo [9/10] 验证打包结果...

set "EXE_PATH=%SCRIPT_DIR%\dist\%EXE_NAME%.exe"
if exist "%EXE_PATH%" (
    for %%i in ("%EXE_PATH%") do set EXE_SIZE=%%~zi
    set /a EXE_SIZE_MB=!EXE_SIZE! / 1048576
    set /a EXE_SIZE_KB=!EXE_SIZE! / 1024
    echo [成功] EXE文件已生成！
    echo   文件位置: %EXE_PATH%
    REM 关键：括号必须用 ^ 转义，否则 cmd 会把它当成子代码块分隔符，
    REM       导致 echo 行被截断、整个 if/else 结构错乱、step10 永远跑不到
    echo   原始大小: !EXE_SIZE_MB! MB ^(!EXE_SIZE_KB! KB^)
) else (
    echo [错误] 未找到生成的EXE文件！
    echo   期望位置: %EXE_PATH%
    echo 请检查打包日志
    pause
    exit /b 1
)

echo.

REM ==================== UPX压缩 ====================
echo [10/10] 使用UPX压缩EXE文件...

REM 多路径探测 UPX 可执行文件（按优先级）
REM 1) D:\UPX\upx.exe       （用户指定路径，优先）
REM 2) D:\tools\upx.exe     （旧路径，兜底）
REM 3) 系统 PATH 中的 upx
set "UPX_PATH="
if exist "D:\UPX\upx.exe" (
    set "UPX_PATH=D:\UPX\upx.exe"
) else if exist "D:\tools\upx.exe" (
    set "UPX_PATH=D:\tools\upx.exe"
) else (
    where upx >nul 2>&1
    if not errorlevel 1 (
        for /f "delims=" %%U in ('where upx') do (
            if not defined UPX_PATH set "UPX_PATH=%%U"
        )
    )
)

if defined UPX_PATH (
    REM 记录压缩前体积
    for %%i in ("%EXE_PATH%") do set ORIGINAL_SIZE=%%~zi
    set /a ORIGINAL_SIZE_MB=!ORIGINAL_SIZE! / 1048576
    set /a ORIGINAL_SIZE_KB=!ORIGINAL_SIZE! / 1024

    echo.
    echo [信息] 找到UPX工具: !UPX_PATH!
    echo [信息] 压缩前大小: !ORIGINAL_SIZE_MB! MB (!ORIGINAL_SIZE_KB! KB)
    echo [信息] 正在压缩（参数: --best --lzma），这可能需要几分钟...

    REM 关键：清空会污染 UPX 启动的环境变量
    REM 系统中残留的 set upx=D:\tools\upx 之类的变量，UPX 启动时会读取 %UPX%
    REM 作为自身目录，若路径无效会直接报：
    REM   "invalid string 'D:\tools\upx' in environment variable 'UPX'"
    REM UPX 只认“合法且与可执行文件同目录”的路径，任意赋值仍会触发校验失败。
    REM 唯一稳妥的做法就是清空它（让其走默认探测），我们仍用绝对路径调用。
    set "UPX="

    REM 备份原文件以便压缩失败时回滚
    copy /y "%EXE_PATH%" "%EXE_PATH%.unpacked" >nul

    REM --force: 允许压缩启用了 GUARD_CF (Windows Control Flow Guard) 的 PE 文件
    REM         UPX 5.x 默认拒绝此类文件，加 --force 后可继续
    "!UPX_PATH!" --best --lzma --force "%EXE_PATH%"

    if errorlevel 1 (
        echo.
        echo [警告] UPX压缩失败 (errorlevel=!errorlevel!)，回滚到原始文件
        if exist "%EXE_PATH%.unpacked" (
            move /y "%EXE_PATH%.unpacked" "%EXE_PATH%" >nul
            echo [信息] 已恢复未压缩版本
        )
    ) else (
        REM 删除备份
        if exist "%EXE_PATH%.unpacked" del /f /q "%EXE_PATH%.unpacked" >nul

        REM 记录压缩后体积
        for %%i in ("%EXE_PATH%") do set COMPRESSED_SIZE=%%~zi
        set /a COMPRESSED_SIZE_MB=!COMPRESSED_SIZE! / 1048576
        set /a COMPRESSED_SIZE_KB=!COMPRESSED_SIZE! / 1024
        set /a SAVED=!ORIGINAL_SIZE! - !COMPRESSED_SIZE!
        set /a SAVED_KB=!SAVED! / 1024
        set /a SAVED_PERCENT=(!SAVED! * 100) / !ORIGINAL_SIZE!

        REM 把压缩前后数据写入日志文件（追加模式）
        set "LOG_FILE=%SCRIPT_DIR%\build_size_log.txt"
        (
            echo ==========================================
            echo 打包时间: !DATE! !TIME!
            echo 版本号:   !PKG_VERSION!
            echo UPX版本:  5.1.1
            echo UPX路径:  !UPX_PATH!
            echo 压缩参数:  --best --lzma
            echo ------------------------------------------
            echo 压缩前:   !ORIGINAL_SIZE! bytes ^(!ORIGINAL_SIZE_MB! MB / !ORIGINAL_SIZE_KB! KB^)
            echo 压缩后:   !COMPRESSED_SIZE! bytes ^(!COMPRESSED_SIZE_MB! MB / !COMPRESSED_SIZE_KB! KB^)
            echo 节省:     !SAVED! bytes ^(!SAVED_KB! KB^)
            echo 压缩率:   !SAVED_PERCENT!%%
            echo ==========================================
            echo.
        ) >> "!LOG_FILE!"

        echo.
        echo [成功] 压缩完成！
        echo   压缩后大小: !COMPRESSED_SIZE_MB! MB ^(!COMPRESSED_SIZE_KB! KB^)
        echo   节省空间:   !SAVED_PERCENT!%% ^(!SAVED_KB! KB^)
        echo   体积日志:   !LOG_FILE!
        echo.
    )
) else (
    echo.
    echo [提示] 未找到UPX工具（已检查 D:\UPX\upx.exe、D:\tools\upx.exe、PATH）
    echo [提示] 跳过压缩步骤，使用原始文件
    echo [提示] 安装UPX后将可减小约 50-70%% 的EXE体积
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
