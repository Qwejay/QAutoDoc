@echo off
:: 设置字符集为 UTF-8 避免中文日志乱码
chcp 65001 >nul
setlocal enabledelayedexpansion

echo ===================================================
echo             QAutoDoc 自动化打包工具
echo ===================================================

:: 1. 检测环境中的 Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未在系统环境中检测到 Python，请先安装 Python 并添加至 PATH。
    goto :EXIT
)

:: 2. 自动定位主 Python 脚本
set "SCRIPT_NAME="
if exist "main.py" (
    set "SCRIPT_NAME=main.py"
) else if exist "qautodoc.py" (
    set "SCRIPT_NAME=qautodoc.py"
) else (
    for %%f in (*.py) do (
        set "SCRIPT_NAME=%%f"
        goto :FOUND_SCRIPT
    )
)

:FOUND_SCRIPT
if "%SCRIPT_NAME%"=="" (
    echo [错误] 当前目录下未找到任何 Python 脚本。
    goto :EXIT
)
echo [信息] 目标打包脚本: %SCRIPT_NAME%

:: 3. 建立临时虚拟环境
echo [步骤 1/4] 正在创建独立的临时虚拟环境 (build_env)...
if exist "build_env" rmdir /s /q "build_env"
python -m venv build_env
if %errorlevel% neq 0 (
    echo [错误] 创建虚拟环境失败。
    goto :EXIT
)

:: 4. 激活虚拟环境并安装依赖
echo [步骤 2/4] 激活虚拟环境并安装依赖依赖项...
call build_env\Scripts\activate.bat

python -m pip install --upgrade pip -i https://pypi.tuna.tsinghua.edu.cn/simple
pip install pyinstaller pyqt5 python-docx python-pptx openpyxl -i https://pypi.tuna.tsinghua.edu.cn/simple
if %errorlevel% neq 0 (
    echo [错误] 依赖安装失败，请检查网络连接。
    call deactivate
    goto :CLEANUP
)

:: 5. 执行打包命令
echo [步骤 3/4] 启动 PyInstaller 执行高压打包...
:: --noconsole: 隐藏黑窗口运行 GUI
:: --onefile: 打包为单文件 EXE
:: --clean: 清理 PyInstaller 自身缓存
pyinstaller --clean --noconsole --onefile --name="QAutoDoc" "%SCRIPT_NAME%"
set "BUILD_STATUS=%errorlevel%"

:: 退出虚拟环境
call deactivate

:: 6. 清理缓存和临时环境
:CLEANUP
echo [步骤 4/4] 正在进行后期清理，移除所有临时缓存...

:: 移除临时虚拟环境
if exist "build_env" (
    echo - 移除虚拟环境文件夹
    rmdir /s /q "build_env"
)

:: 移除 PyInstaller 编译中间件
if exist "build" (
    echo - 移除 PyInstaller 构建缓存 (build)
    rmdir /s /q "build"
)

:: 移除 Spec 配置文件
if exist "QAutoDoc.spec" (
    echo - 移除打包 Spec 配置文件
    del /f /q "QAutoDoc.spec"
)

:: 移除局部 pycache
if exist "__pycache__" (
    echo - 移除字节码缓存 (__pycache__)
    rmdir /s /q "__pycache__"
)

echo ===================================================
if %BUILD_STATUS% equ 0 (
    echo [完成] 打包顺利结束！
    echo [成果] 生成的目标程序位于: dist\QAutoDoc.exe
) else (
    echo [失败] 打包过程中出现异常，请查看上方错误日志。
)
echo ===================================================

:EXIT
pause