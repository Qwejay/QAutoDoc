@echo off
setlocal enabledelayedexpansion

echo ===================================================
echo           Universal Python Packer
echo ===================================================

:: 1. Check Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python not found in PATH.
    goto :EXIT
)

:: 2. Get script file
set "SCRIPT_NAME=%~1"
set "APP_NAME="

if not "%SCRIPT_NAME%"=="" (
    if not exist "%SCRIPT_NAME%" (
        echo [ERROR] File not exist: %SCRIPT_NAME%
        goto :EXIT
    )
    set "APP_NAME=%~n1"
) else (
    if exist "main.py" (
        set "SCRIPT_NAME=main.py"
        set "APP_NAME=main"
    ) else (
        for %%f in (*.py) do (
            set "SCRIPT_NAME=%%f"
            set "APP_NAME=%%~nf"
            goto :FOUND
        )
    )
)

:FOUND
if "%SCRIPT_NAME%"=="" (
    echo [ERROR] No .py file found.
    echo Drag .py file onto this .bat
    goto :EXIT
)

echo [INFO] Script: %SCRIPT_NAME%
echo [INFO] Output: %APP_NAME%.exe
echo.

:: 3. Console window option
set "CONSOLE_CMD=--noconsole"
set /p show_cmd="Show console window? (Y=CLI, N=GUI) [Y/N] (default N): "
if /i "%show_cmd%"=="y" set "CONSOLE_CMD="

:: 4. Create venv
echo [Step 1/4] Creating temporary venv...
if exist "build_env" rmdir /s /q "build_env"
python -m venv build_env
if %errorlevel% neq 0 (
    echo [ERROR] Failed to create venv.
    goto :EXIT
)

:: 5. Install dependencies
echo [Step 2/4] Installing PyInstaller...
call build_env\Scripts\activate.bat

python -m pip install --upgrade pip -i https://pypi.tuna.tsinghua.edu.cn/simple >nul
pip install pyinstaller -i https://pypi.tuna.tsinghua.edu.cn/simple

if exist "requirements.txt" (
    echo [INFO] Installing from requirements.txt...
    pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
)

:: 6. Pack
echo [Step 3/4] Running PyInstaller...
pyinstaller --clean %CONSOLE_CMD% --onefile --name="%APP_NAME%" "%SCRIPT_NAME%"
set "BUILD_STATUS=%errorlevel%"

call deactivate

:: 7. Cleanup
:CLEANUP
echo [Step 4/4] Cleaning up...
if exist "build_env" rmdir /s /q "build_env"
if exist "build" rmdir /s /q "build"
if exist "%APP_NAME%.spec" del /q "%APP_NAME%.spec"
if exist "__pycache__" rmdir /s /q "__pycache__"

echo ===================================================
if %BUILD_STATUS% equ 0 (
    echo [SUCCESS] Packing completed!
    echo [Output] dist\%APP_NAME%.exe
) else (
    echo [FAILED] Packing failed. Check messages above.
)
echo ===================================================

:EXIT
pause
