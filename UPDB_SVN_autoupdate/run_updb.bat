@echo off
chcp 65001 >nul
cd /d "%~dp0"

if "%~1"=="" (
    echo 使用方式: run_updb.bat [成員名單.txt]
    echo 未指定檔案時使用 members.example.txt
    echo.
    set "INPUT=members.example.txt"
) else (
    set "INPUT=%~1"
)

if not exist "%INPUT%" (
    echo 找不到檔案: %INPUT%
    pause
    exit /b 1
)

echo 執行: python main.py %INPUT%
echo.
python main.py "%INPUT%"
pause
