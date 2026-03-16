@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ===== Step 1: 從 OneNote 擷取頁面資料 =====
powershell -ExecutionPolicy Bypass -File src\Get-OneNotePages.ps1
if %ERRORLEVEL% neq 0 (
    echo [錯誤] PowerShell 擷取失敗，請確認 OneNote 已開啟並同步。
    pause
    exit /b 1
)

echo.
echo ===== Step 2: 分析週報並產出 Excel =====
py src\weekly_report_checker.py
if %ERRORLEVEL% neq 0 (
    echo [錯誤] Python 分析失敗。
    pause
    exit /b 1
)

echo.
echo ===== 完成！請查看 output 資料夾中的 Excel 檔案 =====
pause
