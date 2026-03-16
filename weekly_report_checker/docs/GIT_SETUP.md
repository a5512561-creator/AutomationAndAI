# Git 環境安裝與 GitHub 設定

## 1. 安裝 Git for Windows

1. 前往 [https://git-scm.com/download/win](https://git-scm.com/download/win) 下載安裝程式。
2. 執行安裝時，在「Adjusting your PATH environment」步驟請選擇 **Git from the command line and also from 3rd-party software**，以便在 PowerShell 與 CMD 使用 `git` 指令。
3. 其餘選項可維持預設，完成安裝後重新開啟終端機。

## 2. 在專案內初始化 Git

開啟 PowerShell 或命令提示字元，執行：

```powershell
cd D:\AutomationAndAI\weekly_report_checker
git init
```

## 3. 第一次提交

```powershell
git add .
git commit -m "Initial commit: Weekly Report Checker project"
```

## 4. 連接 GitHub 並推送

1. 在 [GitHub](https://github.com) 建立新 repository，名稱例如 `Weekly_Report_Checker`（可不勾選 README，因專案已有）。
2. 在專案目錄執行（請將 `你的帳號` 換成你的 GitHub 使用者名稱）：

```powershell
git remote add origin https://github.com/你的帳號/Weekly_Report_Checker.git
git branch -M main
git push -u origin main
```

之後的版控都在 `D:\AutomationAndAI\weekly_report_checker` 目錄下使用 `git add`、`git commit`、`git push` 等指令即可。
