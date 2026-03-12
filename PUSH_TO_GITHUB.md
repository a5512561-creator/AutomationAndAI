# 上傳此專案到 GitHub

請在 **已安裝 Git 且 `git` 在 PATH 的環境**（例如 Git Bash、或重新開啟的 PowerShell/命令提示字元）中執行以下指令。

在專案根目錄 `d:\CursorProject\AutomationAndAI` 下執行：

```bash
cd d:\CursorProject\AutomationAndAI

git init
git add .
git commit -m "Add weekly_report_checker project"
git branch -M main
git remote add origin https://github.com/a5512561-creator/AutomationAndAI.git
git push -u origin main
```

若之前已經執行過 `git init` 或已設定過 `origin`，可改為：

```bash
cd d:\CursorProject\AutomationAndAI
git add .
git commit -m "Add weekly_report_checker project"
git push -u origin main
```

完成後，程式碼會出現在：  
https://github.com/a5512561-creator/AutomationAndAI  
且 **weekly_report_checker** 會位於倉庫內的 `weekly_report_checker` 資料匣中。
