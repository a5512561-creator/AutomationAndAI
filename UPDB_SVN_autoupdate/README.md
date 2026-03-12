# UPDV_SVN_autoupdate

依文字檔批次將同仁加入 UPDB-manager 專案與群組，並可選是否一併加入 SVN 權限。

## 文件索引

- **[實作方案](docs/PLAN.md)**：整體架構、技術選型、流程與實作順序（換機後從這裡接續開發）。
- **[輸入文字檔格式](docs/INPUT_FORMAT.md)**：來源檔欄位說明與範例。

## 上傳至 AutomationAndAI 儲存庫

此專案可放在 [AutomationAndAI](https://github.com/a5512561-creator/AutomationAndAI) 的 `UPDV_SVN_autoupdate` 資料夾內。在**已安裝 Git 且 `git` 在 PATH 中**的 PowerShell 裡，於本專案目錄執行：

```powershell
.\PUSH_TO_AUTOMATION_AND_AI.ps1
```

腳本會自動 clone（若尚未存在）AutomationAndAI、將本專案檔案複製到 `AutomationAndAI/UPDV_SVN_autoupdate/`、commit 並 push。

## 換機後繼續開發

1. 克隆或拉取 AutomationAndAI（或本專案）後，先閱讀 `docs/PLAN.md` 了解需求與實作步驟。
2. 若為獨立儲存庫且尚未初始化 Git，可執行：
   ```bash
   git init
   git add .
   git commit -m "Initial: docs and project layout"
   git remote add origin <你的遠端倉庫 URL>
   git push -u origin main
   ```
3. 依 `docs/PLAN.md` 的「實作順序建議」從步驟 1 開始實作程式碼。

## 需求摘要

- **輸入**：文字檔，每行指定「專案、群組、是否加 SVN、工號」。
- **行為**：以瀏覽器自動化操作 UPDB-manager（登入時由使用者手動完成 OTP），依序選專案、加群組成員，必要時進入 SVN 頁面加入人員。
- **技術**：預計使用 Python + Playwright，詳見 `docs/PLAN.md`。
