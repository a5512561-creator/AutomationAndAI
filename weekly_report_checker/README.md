# Weekly Report Checker

透過 Power Automate 取得 OneNote 週報頁面清單，並以 Python 比對本週與上週週報，產出檢查結果 Excel。

## 專案結構

- `config/member_list.txt`：需檢查週報的同仁名單（一人一行）
- `src/weekly_report_checker.py`：主程式
- `power_automate/`：Flow 匯入套件與每個 action 說明
- `output/`：Python 輸出的 Excel 存放處

## 使用方式

### 1. 從 OneDrive 取得 JSON

- 在 Power Automate 中手動執行 flow，流程會依 `member_list.txt` 從 OneNote 取得各 section（人名）下的頁面清單，產出 JSON 並存到 OneDrive Business 您指定的資料夾（檔名含時間戳記，例如 `onenote_pages_20260309_101530.json`）。
- 將該 JSON 下載到本機，或使用 OneDrive 同步後取得本機路徑（例如：`C:\Users\你的帳號\OneDrive - 公司\WeeklyReport\onenote_pages_20260309_101530.json`）。

### 2. 執行 Python 檢查程式

```powershell
cd d:\CursorProject\Weekly_Report_Checker
pip install -r requirements.txt
python src/weekly_report_checker.py "C:\路徑\到\onenote_pages_YYYYMMDD_HHMMSS.json"
```

- **不帶參數**：程式會從目前工作目錄或專案內 `output/` 資料夾自動尋找檔名含 `onenote_pages` 且副檔名為 `.json` 的**最新**檔案。
- **指定輸出目錄**：加上 `-o` 或 `--output-dir`，例如  
  `python src/weekly_report_checker.py 路徑\file.json -o D:\Reports`

### 3. 輸出檔案說明

- **檔名**：`WeeklyReport_Check_YYYYMMDD_HHMMSS.xlsx`（以執行當下時間產生）。
- **位置**：預設為專案下的 `output/` 目錄（可透過 `-o` 指定其他目錄）。
- **欄位**：
  - **同仁名**：OneNote section 名稱（即人名）。
  - **weekly report 檢查結果**：第二步檢查的狀況，例如「本週未填寫週報」、「本週有週報但上週未填寫」、「本週與上週皆有週報，內容有更新」或「本週內容與上週完全相同，無新進展」等。

## Git 與 GitHub

若尚未安裝 Git，請先依 [docs/GIT_SETUP.md](docs/GIT_SETUP.md) 安裝 Git for Windows 並將 `git` 加入 PATH。  
在專案目錄執行 `git init` 後，連接 GitHub：在 GitHub 建立 repo，再執行  
`git remote add origin https://github.com/你的帳號/Weekly_Report_Checker.git`、  
`git branch -M main`、  
`git push -u origin main`。  
詳細步驟見 [docs/GIT_SETUP.md](docs/GIT_SETUP.md)。

## 注意事項

若 OneNote connector 無法使用，請參考 `power_automate/FLOW_ACTIONS_README.md` 的替代方案。
