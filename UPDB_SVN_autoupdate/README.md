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
- **技術**：Python + Playwright，詳見 `docs/PLAN.md`。

## 使用方式

1. **安裝依賴**
   ```bash
   pip install -r requirements.txt
   playwright install chromium
   ```
2. **設定**（可選）  
   複製 `config.example.yaml` 為 `config.yaml`，設定 `updb_login_url`、`input_file`、`member_add_wait_seconds` 等。
3. **執行**
   ```bash
   python main.py  members.txt
   ```
   或僅指定 config 內 `input_file` 時：`python main.py`（須在 config 中設定 `input_file`）。  
   程式會開啟 Chromium、載入登入頁；請在瀏覽器完成兩階段 OTP 登入後，回到終端按 **Enter**，程式會依文字檔批次執行 UPDB 加成員與 SVN 權限。
4. **選項**
   - `--config -c`：指定設定檔。
   - `--no-continue`：遇錯誤即中斷，不繼續下一筆。
   - `--log-dir`：日誌目錄（預設 `logs`）。
   - `--clear-cookies`：啟動前清除 Cookie，可避免登入後出現「重新導向次數過多」。

## 疑難排解：重新導向次數過多（ERR_TOO_MANY_REDIRECTS）

若完成 OTP 登入後出現「這個網頁無法正常運作 / 重新導向的次數過多」：

1. **盡快按 Enter**：程式會在您按 Enter 後立即導向專案 UPDB 頁，以避開會觸發迴圈的 main_page.php。
2. **使用全新 session**：加上參數 `--clear-cookies` 再執行，例如  
   `python main.py --clear-cookies members.txt`，然後重新登入 OTP。
3. **手動刪除 Cookie**：若仍失敗，可在瀏覽器針對 `project.rd.realtek.com` 刪除 Cookie 後重試。

## 選擇器與 E2E 測試

頁面操作依 `selectors.yaml` 的連結/按鈕文字（如「變更成員名單」「Add Digital..」「SVN」「編輯權限」等）進行。若 UPDB-manager 或 SVN 頁改版導致點擊失敗，可依實際 HTML 調整 `selectors.yaml` 或 `browser_ops.py` 內對應選擇器。建議以單筆或少量資料先做一次端對端測試確認流程。
