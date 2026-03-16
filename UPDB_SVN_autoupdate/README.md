# UPDB_SVN_autoupdate

依文字檔批次將同仁加入 UPDB-manager 專案與群組。

## 文件索引

- **[輸入文字檔格式](docs/INPUT_FORMAT.md)**：來源檔欄位說明與範例、支援的群組名。
- **[輸入檔處理流程與每專案分次設定](docs/FLOW_AND_BATCHING.md)**：解析與分組方式、執行順序與「每專案等儲存完成再處理下一個」的說明。
- **[分享給其他同事使用](docs/SHARING.md)**：同事如何取得程式、安裝環境、用 .bat 或指令執行。

## 需求摘要

- **輸入**：文字檔，每行指定「專案、群組、是否加 SVN、工號」。
- **行為**：以瀏覽器自動化操作 UPDB-manager（登入時由使用者手動完成 OTP），依序選專案、加群組成員。
- **技術**：Python + Playwright。

## 使用方式

1. **安裝依賴**（僅第一次）
   ```bash
   pip install -r requirements.txt
   playwright install chromium
   ```
2. **設定**（可選）  
   複製 `config.example.yaml` 為 `config.yaml`，設定 `updb_login_url`、`input_file`、`member_add_wait_seconds` 等。
3. **執行**
   - **一鍵執行（建議）**：雙擊 `run_updb.bat`（使用範例名單）或把成員名單檔拖到 `run_updb.bat` 上。
   - **命令列**：`python main.py members.txt` 或 `python main.py`（須在 config 設定 `input_file`）。  
   程式會開啟 Chromium、載入登入頁；請在瀏覽器完成兩階段 OTP 登入後，回到終端按 **Enter**，程式會依文字檔批次執行 UPDB 加成員。
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

頁面操作依 `selectors.yaml` 的連結/按鈕文字（如「變更成員名單」「Add Digital..」「儲存變更」等）進行。若 UPDB-manager 改版導致點擊失敗，可依實際 HTML 調整 `selectors.yaml` 或 `browser_ops.py`。建議以單筆或少量資料先做一次端對端測試確認流程。
