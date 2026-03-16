# 分享給其他同事使用

讓同事在本機直接執行 UPDB 批次加人，無需開 PowerShell，照下列步驟即可。

## 1. 取得程式

任選一種方式：

- **從 Git 複製**（若同事有權限）  
  在要放的資料夾執行：  
  `git clone https://github.com/a5512561-creator/AutomationAndAI.git`  
  程式在 `AutomationAndAI\UPDB_SVN_autoupdate` 資料夾內。

- **壓縮檔**  
  你把整個 `UPDB_SVN_autoupdate` 資料夾壓成 zip 傳給同事，同事解壓到任意路徑（例如 `D:\UPDB_SVN_autoupdate`）。

## 2. 安裝環境（僅第一次）

1. 安裝 **Python 3.10 以上**（從 [python.org](https://www.python.org/downloads/)），安裝時勾選「Add Python to PATH」。
2. 開啟命令提示字元（cmd）或 PowerShell，進入程式資料夾：
   ```bat
   cd /d D:\UPDB_SVN_autoupdate
   ```
3. 安裝依賴與瀏覽器：
   ```bat
   pip install -r requirements.txt
   playwright install chromium
   ```

## 3. 設定與輸入檔（可選）

- **設定檔**：複製 `config.example.yaml` 為 `config.yaml`，若有不同 UPDB 網址或等待秒數再改。
- **成員名單**：依 `docs/INPUT_FORMAT.md` 與 `members.example.txt` 格式準備一個文字檔（Tab 分隔專案、群組、工號等）。

## 4. 執行方式

**方式 A：雙擊 .bat（建議）**

- 雙擊 `run_updb.bat`：會用範例檔 `members.example.txt` 執行（可先試跑）。
- 或把成員名單檔拖到 `run_updb.bat` 上，會用該檔執行。
- 或在資料夾路徑列輸入 `cmd`  Enter，再輸入：  
  `run_updb.bat 你的成員名單.txt`

**方式 B：命令列**

```bat
cd /d D:\UPDB_SVN_autoupdate
python main.py 你的成員名單.txt
```

執行後會開瀏覽器，同事在網頁完成 OTP 登入，回到黑視窗按 Enter，程式會依名單批次加人；log 會寫入 `logs\updb_batch.log`。

## 5. 注意事項

- 需能連上公司 UPDB 網址（如 `https://project.rd.realtek.com/...`）。
- 登入與 OTP 需同事本人操作，程式只做登入後的頁面操作。
- 若出現「重新導向次數過多」，請用 `run_updb.bat --clear-cookies 你的名單.txt` 或參考 README 疑難排解。
