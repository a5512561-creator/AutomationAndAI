# Codebeamer Explorer 規格書

## 1. 目的

從 Codebeamer REST API 讀取指定 tracker 的 work items，並依設定的「項目名稱」篩選、列出第一層或展開子層，供後續整合或檢視規格結構使用。

## 2. 環境與依賴

- Python 3.x
- 依賴套件見 `requirements.txt`：`requests`、`python-dotenv`

安裝：

```bash
pip install -r requirements.txt
```

## 3. 設定檔 .env

所有參數由專案根目錄或執行目錄下的 `.env` 讀取。

- **第一次使用**：複製 `.env.example` 為 `.env`，再填入實際的連線與 tracker 等參數。
- **版控**：`.env` 內含帳號密碼，請勿提交到 Git；專案已提供 `.env.example`（僅範例與欄位說明、無真實帳密）供他人複製使用。

### 3.1 連線（必填）

| 參數 | 說明 | 範例 |
|------|------|------|
| `CB_BASE_URL` | Codebeamer REST API 基底網址（含 `/cb/rest/v3`） | `https://almqa.realtek.com/cb/rest/v3` |
| `CB_USERNAME` | 登入帳號（與 `CB_PASSWORD` 二擇一，或改用 `CB_TOKEN`） | `your_username` |
| `CB_PASSWORD` | 登入密碼 | `your_password` |
| `CB_TOKEN` | API Token（若使用則註解掉 `CB_USERNAME` / `CB_PASSWORD`） | 選填 |

### 3.2 Tracker 與分頁

| 參數 | 說明 | 範例 |
|------|------|------|
| `CB_TRACKER_ID` | 目標 tracker ID（從專案／tracker 網址取得） | `12096` |
| `CB_PAGE_SIZE` | 掃描 tracker 時每頁筆數 | `100` |
| `CB_TEST_ITEM_ID` | 單一 item ID，程式開頭會先讀此筆並印出 JSON 摘要（驗證連線與結構） | `7210` |
| `CB_TRACKER_ITEMS_URL` | （可選）覆蓋 tracker items API 完整 URL，通常留空 | 留空 |

### 3.3 列出項目：要抓哪些、要不要展開子層

| 參數 | 說明 | 範例 |
|------|------|------|
| `CB_LIST_FIRST_LEVEL_ONLY` | `1` = 只列出符合名稱的項目 id/name，不抓 children；`0` 或未設 = 會再抓每個符合項目的 children 並列出 | `1` |
| `CB_TARGET_COMPONENT_NAMES` | 要比對的「項目名稱」清單，對應 API 的 `name` 欄位；多個用分號 `;` 分隔。留空則使用程式內建預設（已統一為 formal）；填了則只找這些名稱（須與 API 回傳的 name 完全一致，不要加 `id=xxx`） | `HwCom_81; [SWITCH] top view` |
| `CB_LIST_FIRST_N` | （除錯用）設為正整數時，只印出 tracker 前 N 筆 item 的 id/name，方便複製到 `CB_TARGET_COMPONENT_NAMES` | `50` |

## 4. 執行方式

在專案根目錄或 `script` 目錄下執行：

```bash
python script/read_cb_hw2_titles.py
```

程式會：

1. 讀取 `CB_TEST_ITEM_ID` 對應的單一 item 並印出 JSON 摘要（確認連線與結構）。
2. 依目前模式掃描 tracker、比對 `CB_TARGET_COMPONENT_NAMES`（或程式內建預設），並列出符合項目的 id/name；若 `CB_LIST_FIRST_LEVEL_ONLY=0`，會再抓每個符合項目的 children 並列出。

## 4.1 由 DOCX 新增項目（測試程式）

此功能用於驗證「由 Word 規格書建立 tracker 樹狀項目」的流程。

### 需要的 .env 參數

- `CB_BASE_URL`, `CB_TRACKER_ID`, `CB_USERNAME/CB_PASSWORD`（或 `CB_TOKEN`）
- `CB_DOCX_PATH`：要讀取的 docx 完整路徑（建議用絕對路徑）

### 行為規則（對應你的測試流程）

- **Hardware Component（父項）**：由檔名擷取，例如 `..._PaddyTest.docx` ⇒ `PaddyTest`，Category = `Hardware Component`
- **章節節點**：依 Word 標題編號（如 `1`, `1.1`, `2`, `2.1`, `3`, `3.1`）建立節點，Category = `Information`
- **2.1 Hardware Part**：在 `2.1` 區段內掃描表格文字，擷取 `HWP_\\d+`（如 `HWP_1`, `HWP_2`），建立於 `2.1` 節點底下，Category = `Hardware Part`

### 執行方式

- 只解析、不呼叫 API（建議先跑確認解析正確）：

```bash
python script/create_cb_items_from_docx.py --dry-run
```

- 實際建立（會呼叫 POST/PUT 變更 Codebeamer；為避免誤建，需加 `--force`）：

```bash
python script/create_cb_items_from_docx.py --apply --force
```

預設會在建立完成後再用 children API 強制重排縮排/順序（避免某些 tracker UI 不依 parent 顯示）。若你確定不需要重排，可加：

```bash
python script/create_cb_items_from_docx.py --apply --force --no-reindent
```

## 4.2 依 Word 編號重排縮排（既有項目）

若你已建立項目但 UI 左側樹狀縮排不符合 Word 編號（例如 `1.1` 沒縮在 `1` 底下），可使用 children API 重新掛載既有節點。

此腳本會呼叫：

- `PATCH /v3/items/{parentId}/children?mode=INSERT`

### 執行方式

- 先看將執行的搬移計畫（不會改動）：

```bash
python script/reindent_cb_items.py
```

- 真的套用重排（會改動 Codebeamer）：

```bash
python script/reindent_cb_items.py --apply
```

## 5. 行為模式（目前兩種）

- **只列第一層（不展開）**  
  `CB_LIST_FIRST_LEVEL_ONLY=1`：掃描整個 tracker 的 itemRefs，以「名稱完全一致」比對 `CB_TARGET_COMPONENT_NAMES`（或程式內建預設），只輸出符合項目的 id/name，不呼叫各 item 的 children。

- **列出並展開 children**  
  `CB_LIST_FIRST_LEVEL_ONLY=0` 或未設：同上比對後，對每個符合的 item 再呼叫 API 取得其 children，並列出每個子項目的 id/name。

若某個名稱在 tracker 中找不到，程式會輸出 `[MISS]` 並嘗試印出「建議 API 名稱」供複製到 `.env` 修正。

## 6. 檔案結構

```
Codebeamer_Explorer/
├── .env                 # 本機設定（勿提交版控）
├── .env.example         # 範例設定檔（無帳密）
├── docs/
│   └── SPEC.md         # 本規格書
├── requirements.txt
└── script/
    └── read_cb_hw2_titles.py
```

## 7. 注意事項

- 名稱比對為**完全一致**（含括號、空格、大小寫），與 API 回傳的 `name` 欄位必須一字不差。
- 若名稱常變動或需探索正確名稱，可先設 `CB_LIST_FIRST_N=50` 取得前 50 筆 id/name，再複製到 `CB_TARGET_COMPONENT_NAMES`。
