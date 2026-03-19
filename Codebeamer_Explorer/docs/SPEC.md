# Codebeamer Explorer 規格書

## 1. 目的

由 Word `.docx` 規格書自動建立 Codebeamer tracker 的樹狀項目，並把章節內的圖片寫入每個項目的 **Description**。

主要行為：

- 根節點：由 `.docx` 檔名擷取 `Hardware Component name`，Category = `Hardware Component`
- 子節點：依 Word 的標題/編號層級建立（例如 `1`, `1.1`, `2`, `2.1`…），Category = `Information`
- `2.1` 節內的 Hardware Part：掃描該節表格文字，擷取 `HWP_\\d+`（例如 `HWP_1`），Category = `Hardware Part`
- 圖片：會先上傳到「host item」取得 `artifact_id`，再用 Codebeamer `Image` macro 產生可顯示的圖片內容

## 2. 環境與依賴

- Python 3.x
- 套件見 `requirements.txt`：`requests`、`python-dotenv`

安裝：

```bash
pip install -r requirements.txt
```

## 3. `.env` 設定

所有參數由專案目錄的 `.env` 讀取。

### 3.1 連線（必填）

- `CB_BASE_URL`：含 `/cb/rest/v3` 的 REST 基底 URL（例如 `https://almqa.realtek.com/cb/rest/v3`）
- `CB_TRACKER_ID`：目標 tracker ID
- 認證（二擇一）
  - `CB_USERNAME` / `CB_PASSWORD`
  - 或 `CB_TOKEN`

### 3.2 DOCX 來源（必填）

- `CB_DOCX_PATH`：要解析的 docx 完整路徑（建議用絕對路徑）

### 3.3 圖片 host item（必填/建議）

程式會把圖片先上傳到固定的 host item，才能在 `Image` macro 內使用 `artifact_id`。

- 建議：`CB_ATTACHMENT_HOST_ITEM_ID`（例如 `7232`）
- 若未設：程式會改用 `CB_TEST_ITEM_ID`

因此至少要提供 `CB_TEST_ITEM_ID`。

## 4. 執行方式

此專案現階段只有一個主程式：

```bash
python script/create_cb_items_from_docx.py
```

### 4.1 只解析、不呼叫 API（建議先跑）

```bash
python script/create_cb_items_from_docx.py --dry-run
```

### 4.2 實際建立（需要避免重複建立）

```bash
python script/create_cb_items_from_docx.py --apply --force
```

### 4.3 重要參數

- `--debug-images`：印出從 docx 抽取的圖片統計（用來確認圖片是否真的被解析到對應章節）
- `--no-images`：不處理 docx 圖片（只建立文字樹狀結構）
- `--debug-docx`：除錯：額外列出 `2.1` 區段 table 的前幾個 cell 文字（用來確認 `HWP_#` 抽取）
- `--no-reindent`：建立完成後不做 `children INSERT` 重排（預設會做）

## 5. Description（文字 + 圖片）寫入規則

在此環境中，Description 欄位用 API PUT 可能會回 `403 not writable`，因此程式採用以下策略：

1. 圖片先上傳到 host item（取得 `artifact_id`）
2. Description 內容於 **POST 建立 item 時一次帶入**
3. 若某節點圖片上傳未成功取得完整 `artifact_id`，該節點會「只寫文字，不寫圖片」

圖片輸出使用：

- `[{Image wiki='[CB:/displayDocument/{filename}?task_id={host_item_id}&artifact_id={att_id}]' width='600' height='400'}]`

並在章節文字存在時：

- 章節文字 + `\n\n` + Image macro 一起組成最终 Description

## 6. 常見驗證方式

### 6.1 確認 docx 有抽到圖片

執行：

```bash
python script/create_cb_items_from_docx.py --dry-run --debug-images
```

看輸出中的 `images extracted` 是否符合你預期的章節編號。

### 6.2 確認圖片是否被寫入 Description

執行建立：

```bash
python script/create_cb_items_from_docx.py --apply --force --debug-images
```

程式會在建立後進行 GET 驗證，檢查 Description 內是否含 `displayDocument`（image macro 的核心片段）。

