# Power Automate Flow：每個 Action 說明

本流程目的：手動觸發後，從人名清單讀取要檢查的 OneNote section，透過 OneNote (Business) 取得各 section 下的頁面（日期），組裝成 JSON 並存到 OneDrive Business，供後續 Python 程式處理。

---

## 步驟 1：手動觸發 (Manual trigger)

| 項目 | 說明 |
|------|------|
| **Connector** | 內建「手動觸發流程」 |
| **用途** | 由使用者手動執行此 flow，不需排程或事件。 |
| **設定** | 無需額外參數。 |
| **輸出** | 無動態內容；僅作為流程起點。 |

**在 Power Automate 中的操作**：新增流程 → 從空白開始 → 選擇「手動觸發流程」。

---

## 步驟 2：取得檔案內容 (OneDrive - Get file content)

| 項目 | 說明 |
|------|------|
| **Connector** | OneDrive for Business（或 SharePoint） |
| **Action** | **取得檔案內容** (Get file content) |
| **用途** | 從 OneDrive 讀取人名清單文字檔，供後續剖析。 |
| **輸入** | **檔案識別碼**：可選「從 OneDrive 挑選檔案」指定 `member_list.txt` 的路徑，或使用動態內容「檔案識別碼」。建議將 `member_list.txt` 上傳到 OneDrive 固定資料夾（例如與 JSON 輸出同資料夾），在此選取該檔案。 |
| **輸出** | `body`（檔案內容，文字）、`$content` 等。後續步驟使用「檔案內容」或 `body`。 |

**注意**：若清單改為放在 SharePoint 文件庫，可改用「SharePoint - 取得檔案內容」，並指定網站與檔案路徑。

---

## 步驟 3：剖析人名清單 (Compose 或 變數 + Split)

| 項目 | 說明 |
|------|------|
| **方式 A** | 使用 **資料作業 → 撰寫** (Compose)，在「輸入」中用運算式：<br/>`split(body('Get_file_content'), '\n')`<br/>若檔案為 UTF-8 且含 BOM，可先以 `replace(body('Get_file_content'), '﻿', '')` 再 `split`。 |
| **方式 B** | 新增 **變數**（類型：陣列），再以 **套用至每個** 搭配 **分割** 將每行加入陣列；或使用 **分割** 動作（若環境有提供）將文字依換行符號分割。 |
| **用途** | 將「一人一行」的文字檔轉成陣列，例如 `["Paddy","Redmin","姍蓉"]`。 |
| **輸出** | 陣列，每個元素為一個人名。下一步「套用至每個人」會反覆此陣列。 |

**運算式範例**（在撰寫的輸入中）：
```text
split(replace(body('Get_file_content'), '﻿', ''), '\n')
```
可再以 `select` 或後續 **篩選** 去掉空字串。

---

## 步驟 4：套用至每個人 (Apply to each)

| 項目 | 說明 |
|------|------|
| **Connector** | 內建 **套用至每個** (Apply to each) |
| **選取輸出** | 上一步的陣列（例如 `outputs('Compose')` 或變數）。 |
| **用途** | 對清單中的每個人名執行一輪「取得 OneNote 頁面 → 組裝一筆資料」。 |
| **內部** | 在迴圈內放置步驟 5、6（取得頁面、組裝該人名的 JSON 節點）。 |

---

## 步驟 5：列出 Section / 頁面 (OneNote connector)

| 項目 | 說明 |
|------|------|
| **Connector** | **OneNote (Business)** 或 **OneNote**（依您的 Power Automate 環境顯示名稱）。 |
| **Action** | **列出筆記本中的區段** 或 **列出區段中的頁面**。 |
| **實際流程** | 若 connector 提供「依 section 名稱取得頁面」：<br/>- 使用 **列出區段** 取得 section 清單，再比對目前迴圈的人名；或<br/>- 使用 **取得區段中的頁面**，並在「區段識別碼」中帶入對應的 section（若前一步有「列出區段」可取得 section ID）。 |
| **輸入** | 筆記本識別碼、section 識別碼（或 section 名稱，依 connector 支援而定）。目前迴圈的人名來自 **套用至每個** 的 `currentItem`（若為字串）或 `currentItem?['value']`（若為物件）。 |
| **輸出** | 頁面清單，每筆通常含頁面名稱（即日期，如 2026/03/05）、頁面 ID 等。 |

**注意**：若貴公司 Microsoft Graph 被阻擋，且 OneNote (Business) connector 底層仍使用 Graph，此步驟可能失敗。替代方案：改由 Power Automate 僅產出「人名清單」的 JSON，不呼叫 OneNote；再由您以本機 OneNote 或其他允許的方式取得頁面清單，手動合併成同一 JSON 格式後交給 Python。

---

## 步驟 6：組裝 JSON (Compose 或 加入陣列)

| 項目 | 說明 |
|------|------|
| **用途** | 將「目前迴圈的人名」與「該 section 下的頁面名稱清單」組成一筆物件，例如 `{ "同仁名": "Paddy", "頁面日期": ["2026/03/05","2026/02/25",...] }`。 |
| **實作** | 在 **套用至每個** 內使用 **撰寫**，輸入為 JSON 物件，例如：<br/>`{ "同仁名": items('Apply_to_each'), "頁面日期": body('List_pages_in_section')?['value'] }` 或依您實際動作輸出調整（例如從「列出頁面」的結果用 **選取** 只取「標題」或「名稱」陣列）。 |
| **迴圈外** | 在 **套用至每個** 之後，再一個 **撰寫** 或 **變數**，將迴圈內每一輪的輸出 **加入陣列**，得到全員的陣列。最終格式建議為：<br/>`{ "members": [ { "同仁名": "Paddy", "頁面日期": ["2026/03/05", ...] }, ... ] }` 或 `{ "Paddy": ["2026/03/05", ...], "Redmin": [...], ... }`，以配合 Python 程式預期的結構。 |

Python 程式預期 JSON 結構範例（依 `weekly_report_checker.py` 實作）：
- 格式一：`{ "同仁名": [ "日期1", "日期2", ... ], ... }`
- 格式二：`{ "members": [ { "同仁名": "名字", "頁面日期": ["日期1", ...] }, ... ] }`  
程式需能解析並對應到「同仁名 → 該員的頁面日期陣列」。

---

## 步驟 7：建立檔案 (OneDrive - Create file)

| 項目 | 說明 |
|------|------|
| **Connector** | OneDrive for Business |
| **Action** | **建立檔案** (Create file) |
| **輸入** | **資料夾路徑**：選擇要存放 JSON 的 OneDrive 資料夾（例如 `/WeeklyReport/` 或您指定的路徑）。<br/>**檔案名稱**：建議含時間戳記，例如：<br/>`concat('onenote_pages_', formatDateTime(utcNow(), 'yyyyMMdd_HHmmss'), '.json')`<br/>**檔案內容**：上一步組裝好的完整 JSON 字串。若上一步為物件，可用 `string( body('Compose_final') )` 或 `string(variables('jsonOutput'))` 轉成字串。 |
| **用途** | 將 JSON 寫入 OneDrive Business，供使用者下載或透過 OneDrive 同步到本機，再由 Python 讀取。 |

---

## 流程摘要圖（概念）

```text
手動觸發
    → 取得 member_list.txt 檔案內容 (OneDrive)
    → 剖析人名清單 (Split/Compose)
    → 套用至每個 [人名]
        → 列出該 section 的頁面 (OneNote)
        → 組裝該人名的 JSON 節點 (Compose)
    → 組裝完整 JSON 並加入陣列
    → 建立檔案 (OneDrive)：寫入 JSON，檔名含時間戳記
```

---

## 替代方案（無法使用 OneNote connector 時）

1. **僅產出人名清單 JSON**：Power Automate 只讀取 `member_list.txt`，組裝成 `{ "members": ["Paddy","Redmin","姍蓉"] }` 存到 OneDrive。Python 端僅根據「需檢查名單」產出 Excel，頁面日期改由您手動維護或從其他允許的來源滙入。
2. **手動合併**：您用本機 OneNote 或其他工具匯出各 section 的頁面日期，存成 JSON 後與 Power Automate 產出的檔案合併，再交給 Python 執行比對與產出 Excel。

匯入本資料夾中的 `flow_export.zip` 後，可依實際 connector 名稱與輸出結構，對照上述說明微調每個步驟的欄位對應。
