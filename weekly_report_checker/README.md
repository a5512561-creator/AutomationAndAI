# Weekly Report Checker

從 OneNote 讀取週報頁面清單與內容，以 Python 分析填寫狀況、比對最新兩週內容，產出檢查結果 Excel。

## 專案結構

- `config/member_list.txt`：需檢查週報的同仁名單（一人一行）
- `config/holidays.txt`：長假日期清單（選配，用於跳週容忍判定）
- `src/Get-OneNotePages.ps1`：透過 OneNote COM 介面讀取頁面清單與最新 2 頁內容，產出 JSON
- `src/fetch_onenote_pages.py`：Python wrapper，呼叫上述 PowerShell 腳本
- `src/fix_onenote_typelib.ps1`：首次設定用，修正 OneNote TypeLib 路徑
- `src/onenote_parser.py`：解析 OneNote XML 表格，提取結構化資料
- `src/weekly_report_checker.py`：讀取 JSON，分析填寫狀況與內容，產出 Excel
- `output/`：JSON 與 Excel 輸出目錄

## 前提條件

- Python 3.9+
- **OneNote 桌面版**（2016 / 2021 / Microsoft 365），非 Windows 10 UWP 版
- 目標 Notebook 已在 OneNote 桌面應用程式中**開啟並同步完成**
- PowerShell 5.1+（Windows 內建）

```powershell
pip install -r requirements.txt
```

### 首次設定（只需一次）

OneNote Click-to-Run 有 TypeLib 路徑 Bug，首次需執行修正：

```powershell
powershell -ExecutionPolicy Bypass -File src/fix_onenote_typelib.ps1
```

## 使用方式

### 步驟 1：從 OneNote 讀取頁面資料

```powershell
cd D:\AutomationAndAI\weekly_report_checker
powershell -ExecutionPolicy Bypass -File src/Get-OneNotePages.ps1
```

或透過 Python wrapper：

```powershell
python src/fetch_onenote_pages.py
```

程式透過 OneNote COM 介面讀取 `config/member_list.txt` 中每位同仁對應的 Section 頁面清單，並自動擷取最新 2 頁的完整內容（XML），產出 JSON 到 `output/`。

**參數：**

| 參數 | 說明 | 預設值 |
|------|------|--------|
| `-NotebookName "名稱"` | 指定 Notebook 名稱 | `Switch-DD member weekly` |
| `-MemberListPath 路徑` | 指定成員清單檔案 | `config/member_list.txt` |
| `-OutputPath 路徑` | 指定 JSON 輸出路徑 | `output/onenote_pages_*.json` |
| `-ContentPages N` | 每位成員讀取內容的頁數 | `2` |

### 步驟 2：執行週報檢查

```powershell
python src/weekly_report_checker.py
```

- **不帶參數**：自動尋找 `output/` 下最新的 `onenote_pages*.json`
- **指定檔案**：`python src/weekly_report_checker.py "output\onenote_pages_20260311_162151.json"`
- **指定輸出目錄**：`python src/weekly_report_checker.py -o D:\Reports`

### 步驟 3：查看結果

- **檔名**：`WeeklyReport_Check_YYYYMMDD_HHMMSS.xlsx`
- **位置**：`output/` 目錄

#### Sheet 1：週報檢查

| 欄位 | 說明 |
|------|------|
| 同仁名 | OneNote section 名稱（人名） |
| 本週狀態 | 已填寫 / 未填寫 / 遲交 |
| 連續未填週數 | 從本週往回算 |
| 最近 12 週填寫率 | 百分比與分數 |
| 最後填寫日期 | 最近一筆週報日期 |
| 備註 | 綜合說明 |

#### Sheet 2：內容分析

針對每位同仁最新 2 頁週報內容進行深度比對，包含：

| 檢查規則 | 說明 |
|----------|------|
| 規則 1 | 最新 2 頁日期間隔是否為 7 天 |
| 規則 2 | 是否在該週四前完成更新 |
| 規則 3-4 | 跳週容忍（搭配 `config/holidays.txt` 長假排除） |
| 規則 5 | 請假容忍（連續 1-2 週無頁面標記「可能請假」） |
| 規則 6 | Item 進度比對：Prgs% 是否增加、BW% 非零、週報有內容 |
| 規則 7 | BW spent% > 0 但進度無變化且無週報內容 |
| 規則 8 | Prgs 100% 持續出現，建議移除已完成任務 |
| 規則 9 | Priority 1 任務本週 BW=0% 或無週報內容 |
| 格式檢查 | Header row 是否與標準 7 欄結構一致 |

每位同仁的 Item 逐項列出比對結果，異常項目以紅底標示，底部彙總各規則違規數量。

### 一行搞定（兩步合一）

```powershell
powershell -ExecutionPolicy Bypass -File src/Get-OneNotePages.ps1; python src/weekly_report_checker.py
```

## 疑難排解

| 問題 | 解法 |
|------|------|
| `GetHierarchy returned too little data` | 確認 OneNote 桌面版已開啟目標 Notebook 並完成同步 |
| `Notebook not found` | 確認 Notebook 名稱正確，可查看 PowerShell 輸出的可用 Notebook 列表 |
| `Section not found` | 確認 Section 名稱與 `member_list.txt` 完全一致 |
| `TYPE_E_CANTLOADLIBRARY` | 執行 `src/fix_onenote_typelib.ps1` 修正 TypeLib 路徑 |

## Git 與 GitHub

若尚未安裝 Git，請先依 [docs/GIT_SETUP.md](docs/GIT_SETUP.md) 安裝 Git for Windows 並將 `git` 加入 PATH。  
詳細步驟見 [docs/GIT_SETUP.md](docs/GIT_SETUP.md)。
