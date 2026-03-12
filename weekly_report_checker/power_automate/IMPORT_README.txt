Power Automate 流程匯入說明
============================

本資料夾內的 definition.json 為流程邏輯的參考結構（Logic Apps 相容格式）。
flow_export.zip 內含此定義，供您嘗試匯入。

若在 Power Automate 中「匯入」→「匯入套件」無法成功匯入此 zip：
1. 請改依 power_automate/FLOW_ACTIONS_README.md 的步驟，在 Power Automate 中手動建立流程。
2. 每個 action 的設定與欄位對應皆已寫在 FLOW_ACTIONS_README.md 中。

匯入後請記得：
- 設定 OneDrive 與 OneNote 的連線。
- 將參數 MemberListPath 設為 member_list.txt 在 OneDrive 的檔案識別碼或路徑。
- 將參數 OutputFolderPath 設為要存放 JSON 的 OneDrive 資料夾路徑。
