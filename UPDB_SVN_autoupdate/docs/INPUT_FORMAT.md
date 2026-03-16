# 輸入文字檔格式

程式從單一文字檔讀取要加入的專案與同仁，格式如下。

## 欄位說明

每行一筆，以 **Tab** 區隔「第一欄」與「第二欄」；第二欄起為工號與其他不需程式使用的欄位。

| 位置 | 內容 | 程式是否使用 |
|------|------|--------------|
| 第一欄（Tab 前） | 以空白分隔：`專案名 群組名 [SVN] RT` | 是：專案、群組、是否加 SVN |
| 第二欄（Tab 後） | 工號（如 R8943） | 是：UPDB-manager 與 SVN 皆用工號加入 |
| 其餘欄位 | 姓名、分機、部門、Email 等 | 否 |

- **專案名**：例如 `RL6665`。
- **群組名**：表示要加入該專案下的哪個群組，**必須**為下列其一（與 UPDB 畫面「Add XXX..」對應）：
- **RT**：忽略。

### 群組名（目前支援選項）

| 輸入的群組名 | 說明／對應 UPDB 按鈕 |
|-------------|----------------------|
| `Analog`    | Add Analog..         |
| `Digital`   | Add Digital..        |
| `DV`        | Add DV..             |
| `Layout`    | Add Layout..         |
| `APR`       | Add APR..            |
| `CTC`       | Add CTC DTD..        |
| `CTC DTD`   | Add CTC DTD..（同上） |
| `Planner`   | Add Planner..        |
| `Testing`   | Add Testing..        |

輸入檔第一欄的群組名若不在上表，程式會記錄「不支援的群組」並略過該筆。
- **工號**：同仁工號，UPDB-manager 與 SVN 都以此工號加入。

## 範例

```
RL6665 Digital SVN RT	R8943	胡定安	3519497	通訊網路事業群	 andy.hu@realtek.com
RL6665 Analog SVN RT	R8067	王泓閔	3510572	通訊網路事業群	 redmin.wang@realtek.com
```

- 第一行：專案 `RL6665`、群組 `Digital`、要加 SVN，工號 `R8943`。
- 第二行：專案 `RL6665`、群組 `Analog`、要加 SVN，工號 `R8067`。

## 編碼

建議文字檔使用 **UTF-8** 編碼，程式讀取時明確指定 `encoding='utf-8'`，避免中文欄位亂碼。
