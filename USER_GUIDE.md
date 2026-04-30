 # BOM Excel Tool 使用者文件

## 1. 工具用途

`bom_excel_tool.py` 用來整理 BOM 類 Excel（主料來源 `BOM COST`、子料來源 `EMS BOM COST`）：

- 自動找到第一欄為 `Level` 的標題中心列
- 將「上一列 + 中心列 + 下一列」合併成欄位名稱
- 讀出後續資料列
- 新增 `Board` 欄位：當 `Level = 1` 時，取該列 `Description`，並向下填入直到下一個 `Level = 1` 覆蓋
- 新增 `M/S` 欄位（固定填 `M`）與 `Main Source` 欄位（填入 `Item`）
- 將 `2ND_SOURCE_TOTAL` 起的尾端欄位依規則重命名為 `Sub_*`
- 依子料來源的 `Sub_*` 展開產生子料列（`M/S = S`），插入在主料列後方

## 2. 環境需求

- Python 3.10 以上（建議）
- 套件：
  - `openpyxl>=3.1.0`
  - `pandas>=2.0.0`

安裝：

```bash
pip install -r requirements.txt
```

## 3. 基本用法

在專案目錄執行：

```bash
python bom_excel_tool.py "你的Excel檔案.xlsx"
```

每次執行會輸出兩份檔案：

- 完整欄位檔（所有欄位）
- 指定欄位檔（固定欄位順序）

若未指定輸出檔，完整欄位檔預設為：

- 與原檔同目錄
- 檔名為 `原檔名_flat.csv`
- 編碼為 `utf-8-sig`（Excel 開啟中文較友善）

指定欄位檔會在完整欄位檔名後加上 `_selected`，例如：

- 完整：`result.csv`
- 指定欄位：`result_selected.csv`

## 4. 命令列參數

```bash
python bom_excel_tool.py INPUT [-o OUTPUT] [--sheet SHEET] [--marker MARKER] [--sep SEP] [--preview] [--sub-source SUB_SOURCE] [--sub-sheet SUB_SHEET] [--selected-only | --no-selected]
```

- `INPUT`：輸入 `.xlsx` 路徑（必要）
- `-o, --output`：輸出檔（支援 `.csv`, `.xlsx`, `.xlsm`）
- `--sheet`：指定工作表名稱（未指定時會掃描並合併所有可用工作表）
- `--marker`：標題中心列第一欄文字（預設 `Level`）
- `--sep`：三列表頭合併分隔字元（預設空白字元 ` `）
- `--preview`：僅預覽，不寫檔
- `--sub-source`：子料來源檔（選填；預設解讀為 `EMS BOM COST`；提供 `Sub_*` 的幣別/價格）
  - 未提供時會使用 `BOM COST` 內建的 `Sub_*` 展開
  - 相容舊參數：`--ecode-source` 仍可用，但等同 `--sub-source`
- `--sub-sheet`：子料來源檔工作表（指定單一工作表；未指定時會掃描所有工作表）
- `--selected-only`：只輸出指定欄位檔（`*_selected`）
- `--no-selected`：只輸出完整欄位檔

## 5. 常見範例

### 5.1 只做預覽（不輸出）

```bash
python bom_excel_tool.py "data/BOM COST.xlsx" --preview
```

### 5.2 輸出 CSV

```bash
python bom_excel_tool.py "data/BOM COST.xlsx" -o "output/result.csv"
```

### 5.3 輸出 Excel

```bash
python bom_excel_tool.py "data/BOM COST.xlsx" -o "output/result.xlsx"
```

### 5.4 指定工作表與分隔符號

```bash
python bom_excel_tool.py "data/sample.xlsx" --sheet "91-017-507025B.401" --sep " - "
```

### 5.4-2 主檔多工作表合併（預設）

```bash
python bom_excel_tool.py "data/sample.xlsx" -o "output/result.csv"
```

未指定 `--sheet` 時，主檔（`INPUT`）會自動掃描並合併所有含 `Level` 標題的工作表。

### 5.5 （選填）指定子料來源（BOM COST + EMS BOM COST）

```bash
python bom_excel_tool.py "data/BOM COST.xlsx" --sub-source "data/EMS BOM COST.xlsx" -o "output/result.csv"
```

說明：
- `INPUT(BOM COST)` 會直接產生主料列（`M/S = M`），並從 `BOM COST` 取得 `Ecode`。
- `--sub-source(EMS BOM COST)` 提供 `Sub_n` 的子料列（`M/S = S`），其 `Last BPA Currency/Last BPA Price` 來自 `Sub_n_Currency/Sub_n_Price`。

### 5.6 展開 Sub 為子料列

輸出時會自動將每列 `Sub_n` 展開為新列，並插在原主料列後方：

- 子料 `Item` = `Sub_n`
- 子料 `M/S` = `S`
- 子料 `Last BPA Currency` = `Sub_n_Currency`
- 子料 `Last BPA Price` = `Sub_n_Price`
- 子料只填入主料對應欄位：`Ecode`、`Model`、`Assembly`、`Board`、`Quantity`、`Main Source`（=主料）、`Time`；其他欄位保留空值
- 會新增 `主料` 欄位：主料列為自身 `Item`，子料列為對應主料 `Item`

### 5.7 只輸出指定欄位檔

```bash
python bom_excel_tool.py "data/BOM COST.xlsx" --selected-only -o "output/result.csv"
```

只會輸出：`output/result_selected.csv`

### 5.8 只輸出完整欄位檔

```bash
python bom_excel_tool.py "data/BOM COST.xlsx" --no-selected -o "output/result.csv"
```

只會輸出：`output/result.csv`

## 6. 輸出資料說明

### 6.1 完整欄位檔

完整欄位輸出包含：

- 三列表頭合併後欄位
- 新增欄位 `Board`
- 新增欄位 `M/S`（固定 `M`）
- 新增欄位 `Main Source`（取 `Item` 欄位值）
- `2ND_SOURCE_TOTAL` 起依四欄一組轉為 `Sub_*` 欄位群組
- 新增欄位 `主料`，並將 `Sub_*` 展開為子料列（`M/S = S`）

### 6.2 指定欄位檔（固定順序）

指定欄位檔只保留以下欄位，且順序固定：

1. `Model`
2. `Assembly`
3. `Board`
4. `Item`
5. `Quantity`
6. `Last BPA Currency`
7. `Last BPA Price`
8. `Lead Time`
9. `MFG`
10. `MPN`
11. `Ecode`
12. `M/S`
13. `Main Source`
14. `Time`

`Board` 產生規則：

1. 逐列讀取資料
2. 若 `Level` 為 `1`，先檢查 `Description` 是否包含以下關鍵字（不分大小寫）：
   - `MECHANICAL`、`WIFI HIGH`、`KEYPAD`、`LED`、`PACKAGE`
   - `WIFI LOW`、`PALLET`、`MAIN BOARD`、`POE`、`NVME`
3. 若命中關鍵字，`Board` 取該標準字串；未命中則保留原 `Description` 完整值
4. 其後列沿用該 `Board`
5. 遇到下一個 `Level = 1` 時覆蓋

`M/S` 與 `Main Source` 規則：

1. 若找到 `Item` 欄位，新增 `M/S` 全欄固定 `M`
2. 新增 `Main Source`，每列值等於該列 `Item`

`Sub_*` 規則（由 `2ND_SOURCE_TOTAL` 起算）：

1. 每四欄一組
2. 第 1 欄改名 `Sub_n`
3. 第 2 欄改名 `Sub_n_Currency`
4. 第 3 欄改名 `Sub_n_Price`
5. 第 4 欄刪除
6. 下一組依序為 `Sub_2`、`Sub_2_Currency`、`Sub_2_Price`...

`Sub_*` 展開規則：

1. 每列保留主料列（`M/S = M`）
2. 對每個非空 `Sub_n` 產生一列子料
3. 子料列 `Item` 取 `Sub_n`
4. 子料列 `M/S` 固定為 `S`
5. 子料列 `Last BPA Currency` / `Last BPA Price` 分別取 `Sub_n_Currency` / `Sub_n_Price`
6. 子料列只帶入主料欄位：`Ecode`、`Model`、`Assembly`、`Board`、`Quantity`、`Main Source`（=主料）、`Time`；其他欄位在子料列維持空值（留白），僅 `Last BPA Currency/Last BPA Price` 會帶入 `Sub_n_Currency/Sub_n_Price`

執行順序重點：

1. 讀取 `INPUT(BOM COST)`：產生主料列、並填入 `Board/Model/Time`
2. 從 `BOM COST` 取得 `Ecode`
3. 若提供 `--sub-source(EMS BOM COST)`：用 EMS 的 `Sub_*` 展開產生子料列
4. 子料列的 `Last BPA Currency/Last BPA Price` 取自 EMS 的 `Sub_n_Currency/Sub_n_Price`

## 7. 錯誤排除

- `找不到檔案`：確認 `INPUT` 路徑正確
- `找不到第一欄為「Level」的標題列`：
  - 檢查該工作表第一欄是否真有 `Level`
  - 或改用 `--marker` 指定實際文字
- `輸出副檔名請使用 .csv 或 .xlsx`：請改用支援的輸出副檔名

## 8. 注意事項

- 目前讀取模式為 `read_only=True`，適合大檔案
- 空白資料列會自動略過
- 欄位範圍會同時參考「表頭」與「資料列」；即使尾端沒有欄名，只要資料有值仍會保留
- 去重/唯一鍵：工具會以 `(Board + Model + Item)` 做主料與子料匹配，並嚴格檢查 `BOM COST` / `EMS BOM COST` 任一端若出現重複 key 會直接拋出 `ValueError`
