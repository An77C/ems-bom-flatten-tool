# BOM Excel Tool 開發者文件

## 1. 專案概觀

此專案目前核心檔案：

- `bom_excel_tool.py`：主程式（CLI + 讀取邏輯）
- `requirements.txt`：相依套件

工具目標是將 BOM 類 Excel 扁平化，並補上 `Board`、`M/S`、`Main Source`、`Sub_*` 與子料展開列供後續分析。

## 2. 程式流程

主流程位於 `main()`：

1. `_build_parser()` 解析 CLI 參數
2. `_read_main_dataframe()` 讀取主檔（單表或多表合併）並轉換資料
3. `_apply_ecode_if_needed()` 套用 Ecode 對照（可略過）
4. `_apply_final_transforms()` 套用後處理（如 Sub 展開）
5. `--preview` 時走 `_print_preview()`
6. 否則走 `_write_outputs()` 輸出完整檔與指定欄位檔

## 3. 核心資料結構

### `BomReadInfo`

`@dataclass(frozen=True)`，回傳讀取結果摘要：

- `sheet`
- `level_row`
- `header_rows`
- `data_start_row`
- `column_count`

可作為上層流程的可觀測資訊（log/debug）。

## 4. 主要函式責任

- `find_level_header_row()`：尋找標題中心列（第一欄值匹配 marker）
- `merge_three_row_headers()`：將三列表頭合併
- `uniquify_column_names()`：欄名去重並加流水號
- `_read_header_rows()`：讀取中心列上下三列
- `_read_data_rows()`：串流讀取資料列、忽略空列
- `_effective_max_col()`：以「表頭 + 資料」共同決定最終欄位範圍
- `_build_board_column()`：依 `Level=1` 的 `Description` 向下填補 `Board`
- `BOARD_KEYWORDS`：`Board` 關鍵字常數（集中維護分類詞）
- `_build_main_source_columns()`：建立 `M/S` 與 `Main Source`
- `_rename_sub_columns()`：將 `2ND_SOURCE_TOTAL` 起的尾端欄位轉為 `Sub_*` 群組並刪除每組第 4 欄
- `_build_item_to_ecode_map()`：由對照檔建立 `(Assembly, Item) -> Customer PN` 對照
- `_apply_ecode_mapping()`：將對照回填至主檔 `Ecode` 欄位
- `_expand_sub_rows()`：將 `Sub_n` 展開為子料列，並沿用主料關鍵欄位
- `read_bom_multi_sheet()`：未指定 `--sheet` 時合併主檔所有可用工作表
- `_write_output()`：依副檔名寫出 CSV / XLSX

## 5. 欄位對應規則

### 5.1 表頭規則

- 預設 marker：`Level`
- 合併順序：上一列 -> 中心列 -> 下一列
- 分隔符號：預設空白字元 ` `
- `-----` 這類分隔字元視為空字串

### 5.2 `Board` 規則

- 找 `Level` 欄與 `Description` 欄（先精準比對，找不到用前綴比對）
- `Level` 正規化後等於 `"1"` 才更新當前 `Board`
- 更新 `Board` 前，會先以 `BOARD_KEYWORDS` 進行關鍵字比對（不分大小寫）
- 命中關鍵字時寫入標準分類字串；未命中則保留原始 `Description`
- 每列寫入當前 `Board`，直到下一個 `Level=1`

### 5.3 `M/S` 與 `Main Source` 規則

- 找到 `Item` 欄後：
  - 新增 `M/S`，全欄固定 `M`
  - 新增 `Main Source`，值來自該列 `Item`

### 5.4 `Sub_*` 規則

- 以 `2ND_SOURCE_TOTAL` 為 anchor
- 每四欄為一組：
  - 1 -> `Sub_n`
  - 2 -> `Sub_n_Currency`
  - 3 -> `Sub_n_Price`
  - 4 -> 刪除
- 轉換順序必須先 `_rename_sub_columns()`，再 `_build_main_source_columns()`，避免尾端分組覆蓋 `M/S`、`Main Source`

### 5.5 `Ecode` 對照規則（雙檔模式）

- 以 `--ecode-source` 指定對照檔
- 未指定 `--ecode-sheet` 時，掃描對照檔所有工作表
- 對照檔以「工作表名稱」作為 `Assembly` 維度
- 每個工作表中使用 `Item` 與 `Customer PN`（或 `Customer PN` 近似欄名）建立映射
- 主檔以 `(Assembly, Item)` 回填 `Ecode`
- 若主檔已有 `Ecode`，僅在新對照有值時覆蓋

### 5.6 `Sub_*` 展開子料列規則

- 若存在 `Item` 與 `M/S` 欄位，將每列保留為主料列
- 若不存在 `主料` 欄位，會自動新增：
  - 主料列 `主料 = Item`
  - 子料列 `主料 = 主料 Item`
- 每個非空 `Sub_n` 產生一列子料：
  - `Item = Sub_n`
  - `M/S = S`
  - `Last BPA Currency = Sub_n_Currency`
  - `Last BPA Price = Sub_n_Price`
  - 沿用主料欄位：`Ecode`、`Model`、`Assembly`、`Board`、`Quantity`、`Main Source`、`Time`
- 其餘欄位預設為空值，避免誤帶主料專屬資訊（如 `MFG`、`MPN`、`Lead Time`）

## 6. 擴充建議

### 6.1 新增衍生欄位

建議在 `read_bom_with_merged_headers()` 中，採獨立函式方式處理，保持單一職責：

- 例如新增 `_build_<new_column>_column(df)`
- 最後在主流程串接

### 6.2 自訂輸出格式

若要支援 JSON / Parquet，建議擴充 `_write_output()`：

- 新增副檔名分支
- 保持錯誤訊息一致

### 6.3 可測試性

核心邏輯函式已接近純函式，可直接單元測試：

- `merge_three_row_headers()`
- `uniquify_column_names()`
- `_normalize_level_for_compare()`
- `_build_board_column()`

## 7. 測試建議

建議建立 `tests/` 並使用 `pytest`：

- `test_merge_three_row_headers.py`
  - 測試底線分隔列、空值、重複字串
- `test_board_fill.py`
  - 測試多段 `Level=1` 覆蓋行為
- `test_main_source_columns.py`
  - 測試 `M/S` 固定值與 `Main Source` 對應 `Item`
- `test_sub_columns.py`
  - 測試 `2ND_SOURCE_TOTAL` 起的四欄一組改名與刪除
- `test_sub_expand_rows.py`
  - 測試 `Sub_n` 展開為子料列與欄位帶入規則
- `test_output_writer.py`
  - 測試副檔名合法性與錯誤訊息

範例（手動 smoke test）：

```bash
python bom_excel_tool.py "data/91-017-507025B EMS BOM COST.xlsx" --preview
python bom_excel_tool.py "data/91-017-507025B EMS BOM COST.xlsx" --ecode-source "data/BOM COST.xlsx" -o "output/result.csv"
```

## 8. 維護注意事項

- 目前使用 `openpyxl.load_workbook(..., read_only=True, data_only=True)`：
  - 適合大檔
  - 請避免再以 `max_row` 做全範圍掃描
- 欄位右邊界以「表頭 + 資料列」共同判定，避免漏掉無欄名但有值的尾端欄位
- 編碼輸出採 `utf-8-sig` 是為了 Excel 相容性
- 任何欄名變更都可能影響 `Board` 欄位定位，改動時需同步測試
