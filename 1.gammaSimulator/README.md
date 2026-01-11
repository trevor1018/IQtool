# Gamma AE Precheck

比較 Target 與 Original 兩組 Gamma 數據，自動彙整至 Excel 報表。

## 功能

- 批次讀取 CSV 格式的 Gamma 數據
- 自動對應 15 組測試場景的欄位排列
- 支援不同版本的 CSV 格式（自動判斷 3.7 版本）
- 輸出含 VBA 巨集的 Excel 報表

## 資料夾結構

```
gammaAEprecheck/
├── gammaAEprecheck.py
├── gammaSummary.xlsm      ← 範本檔案（勿刪）
├── 0.target/              ← 目標數據
│   ├── scene01.csv
│   ├── scene02.csv
│   └── ...
└── 1.original/            ← 原始數據
    ├── scene01.csv
    ├── scene02.csv
    └── ...
```

> ⚠️ 兩個資料夾的 CSV 檔案數量與排序需一致

## 使用方式

```bash
python gammaAEprecheck.py
```

## 輸出

程式會產生 `gammaSummary_YYYY_MM_DD_XXXXX.xlsm`，包含：
- Target 數據（左側欄位）
- Original 數據（右側欄位，偏移 18 欄）
- 20 筆 Gamma 數據點

## 相依套件

```bash
pip install openpyxl numpy
```

## 欄位對應

程式內建 15 組場景的欄位順序：
```
18, 9, 6, 10, 7, 4, 11, 12, 13, 8, 5, 14, 15, 16, 17
```

如需調整順序，修改 `order` 變數即可。
