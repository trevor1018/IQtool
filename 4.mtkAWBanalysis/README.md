# MTK AWB Analysis

解析 MediaTek 平台的 AWB（Auto White Balance）參數，自動產生分析報表與 Gray World 色彩空間圖。

## 功能

- 解析 `.exif` / `.txt` 格式的 AWB 參數檔案
- 自動擷取多項 AWB 關鍵數據：
  - CCT（色溫）
  - Output Gain (R/B)
  - 各光源 Neutral PB Number
  - Light Source Probability (P1/P2)
  - Spatial Gain
  - Exclude / Extra Color 區域
- 繪製 Gray World 色彩空間圖，標註各光源區域
- 支援成對圖片比較（refer 模式）
- 輸出含圖片與圖表的 Excel 報表

## 資料夾結構

```
mtkAWBanalysis/
├── mtkAWBanalysis.py
├── mtkAWBanalysis.xlsm       ← 範本檔案（勿刪）
└── Exif/                     ← 放置待分析的檔案
    ├── 001_xxx.exif
    ├── 001_xxx.jpg
    ├── 002_xxx.exif
    ├── 002_xxx.jpg
    └── ...
```

## 使用方式

```bash
python mtkAWBanalysis.py
```

執行後：
1. 選擇 AWB.cpp 參數檔案（用於讀取 Light Source Probability）
2. 輸入是否有參考圖（0: 無, 1: 有）

## 輸出

程式會產生 `mtkAWBanalysis_YYYY_MM_DD_XXXXX_start_end.xlsm`，包含：
- AWB 參數數據表
- 原圖縮圖
- Gray World 色彩空間圖

## Gray World 光源區域顏色對照

| 顏色 | 光源 |
|------|------|
| 🔴 紅色 | T (Tungsten / 鎢絲燈) |
| 🟠 橘色 | WF (Warm Fluorescent) |
| 🟡 黃色 | F (Fluorescent) |
| 🟢 綠色 | CWF (Cool White Fluorescent) |
| 🔵 藍色 | D (Daylight) |
| 🔵 深藍 | S (Shade) |
| 🔵 青色 | DF (Daylight Fluorescent) |

## 相依套件

```bash
pip install opencv-python numpy openpyxl matplotlib pillow
```

## 注意事項

- `.exif` / `.txt` 與 `.jpg` 需同名配對
- 每 20 組自動分檔儲存
- 需要 Microsoft Excel 開啟 `.xlsm` 報表
