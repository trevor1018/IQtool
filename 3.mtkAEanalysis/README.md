# MTK AE Analysis

解析 MediaTek 平台的 AE（Auto Exposure）參數，自動產生分析報表與 Histogram 圖表。

## 功能

- 解析 `.exif` / `.txt` 格式的 AE 參數檔案
- 自動擷取多項 AE 關鍵數據：
  - TargetMidRatioTbl
  - MainTHD（BV / based / exp）
  - HS（Highlight Suppression）相關參數
  - Face AE 參數
  - Night Scene 參數
- 產生圖片的灰階 Histogram，標註 BT/MT/DT 閾值線
- 支援成對圖片比較（refer 模式）
- 輸出含圖片與圖表的 Excel 報表

## 資料夾結構

```
mtkAEanalysis/
├── mtkAEanalysis.py
├── mtkAEanalysis_SX3.xlsm    ← 範本檔案（勿刪）
├── faceCase.png              ← Face AE 參考圖（選用）
└── [yourPath]/               ← 放置待分析的檔案
    ├── 001_xxx.exif
    ├── 001_xxx.jpg
    ├── 002_xxx.exif
    ├── 002_xxx.jpg
    └── ...
```

## 使用方式

```bash
python mtkAEanalysis.py
```

執行後會透過檔案對話框選擇要分析的資料夾。

## 輸出

程式會產生 `mtkAEanalysis_YYYY_MM_DD_XXXXX_start_end.xlsm`，包含：
- AE 參數數據表
- 原圖縮圖
- 灰階 Histogram（含 BT/MT/DT 閾值標線）
- Face AE 資訊（如有）

## Histogram 標線說明

| 顏色 | 線型 | 說明 |
|------|------|------|
| 🔴 紅色實線 | — | BT (Bright Tone) THD |
| 🔴 紅色虛線 | ⋯ | BT Final Y |
| 🔵 藍色實線 | — | MT (Mid Tone) THD |
| 🔵 藍色虛線 | ⋯ | MT Final Y |
| 🟢 綠色實線 | — | DT (Dark Tone) THD |
| 🟢 綠色虛線 | ⋯ | DT Final Y |

## 相依套件

```bash
pip install opencv-python numpy openpyxl matplotlib pillow
```

## 注意事項

- `.exif` / `.txt` 與 `.jpg` 需同名配對
- 每 20 組自動分檔儲存
- 需要 Microsoft Excel 開啟 `.xlsm` 報表
