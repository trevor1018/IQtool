# CCM CV Simulator

使用 ColorChecker 色卡擷取兩張圖片的 24 色塊 sRGB 數值，用於 CCM（Color Correction Matrix）分析與模擬。

## 功能

- 自動偵測圖片中的 Macbeth ColorChecker 色卡
- 擷取並轉換 24 色塊為 sRGB 值
- 輸出成對比較的 Excel 報表

## 使用方式

將**兩張**圖片放入 `Macbeth/` 資料夾：

```
Macbeth/
├── 1_before.jpg
└── 2_after.jpg
```

執行：
```bash
python CCMCVsimulator.py
```

## 輸出

程式會產生 `CCMCVsimulator_YYYY_MM_DD_XXXXX.xlsm`，包含：
- 第一張圖片的 24 色塊 sRGB（B~D 欄）
- 第二張圖片的 24 色塊 sRGB（I~K 欄）

## 相依套件

```bash
pip install opencv-python numpy xlwings colour-checker-detection
```

## 注意事項

- 每次只能處理一組（2 張）照片
- 需要安裝 Microsoft Excel
- `CCMCVsimulator.xlsm` 為範本檔案，請勿刪除
- 圖片格式僅支援 `.jpg` / `.JPG`
