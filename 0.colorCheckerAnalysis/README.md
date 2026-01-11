# Color Checker Analysis

自動偵測 Macbeth ColorChecker 色卡並分析 24 色塊的 sRGB 數值，結果輸出至 Excel 報表。

## 功能

- 自動偵測圖片中的 ColorChecker 色卡位置
- 擷取 24 色塊的 RGB 值並轉換為 sRGB
- 支援成對比較（如：調整前 vs 調整後）
- 自動產生帶有圖表的 Excel 報表

## 使用方式

### 標準模式（成對比較）

將圖片放入 `Macbeth/` 資料夾，檔名需成對且依序排列：
```
Macbeth/
├── 1_default.jpg
├── 1_modify.jpg
├── 2_default.jpg
├── 2_modify.jpg
```

執行：
```bash
python colorCheckerAnalysis.py
# 或直接執行
run.bat
```

### 基準比較模式

以 `1_` 開頭的圖片作為基準，與所有 `2_` 開頭的圖片進行比較：
```
Macbeth/
├── 1_baseline.jpg
├── 2_sceneA.jpg
├── 2_sceneB.jpg
├── 2_sceneC.jpg
```

執行：
```bash
python colorCheckerAnalysis_modify.py
# 或直接執行
run_modify.bat
```

## 輸出

程式會產生 `colorCalculate_YYYY_MM_DD_XXXXX.xlsm`，包含：
- 每組比較的獨立工作表
- 24 色塊的 sRGB 數值
- 自動生成的色差圖表

![輸出範例](images/exceloutput.jpg "Excel 輸出結果")

## 相依套件

```bash
pip install opencv-python numpy xlwings colour-checker-detection

* colour_checker_detection直接安裝版本有可能無法使用，請替換成專案中的版本
```

## 附加工具

- **Diagram.py** - 繪製 Lab 色彩空間的 a*b* 平面圖，用於視覺化色彩分布

## 注意事項

- 需要安裝 Microsoft Excel（xlwings 需要）
- `colorCalculate.xlsm` 為範本檔案，請勿刪除
- 圖片格式僅支援 `.jpg` / `.JPG`
