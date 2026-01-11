import cv2
import numpy as np
import os
import time
import xlwings as xw
from colour_checker_detection import detect_colour_checkers_segmentation

# Global variables
drawing = False
top_left_pt, bottom_right_pt = (-1, -1), (-1, -1)
coordinates = []

def file_filter(f):
    if f[-4:] in ['.jpg', '.JPG']:
        return True
    else:
        return False
    
def RGBtosRGB(rgb):
    srgb = []
    for i in range(0,3):
        if rgb[i] > 0.00304:
            V = (1+0.055)*((rgb[i])**(1/2.4))-0.055
            srgb.append(round(V*255,2))
        else:
            V = 12.92*rgb[i]
            srgb.append(round(V*255,2))
    return srgb
            
def detect_color(img, item, ws, index):
    try:
        img_crop = detect_colour_checkers_segmentation(img, additional_data=True)[0]
        for j in range(24):
            colors = RGBtosRGB(img_crop[0][j])
            columns = [2, 3, 4] if index == 1 else [7, 8, 9]
            for col, color in zip(columns, colors):
                ws.range((j + 12, col)).value = color  # 使用 xlwings 的方式指定儲存格
        if index == 2:
            print(f"{os.path.basename(item)} is ok!")
    except:
        columns = [2, 3, 4] if index == 1 else [7, 8, 9]
        for j in range(24):
            for col in columns:
                ws.range((j + 12, col)).value = "error"  # 使用 xlwings 的方式指定儲存格
        print(f"***** {os.path.basename(item)} *****")

print("colorCheckerAnalysis is runing...")
localtime = time.localtime()
clock = str(60*60*localtime[3] + 60*localtime[4] + localtime[5])

yourPath = "Macbeth"
allFileList = [os.path.join(yourPath, file) for file in os.listdir(yourPath)]
allFileList = np.sort(allFileList,axis=0)
allFileList = list(filter(file_filter, allFileList))

# 定義一個函數從路徑中提取檔名
def extract_filename(filepath):
    return os.path.basename(filepath)

# 找到基準圖片
base_img_file = next((item for item in allFileList if extract_filename(item).startswith('1_')), None)

if base_img_file:
    # 開始
    app = xw.App(visible=False)
    fn = 'colorCalculate.xlsm'
    wb = app.books.open(fn)  # 使用xlwings加載工作簿

    # 儲存新的工作簿
    file = f'colorCalculate_{localtime[0]}_{localtime[1]}_{localtime[2]}_{clock}.xlsm'
    wb.save(file)

    macro_vba = wb.app.macro('CopySheetWithChart')
    
    # 取得所有2_開頭的圖片
    other_files = [item for item in allFileList if extract_filename(item).startswith('2_')]
    
    if other_files:
        for idx, item in enumerate(other_files, start=0):
            # 用2_後的文字作為target.title
            base_name = os.path.splitext(extract_filename(item))[0][2:]
            macro_vba(base_name)
        wb.sheets[0].activate()
        wb.save()
        app.quit()
    else:
        print("No other photos with prefix '2_' found!")
        app.quit()
else:
    print("Base photo with prefix '1_' not found!")
    app.quit()

if base_img_file:
    base_img = cv2.imread(base_img_file, cv2.IMREAD_COLOR)
    if other_files:
        app = xw.App(visible=False)
        wb = app.books.open(file)  # 使用xlwings打開工作簿

        for idx, item in enumerate(other_files, start=0):
            other_img = cv2.imread(item, cv2.IMREAD_COLOR)
            ws = wb.sheets[int(idx+1)]  # 使用xlwings選取工作表
            detect_color(base_img, base_img_file, ws, 1)
            detect_color(other_img, item, ws, 2)
        
        # 刪除指定的工作表，並激活第一個工作表
        wb.sheets['(default)'].delete()
        wb.sheets[0].activate()

        wb.save(file)
        app.quit()
    else:
        print("No other photos with prefix '2_' found!")
else:
    print("Base photo with prefix '1_' not found!")

print("colorCheckerAnalysis is ok!")
os.system("pause")