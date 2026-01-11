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
                ws.range((j + 12, col)).value = color
        print(f"{os.path.basename(item)} is ok!")
    except:
        columns = [2, 3, 4] if index == 1 else [7, 8, 9]
        for j in range(24):
            for col in columns:
                ws.range((j + 12, col)).value = "error"
        print(f"***** {os.path.basename(item)} *****")

print("colorCheckerAnalysis is runing...")
localtime = time.localtime()
clock = str(60*60*localtime[3] + 60*localtime[4] + localtime[5])

yourPath = "Macbeth"
allFileList = [os.path.join(yourPath, file) for file in os.listdir(yourPath)]
allFileList = np.sort(allFileList,axis=0)
allFileList = list(filter(file_filter, allFileList))

def extract_filename(filepath):
    return os.path.basename(filepath)

app = xw.App(visible=False)
fn = 'colorCalculate.xlsm'
wb = app.books.open(fn)

file = f'colorCalculate_{localtime[0]}_{localtime[1]}_{localtime[2]}_{clock}.xlsm'
wb.save(file)

macro_vba = wb.app.macro('CopySheetWithChart')

if np.size(allFileList) % 2 == 0:
    for i, (item1, item2) in enumerate(zip(allFileList[0::2], allFileList[1::2]), start=0):
        
        img1 = cv2.imread(item1, cv2.IMREAD_COLOR)
        img2 = cv2.imread(item2, cv2.IMREAD_COLOR)
        
        base = os.path.splitext(os.path.basename(item1))[0][:-2]
        macro_vba(base)
        ws = wb.sheets[base]

        detect_color(img1, item1, ws, 1)
        detect_color(img2, item2, ws, 2)
    
    wb.sheets['(default)'].delete()
    wb.sheets[0].activate()
    wb.save(file)
    app.quit()
else:
    print("The photos are not in pairs!")

print("colorCheckerAnalysis is ok!")
os.system("pause")