# -*- coding: utf-8 -*-
"""
Created on Tue Aug 29 11:37:16 2023

@author: TrevorEChen
"""

import os
import sys
import re

def rename_file(filepath):
    # 分割檔案路徑和檔案名
    dir_name, file_name = os.path.split(filepath)
    # 分割檔名和副檔名
    base_name, ext = os.path.splitext(file_name)
    
    # 使用正則表達式匹配並提取檔名中的特定部分
    match = re.search(r'Output_\d+_\d+_\d+_\d+_(.*)', base_name)
    if match:
        new_name = '2_' + match.group(1) + ext
    else:
        # 如果不匹配，則不更改檔名
        new_name = file_name
    
    # 得到新的檔案路徑
    new_filepath = os.path.join(dir_name, new_name)
    
    # 重新命名
    os.rename(filepath, new_filepath)
    print(f"Renamed {file_name} to {new_name}")

if __name__ == "__main__":
    # 檢查是否有提供檔案參數
    if len(sys.argv) < 2:
        print("Please drag and drop a file onto this exe.")
    else:
        for filepath in sys.argv[1:]:
            if os.path.exists(filepath):
                rename_file(filepath)
            else:
                print(f"File {filepath} does not exist.")
    input("\nPress Enter to exit...")
