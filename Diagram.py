# -*- coding: utf-8 -*-
"""
Created on Thu Aug 31 11:12:56 2023

@author: TrevorEChen
"""

import numpy as np
import matplotlib.pyplot as plt
from skimage.color import lab2rgb

def lab_to_rgb(L, a, b):
    # 使用skimage库进行转换
    lab = np.array([[[L, a, b]]])
    rgb = lab2rgb(lab)
    return rgb[0][0]

fig, ax = plt.subplots(figsize=(8, 8))

# 创建一个A和B的网格
a = np.linspace(-90, 90, 500)
b = np.linspace(-90, 90, 500)
a, b = np.meshgrid(a, b)

# 使用固定的L值
L = 90

# 转换每对AB值到RGB
w, h = a.shape
image = np.zeros((w, h, 3))
for i in range(w):
    for j in range(h):
        image[i, j] = lab_to_rgb(L, a[i, j], b[i, j])

ax.imshow(image, origin='lower', extent=(-90, 90, -90, 90))

# 画a=0 和 b=0 的网格线
ax.axvline(0, color='gray', linestyle='--', linewidth=1)
ax.axhline(0, color='gray', linestyle='--', linewidth=1)
ax.set_xlim(-90, 90)
ax.set_ylim(-90, 90)

plt.show()