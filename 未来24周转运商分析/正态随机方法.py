import openpyxl
# import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import random
import math

path = './附件2 近5年8家转运商的相关数据.xlsx'
print('正在加载文件')
wb = openpyxl.load_workbook(path)
sh = wb['Sheet2']

均值 = [1.904769167,
0.921370417,
0.186055556,
1.570482353,
2.889825301,
0.543761111,
2.078833333,
1.010282759
]
标准差 = [0.657302318,
0.480451243,
0.477394567,
1.937674233,
1.970063605,
1.351074514,
1.435411881,
1.115299628
]

for i in range(0,8):
    row = list(sh.rows)[i]  # 第几行就是第几家供应商

    for date in range(0,24):
        x = 0
        random3 = np.random.randn(1) # 生成一个标准正态分布量
        x = 标准差[i-1] * random3 + 均值[i-1]
        # x = max(0, x)  # 不能生产负数
        x = abs(x)
        x = float(x)
        row[date].value = x
        # print(date)

wb.save(path)
print('finished.')