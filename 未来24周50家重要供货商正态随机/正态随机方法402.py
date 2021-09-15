import openpyxl
# import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import random
import math
path = './工作簿2.xlsx'
print('正在加载文件')
wb = openpyxl.load_workbook(path)
sh = wb['Sheet1']

均值 = []
标准差 = []
tp = []
print('正在加载均值和标准差')
for row in list(sh.rows)[1:403]:  # 第几行就是第几家供应商
    # tp.append(str(row[1].value))
    均值.append(float(row[3].value))
    标准差.append(float(row[2].value))

# print(均值)
# print(标准差)

for i in range(1,403):
    row = list(sh.rows)[i]  # 第几行就是第几家供应商
    print(i)
    for date in range(4,4+24):
        # print(date)
        x = 0
        random3 = np.random.randn(1) # 生成一个标准+正态分布量
        x = 标准差[i-1] * random3 + 1.7*均值[i-1]
        x = max(0, x)  # 不能生产负数
        # x = math.fabs(x)
        x = float(x)
        row[date].value = x
        # print(date)
        print( row[date].value)

wb.save(path)
print('finished.')