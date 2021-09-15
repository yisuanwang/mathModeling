import openpyxl
# import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import random
import math

path = './工作簿2.xlsx'
print('正在加载文件')
wb = openpyxl.load_workbook(path)
sh = wb['Sheet3']

均值 = []
标准差 = []
tp = []
print('正在加载均值和标准差')
for row in list(sh.rows)[0:410]:
    for cell in row[0:24*8]:
        v = random.random() * 28
        print(v)
        if v == 0:
            cell.value = ''
        else:
            cell.value = float(v)

wb.save(path)
print('finished.')
