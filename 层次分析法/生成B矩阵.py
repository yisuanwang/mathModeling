import openpyxl
# import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

path = './聚类结果.xlsx'
wb = openpyxl.load_workbook(path)
sh = wb['Sheet2']
data_dhl = []

最大供货量 = []
订单次数 = []
完成率 = []
平均供货量 = []

for row in list(sh.rows)[1:376]:
    最大供货量.append(row[3].value)
    订单次数.append(row[4].value)
    完成率.append(row[5].value)
    平均供货量.append(row[6].value)

# print(最大供货量)

B_最大供货量 = []
B_订单次数 = []
B_完成率 = []
B_平均供货量 = []

for i in range(0, 374):
    tp1, tp2, tp3, tp4 = [], [], [], []
    # print(i, len(tp1))
    for j in range(0, 374):
        # print('i = ', i, '  j = ', j)
        tp1.append(最大供货量[i] / 最大供货量[j])
        tp2.append(订单次数[i] / 订单次数[j])
        tp3.append(完成率[i] / 完成率[j])
        tp4.append(平均供货量[i] / 平均供货量[j])
        pass
    # print('tp1 len = ', len(tp1))
    B_最大供货量.append(tp1)
    B_订单次数.append(tp2)
    B_完成率.append(tp3)
    B_平均供货量.append(tp4)
