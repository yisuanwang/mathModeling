# 将数据处理成csv文件

import openpyxl
# import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

path = './data/data1.xlsx'
wb = openpyxl.load_workbook(path)
sh = wb['企业的订货量（m³）']
data_dhl = []
for row in list(sh.rows):
    data_dhl.append([ cell.value for cell in row[2:242]])

# print(data_dhl)

sh = wb['供应商的供货量（m³）']
data_ghl = []
for row in list(sh.rows):
    data_ghl.append([ cell.value for cell in row[2:242]])

# test=pd.DataFrame(data_dhl[1:])#数据有三列，列名分别为one,two,three

week_list = np.linspace(1,240,240)

# 147 283 407
line = 100
# print(week_list.__len__())
print(data_dhl[line])
print(data_ghl[line])


fig = plt.figure()
ax = fig.add_subplot(111)

plt.title('raw material A')  # 标题
plt.grid()

plt.plot(week_list,data_dhl[line], linestyle = '-',c='blue',label='Order quantity')
plt.plot(week_list,data_ghl[line], linestyle = '-',c='red',label='Supply quantity')

ax.set_xlabel('week')
ax.set_ylabel('value')
print('drawing...')
plt.savefig("./240周各原料订货量和供货量统计/A.png", dpi=1000, bbox_inches='tight')
plt.show()

# print(test)

# test.to_csv("e:/testcsv.csv",encoding="gbk")

# for d in data[1]:
#     print(d)
# wb.save(path)
wb.close()