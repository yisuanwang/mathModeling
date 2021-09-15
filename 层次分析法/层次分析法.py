# 程序文件Pan9_1.py
from scipy.sparse.linalg import eigs
import openpyxl
from numpy import array, hstack

B_最大供货量 = []
B_订单次数 = []
B_完成率 = []
B_平均供货量 = []

def makeB():
    path = './聚类结果分析.xlsx'
    wb = openpyxl.load_workbook(path)
    sh = wb['Sheet2']
    data_dhl = []

    最大供货量 = []
    订单次数 = []
    完成率 = []
    平均供货量 = []

    for row in list(sh.rows)[1:378]:
        # print(row[3].value)
        最大供货量.append(row[3].value)
        订单次数.append(row[4].value)
        完成率.append(row[5].value)
        平均供货量.append(row[6].value)

    # print(最大供货量)

    for i in range(0, 377):
        tp1, tp2, tp3, tp4 = [], [], [], []
        # print(i, len(tp1))
        for j in range(0, 377):
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
    pass

makeB() # 生成B矩阵

a = array([[1, 1 / 3, 1 / 4, 1 / 3],
           [3, 1, 1 / 2, 1 / 2],
           [4, 2, 1, 1],
           [3, 2, 1, 1]])

# a = array([[1, 1 / 2, 5, 5, 3], [2, 1, 7, 7, 5], [1 / 5, 1 / 7, 1, 1 / 2, 1 / 3],
#            [1 / 5, 1 / 7, 2, 1, 1 / 2], [1 / 3, 1 / 5, 3, 2, 1]])
L, V = eigs(a, 1) # L ：最大特征值
print('L:',L,'\nV:',V)
CR = (L - 5) / 4 / 1.12  # 计算矩阵A的一致性比率
W = V / sum(V)
print("最大特征值 L =", L)
print("最大特征值对应的特征向量 W =\n", W)
print("CR=", CR)

B1 = array(B_最大供货量)
L1, P1 = eigs(B1, 1)
P1 = P1 / sum(P1)
print("P1=", P1)

B2 = array(B_订单次数)
t2, P2 = eigs(B2, 1)
P2 = P2 / sum(P2)
print("P2=", P2)

B3 = array(B_完成率)
t3, P3 = eigs(B3, 1)
P3 = P3 / sum(P3)
print("P3=", P3)

B4 = array(B_平均供货量)
t4, P4 = eigs(B4, 1)
P4 = P4 / sum(P4)
print("P4=", P4)

# print(type(B4[0][0]))

K = hstack([P1, P2, P3, P4]) @ W  # 矩阵乘法
print('*'*40)
# print("K=", K)
for i in K:
    # print(str(i[0]))
    print(str(i[0])[1:20])
    pass
