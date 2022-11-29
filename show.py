#!C:\Users\xsy\AppData\Local\Programs\Python\Python38\python.exe
# -*- Coding: UTF-8 -*-
# Author: 小圆子
# CreateDate: 2022/11/17 19:52
import sys
import csv
import pandas as pd
wb = pd.read_excel(sys.argv[1], sheet_name=None)
x = {"周一": "1", "周二": "2", "周三": "3", "周四": "4", "周五": "5"}
y = {"第四周": "4", "第五周": "5", "第六周": "6", "第七周": "7", "第八周": "8", "第九周": "9", "第十周": "10", "第十一周": "11", "第十二周": "12", "第十三周": "13", "第十四周": "14", "第十五周": "15", "第十六周": "16", "第十七周": "17", "第十八周": "18", "第十九周": "19"}
sheet, week, day, work = input("Input(天山堂 第四周 周一 7-8节):\n").split(' ')
for _ in wb:
    if _ == sheet:
        print("Right!")
        break
else:
    print("Not Right!")
    exit()

for i, n in enumerate(wb[sheet].values[0]):
    if n == "到课率":
        h = i
        break
for i, n in enumerate(wb[sheet].values[0]):  # i为序号
    if n == "课程名称":
        a = i
    elif n == "开课院系":
        b = i
    elif n == "课程号":
        c = i
    elif n == "主讲教师姓名":
        d = i
    elif n == "节次":
        e = i
    elif n == "教室":
        f = i
    elif n == "备注":
        g = i
    elif n == "星期":
        m = i

work_lst = [work]
if work == "9-11节":
    work_lst += ["9-10节", "7-10节"]

data_fill = []
number = 0
for i in wb[sheet].values:
    if i[m] == x[day] and i[e] in work_lst and (i[h] < 0.7 or i[g] == "线上（老师迟到15min）"):
        number += 1
        if i[h] < 0.6:
            print(number, week, day, i[e], "到课率为%-5.2f" %(i[h] * 100), i[g], i[f], "ATTENTION!!!")
        else:
            print(number, week, day, i[e], "到课率为%-5.2f" %(i[h] * 100), i[g], i[f])
print("finish!")

