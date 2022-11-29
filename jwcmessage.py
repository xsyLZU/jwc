#!/bin/env python
# -*- Coding: UTF-8 -*-
# Author: hashcaaat
# CreateDate: 2022/9/15 20:54
import sys
import pandas as pd
wb = pd.read_excel(sys.argv[1], sheet_name=None)
x = {"周一": "1", "周二": "2", "周三": "3", "周四": "4", "周五": "5"}
sheet, week, day, work = input("Input(秦岭堂 第四周 周一 7-8节):\n").split(' ')
for _ in wb:
    if _ == sheet:
        print("Right!")
        break
else:
    print("Not Right!")
    exit()

print(f"{sheet} {week} {day} {work}")
test = 0
work_lst = [work]
if work == "9-11节":
    work_lst += ["9-10节", "7-10节"]

for i, n in enumerate(wb[sheet].values[0]):
    if n == "到课人数":
        p = i
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
    elif n == "选课人数":
        h = i
    elif n == "星期":
        m = i

for i in wb[sheet].values:
    if i[m] == x[day] and i[e] in work_lst:
        if i[g] == "线上授课":
            print(f"{i[f]} {i[a]} {i[b]} {i[c]} {i[d]}老师 线上授课")
            test = 1
        # elif i[g] == "到课率不足70%":
        #     print(f"{i[f]} {i[a]} {i[b]} {i[c]} {i[d]}老师 到课率不足70% 应到{i[h]}人 实到{i[p]}人")
        #     test = 1
        elif i[g] == "无师生":
            print(f"{i[f]} {i[a]} {i[b]} {i[c]} {i[d]}老师 无师生")
            test = 1
if test == 0:
    print("无异常情况")

print("其余教室均正常到课")

