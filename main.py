import xlwings as xls
import os
import pandas as pd
import re
import config # config文件视人数而定，其中columnList列表的元素由转置之后的列数而定。简而言之，从字母“B”开始，列表长度等于班级人数。

# findCreditInStr函数：从课程标题中提取本课程的学分
def findCreditInStr (String):
    rule = "\/\.?\d\.?\d?"
    findList = re.findall (rule, String)
    rawCredit = findList[0]
    rawCredit = rawCredit.replace ('/','')
    credit = float (rawCredit)
    return credit

# handleGrade函数：将分数数据转化成标准五分制绩点
def handleGrade (grade):
    rule = "-?\d{1,}\.\d{1,}E-?\d{1,}|-?\d{1,}\.\d{1,}|-?\d{1,}"
    gradeText = str (grade)
    if (gradeText == "优秀"):
        gradeText = "95"
    elif (gradeText == "良好"):
        gradeText = "85"
    elif (gradeText == "中等"):
        gradeText = "75"
    elif (gradeText == "合格" or gradeText == "及格"):
        gradeText = "65"
    elif (gradeText == "不合格" or gradeText == "缺考" or gradeText == "不及格"):
        gradeText = "0"
    findList = re.findall (rule, gradeText)
    rawGrade = findList[0]
    rawGrade = float (rawGrade)
    if (rawGrade < 60):
        finalGrade = 0
    else:
        finalGrade = (rawGrade - 50) / 10
    return finalGrade

current_dir = os.path.dirname(os.path.abspath(__file__))
file_name_list = os.listdir (current_dir)

data = pd.read_excel ('2021-2023学年总表.xlsx') # 需要处理的原始数据文件，可自由替换，需保证main.py与xls文件在同一目录下。
data = data.T
data.to_excel ("1.xls",header = False) # 为方便操作和读写，将原表格进行了转置操作。1.xls为转置后的xls，GPA数据也会输出在此文件中。可自由命名。

wb = xls.Book ("1.xls")
sheet1 = wb.sheets["sheet1"]
titleList = []
index = 1
while (True):
    cellName = 'A' + str (index)
    cellValue = sheet1[cellName].value
    if (cellValue is not None):
        titleList.append (cellValue)
        index += 1
    else:
        break

len_titleList = len (titleList)
creditList = []
for i in range (3,len_titleList):
    creditList.append (findCreditInStr(titleList[i]))
for column in config.columnList:
    gradeList = []
    for i in range (4, len_titleList + 1):
        cellName = column + str (i)
        cellValue = handleGrade (sheet1[cellName].value) if sheet1[cellName].value is not None else None
        gradeList.append (cellValue)

    num, den = 0, 0
    for i in range (len_titleList - 3):
        if (gradeList[i] is not None):
            num += gradeList[i] * creditList[i]
            den += creditList[i]
    
    gpa = num / den
    print (gpa)

    targetCell = column + str (len_titleList + 1)

    sheet1.range (targetCell).value = gpa
    wb.save ()
