# coding: UTF-8

import openpyxl
import numpy as np
import pandas as pd
from PIL import Image
import calculate
import chart_python
import matplotlib.font_manager

#一人分のアンケート結果を出力
def question_result_output(personal_info, answerdata):
    wb = openpyxl.load_workbook('templete.xlsx')
    ws = wb.worksheets[0]

    #個人情報挿入
    for i, data in enumerate(personal_info):
        ws.cell(4,1+i,data)
    #ws.cell(4,1,personal_info)
    
    #数字データ結果挿入
    for i, yeardata in enumerate(answerdata):
        for j,data in enumerate(yeardata):
            ws.cell(7+i,2+j,data)

    #グラフ画像挿入
    chart_python.draw_chart(answerdata)
    img = openpyxl.drawing.image.Image('graph.png')
    ws.add_image( img, 'B13' )

    wb.save('output.xlsx')

if __name__ == "__main__":
    classdata,answerdatas=calculate.search("2","Z")
    student_number = 1018094
    personal_info = classdata.loc[student_number].values
    personal_info = np.insert(personal_info,0,student_number)
    answerdata = []
    for yeardata in answerdatas:
        answerdata.append(yeardata.loc[student_number].values)
    for _ in range(4-len(answerdata)):
        answerdata.append([0,0,0,0,0,0,0,0]) 
    print(personal_info,answerdata)
    question_result_output(personal_info,answerdata)