# coding: UTF-8

import sys
import openpyxl
import numpy as np
import pandas as pd
from PIL import Image
import calculate
import chart_python
import matplotlib.font_manager

#一人分のアンケート結果を出力
def question_result_output(personal_info, answerdata,ws):

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

#一クラス分のExcel生成
def output_class(stgrade, stclass):
    classdata,answerdatas=calculate.search(stgrade,stclass)
    fn = 'output' + stgrade + stclass + '.xlsx'
    wb = openpyxl.load_workbook('templete.xlsx')
    ws = wb.worksheets[0]
    if classdata.index.size == 0:
        print ("クラスのデータが取得できませんでした")
    for sheet_name in list(classdata.index):
        ws_copy = wb.copy_worksheet(ws)
        ws_copy.title=str(sheet_name)
    wb.save(fn)
    for i, student_number in enumerate(list(classdata.index)):
        wb = openpyxl.load_workbook(fn)
        personal_info = classdata.loc[student_number].values
        personal_info = np.insert(personal_info,0,student_number)
        answerdata = []
        for yeardata in answerdatas:
            if student_number in list(yeardata.index):
                answerdata.append(yeardata.loc[student_number].values)
        for _ in range(4-len(answerdata)):
            answerdata.append([0,0,0,0,0,0,0,0]) 
        question_result_output(personal_info,answerdata,wb.worksheets[i+1])
        wb.save(fn)

if __name__ == "__main__":
    if len(sys.argv) == 3:
        args = sys.argv
        stgrade = args[1]
        stclass = args[2]
    else:
        print ("Usage: python "+sys.argv[0]+" <学年> <クラス>")
        sys.exit()
    for g in stgrade:
        for c in stclass:
            output_class(g, c)
