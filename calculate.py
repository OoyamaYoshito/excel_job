# coding: UTF-8

import pandas as pd
import glob
import sys

#データを整形し、学年とクラス名が一致する情報のみを出力する
def search(age, class_name, studentpath="studentlist", answerpath="answersdata"):
    input_files = glob.glob(answerpath+"/*.xls") + glob.glob(answerpath+"/*.xlsx")
    input_files.sort()
    df_answerdata = []
    for i, filename in enumerate(input_files):
        print (filename)
        input_book = pd.ExcelFile(filename)
        #sheet_namesメソッドでExcelブック内の各シートの名前をリストで取得できる
        input_sheet_name = input_book.sheet_names

        #シートの１番目をDataflameに変換
        input_sheet_df = input_book.parse(input_sheet_name[0])

        #indexを学籍番号に変更
        #index名に学籍番号（は不要）が含まれるので削除したい
        
        student_number_cell=[s for s in list(input_sheet_df.columns.values) if "学籍番号" in s]
        skip=0
        #「学籍番号」を含むセルがなかったら、見つけるまで行を読み飛ばす
        while len(student_number_cell) == 0:
            skip = skip + 1
            input_sheet_df = input_book.parse(input_sheet_name[0], skiprows=skip)
            columns = [str(s) for s in input_sheet_df.columns.values]
            student_number_cell=[s for s in columns if "学籍番号" in s]
        df_answers = input_sheet_df.set_index(student_number_cell[0])
        #print(df_answers)

        #直書きしてるが、表を参照するようにしたい

        try:
            df_program=df_answers.loc[:,["# 回答1.55"]].mean(axis='columns') 
            df_know_1 =df_answers.loc[:,["# 回答1.6","# 回答1.19","# 回答1.24","# 回答1.38","# 回答1.44","# 回答1.54"]].mean(axis='columns') 
            df_know_2 =df_answers.loc[:,["# 回答1.9","# 回答1.13"]].mean(axis='columns') 
            df_know_3 =df_answers.loc[:,["# 回答1.46","# 回答1.51"]].mean(axis='columns') 
            df_know_4 =df_answers.loc[:,["# 回答1.21","# 回答1.31","# 回答1.36","# 回答1.39","# 回答1.43","# 回答1.48"]].mean(axis='columns') 
            df_atti_1 =df_answers.loc[:,["# 回答1.4","# 回答1.16","# 回答1.22","# 回答1.26","# 回答1.53"]].mean(axis='columns') 
            df_atti_2 =df_answers.loc[:,["# 回答1.17","# 回答1.34","# 回答1.52"]].mean(axis='columns') 
            df_atti_3 =df_answers.loc[:,["# 回答1.5","# 回答1.50"]].mean(axis='columns') 
        except KeyError:
            df_program=df_answers.loc[:,["# 回答1.31"]].mean(axis='columns') 
            df_know_1 =df_answers.loc[:,["# 回答1.4","# 回答1.9","# 回答1.14","# 回答1.20","# 回答1.25","# 回答1.29"]].mean(axis='columns') 
            df_know_2 =df_answers.loc[:,["# 回答1.11","# 回答1.17","# 回答1.23"]].mean(axis='columns') 
            df_know_3 =df_answers.loc[:,["# 回答1.5","# 回答1.19","# 回答1.27","# 回答1.30"]].mean(axis='columns') 
            df_know_4 =df_answers.loc[:,["# 回答1.7","# 回答1.10","# 回答1.13","# 回答1.16","# 回答1.22","# 回答1.28"]].mean(axis='columns') 
            df_atti_1 =df_answers.loc[:,["# 回答1.1","# 回答1.6","# 回答1.12","# 回答1.18","# 回答1.26"]].mean(axis='columns') 
            df_atti_2 =df_answers.loc[:,["# 回答1.2","# 回答1.8","# 回答1.24"]].mean(axis='columns') 
            df_atti_3 =df_answers.loc[:,["# 回答1.3","# 回答1.15","# 回答1.21"]].mean(axis='columns') 
        df_answers = pd.concat([df_program,df_know_1,df_know_2,df_know_3,df_know_4,df_atti_1,df_atti_2,df_atti_3],axis=1)

        df_answers.columns=["プログラミング得意度","構想・設計","エラーメッセージ理解","デバッグ","文法知識","積極性","他者活用","Web活用"]
        df_answerdata.append(df_answers)

    if not df_answerdata:
        print ("アンケート回答データが取得できません")
        sys.exit()
    df_answer = pd.concat(df_answerdata)

    #学生のリストを読み込む
    input_files =  glob.glob(studentpath + "/*.xlsx")
    df_studentlists = []
    for i, filename in enumerate(input_files):
        input_book = pd.ExcelFile(filename)

        #sheet_namesメソッドでExcelブック内の各シートの名前をリストで取得できる
        input_sheet_name = input_book.sheet_names

        #シートの１番目をDataflameに変換
        input_sheet_df = input_book.parse(input_sheet_name[0])

        #学生情報のdataframeを作成
        try:
            df_studentlists.append(input_sheet_df.set_index("ID"))
        except KeyError:
            df_studentlists.append(input_sheet_df.set_index("学籍番号"))

    if not df_studentlists:
        print ("学生情報データ(studentlist)が取得できません")
        sys.exit()
    df_studentlist = pd.concat(df_studentlists)

    #指定されたクラスのデータを引き出す
    df_classdata = df_studentlist.query("(Grade == @age)&(Class == @class_name)")
    if df_classdata.index.size == 0:
        class_name = age + class_name
        df_classdata = df_studentlist.query("(Grade == @age)&(Class == @class_name)")
    df_classdata = df_classdata.drop(columns=["Absence","Ent.year"])
    df_classdata = df_classdata.drop(columns=["Name(J_Kana)","Name(E) "],errors='ignore')
    df_classdata = df_classdata.drop(columns=["Name\n(J Kana)","Name(E)"],errors='ignore')
    df_classdata = df_classdata.drop(columns=["Name_J kana","Name_E"],errors='ignore')
    df_classdata = df_classdata.iloc[:,[0,1,3,4,2]]
    print(df_classdata)

    #アンケート結果用の数字データをもつ配列
    df_combineds = []
    df_mean = []
    for i, df_answer in enumerate(df_answerdata):
        df_combined = pd.concat([df_answer,df_classdata], axis=1, join="inner")
        df_combined =df_combined.drop(columns=["Name(J)"],errors='ignore')
        df_combined =df_combined.drop(columns=["Name_J"],errors='ignore')
        df_combined =df_combined.drop(columns=["Sex","Dept. & Course","Grade","Class"])
        if len(df_combined.index) > 0:
            df_mean = df_combined.mean()
        df_combineds.append(df_combined)
    print(df_mean)

    #excelファイルとして出力
    #df_output.to_excel("output.xlsx")

    return df_classdata, df_combineds, df_mean

if __name__ == "__main__":
    print(search('2','Z')) 
