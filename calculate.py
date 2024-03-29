# coding: UTF-8

import pandas as pd
import glob
import sys
import datetime
import re

#データを整形し、学年とクラス名が一致する情報のみを出力する
def search(age, class_name, studentpath="studentlist", answerpath="answersdata"):
    df_answerdata = get_answerdata(answerpath)
    df_classdata = get_classstudents(age, class_name, studentpath)
    df_combineds, df_mean = calc_mean(df_answerdata, df_classdata)
    return df_classdata, df_combineds, df_mean

### アンケート回答データを収集する
def get_answerdata_from_xls(input_files):
    input_files.sort()
    df_answerdata = []
    for i, filename in enumerate(input_files):
        print (filename)
        input_book = pd.ExcelFile(filename)
        #sheet_namesメソッドでExcelブック内の各シートの名前をリストで取得できる
        input_sheet_name = input_book.sheet_names

        #シートの１番目をDataflameに変換
        input_sheet_df = input_book.parse(input_sheet_name[0])

        student_number_cell=[s for s in list(input_sheet_df.columns.values) if "学籍番号" in s]
        skip=0
        #「学籍番号」を含むセルがなかったら、見つけるまで行を読み飛ばす
        while len(student_number_cell) == 0:
            skip = skip + 1
            input_sheet_df = input_book.parse(input_sheet_name[0], skiprows=skip)
            columns = [str(s) for s in input_sheet_df.columns.values]
            student_number_cell=[s for s in columns if "学籍番号" in s]
        df_answers = input_sheet_df.set_index(student_number_cell[0])

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
        yearmonth=re.sub('^[^/]*/','',filename)[:6]
        df_answerdata.append((yearmonth[:4]+'/'+yearmonth[5:],df_answers))
        #print(df_answers)
    return df_answerdata

def get_answerdata_from_csv(input_files):
    input_files.sort()
    df_answerdata = []
    for i, filename in enumerate(input_files):
        print (filename)
        input_csv = pd.read_csv(filename)
        input_csv['学籍番号']=input_csv['ユーザ名'].replace('b([0-9]+)',r'\1',regex=True)
        input_csv['学籍番号']=pd.to_numeric(input_csv['学籍番号'], errors='coerce')
        #print(input_csv)
        df_answers = input_csv.set_index('学籍番号')
        #print(df_answers)
        df_program=df_answers.loc[:,["Q31_31) プログラミングが得意である"]].mean(axis='columns') 
        df_know_1 =df_answers.loc[:,["Q04_4) 仕様書を読んでプログラムを作成するとき，プログラムの完","Q09_9) 類似のプログラムがなくても，一からプログラムを記述でき","Q14_14) 順を追って論理的にプログラムを記述できる","Q20_20) プログラムを読むことで，大まかな動作を想像できる","Q25_25) プログラムを読むことにより，そのプログラムがどのよう","Q29_29) プログラムが思ったとおりに動作しないとき，別のやり方"]].mean(axis='columns') 
        df_know_2 =df_answers.loc[:,["Q11_11) エラーメッセージの英語の内容を理解できる","Q17_17) エラーメッセージが表示されたとき，そのエラーメッセー","Q23_23) エラーメッセージを読むと，どのような問題が起こってい"]].mean(axis='columns') 
        df_know_3 =df_answers.loc[:,["Q05_5) 正しく動作しないとき，正しく動作するまでプログラムを少","Q19_19) プログラムが正しく動作しないとき，どこにバグがあるの","Q27_27) プログラムが正しく動作しないとき，プログラムのロジッ","Q30_30) プログラムが正しく動作しないとき，変数の値などを変更"]].mean(axis='columns') 
        df_know_4 =df_answers.loc[:,["Q07_7) 実行文（代入，条件分岐，繰り返し）を使うべき場所で正し","Q10_10) 複数の実行文（代入，条件分岐，繰り返し）を組み合わせ","Q13_13) 配列を使うべき場所で正しく使える","Q16_16) 関数(サブプログラム，ブロック，メソッドなど）を使う","Q22_22) ライブラリ（ライブラリ関数，外部クラスなど）を使うべ","Q28_28) 機能が似ている実行文（代入，条件分岐，繰り返し）の違"]].mean(axis='columns') 
        df_atti_1 =df_answers.loc[:,["Q01_1) プログラミングを積極的に勉強している","Q06_6) 講義以外でも，独学でプログラミングを勉強する","Q12_12) 難しいプログラミングにも挑戦する","Q18_18) プログラミングを学習するとき，さまざまなプログラム作","Q26_26) プログラミングの経験を積むために，さまざまプログラム"]].mean(axis='columns') 
        df_atti_2 =df_answers.loc[:,["Q02_2) プログラムが正しく動かないときは，友人に相談する","Q08_8) プログラミングを学習するとき，友人と協力する","Q24_24) プログラムが正しく動かないときは，TAやチューターに"]].mean(axis='columns') 
        df_atti_3 =df_answers.loc[:,["Q03_3) プログラミングを学習するとき，Web上にある情報やリフ","Q15_15) プログラムが正しく動かないとき，Web上にある情報や","Q21_21) エラーメッセージの意味がわからないときは，辞書を使っ"]].mean(axis='columns') 
        df_answers = pd.concat([df_program,df_know_1,df_know_2,df_know_3,df_know_4,df_atti_1,df_atti_2,df_atti_3],axis=1)
        df_answers.columns=["プログラミング得意度","構想・設計","エラーメッセージ理解","デバッグ","文法知識","積極性","他者活用","Web活用"]
        yearmonth=re.sub('^[^/]*/','',filename)[:6]
        df_answerdata.append((yearmonth[:4]+'/'+yearmonth[5:],df_answers))
        #print(df_answers)
    return df_answerdata

def get_answerdata(answerpath="answersdata"):
    today=datetime.date.today()
    year=today.year
    if today.month < 4:
        year=year-1
    yearre="/" + str(year)
    prevre="/" + str(year-1)
    nextre="/" + str(year+1)
    input_files = glob.glob(answerpath+prevre+"0[4-9]*.xls*") \
        + glob.glob(answerpath+prevre+"0[4-9]*/*.xls*") \
        + glob.glob(answerpath+prevre+"1[0-2]*.xls*") \
        + glob.glob(answerpath+prevre+"1[0-2]*/*.xls*") \
        + glob.glob(answerpath+yearre+"0[1-3]*.xls*") \
        + glob.glob(answerpath+yearre+"0[1-3]*/*.xls*") \
        + glob.glob(answerpath+yearre+"0[4-9]*.xls*") \
        + glob.glob(answerpath+yearre+"0[4-9]*/*.xls*") \
        + glob.glob(answerpath+yearre+"1[0-2]*.xls*") \
        + glob.glob(answerpath+yearre+"1[0-2]*/*.xls*") \
        + glob.glob(answerpath+nextre+"0[1-3]*.xls*") \
        + glob.glob(answerpath+nextre+"0[1-3]*/*.xls*")
    df_answerdata = get_answerdata_from_xls(input_files)
    print(len(df_answerdata))

    # support Moodleアンケート
    input_files = glob.glob(answerpath+prevre+"0[4-9]*/*.csv") \
        + glob.glob(answerpath+prevre+"1[0-2]*/*.csv") \
        + glob.glob(answerpath+yearre+"0[1-3]*/*.csv") \
        + glob.glob(answerpath+yearre+"0[4-9]*/*.csv") \
        + glob.glob(answerpath+yearre+"1[0-2]*/*.csv") \
        + glob.glob(answerpath+nextre+"0[1-3]*/*.csv")
    df_answerdata.extend(get_answerdata_from_csv(input_files))
    print(len(df_answerdata))

    if not df_answerdata:
        print ("01アンケート回答データが取得できません")
        raise ValueError("01アンケート回答データが取得できません")
    #df_answer = pd.concat(df_answerdata)
    return df_answerdata

### 指定された学生の学生情報を収集する
def get_studentdata(students, studentpath="studentlist"):
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
        print ("02学生情報データ(studentlist)が取得できません")
        raise ValueError("02学生情報データ(studentlist)が取得できません")
    df_studentlist = pd.concat(df_studentlists)

    #指定されたクラスのデータを引き出す
    try:
        df_classdata = df_studentlist.query("(ID in @students)")
    except pd.core.computation.ops.UndefinedVariableError:
        df_classdata = df_studentlist.query("(学籍番号 in @students)")
    df_classdata = df_classdata.drop(columns=["Absence","Ent.year"])
    df_classdata = df_classdata.drop(columns=["Name(J_Kana)","Name(E) "],errors='ignore')
    df_classdata = df_classdata.drop(columns=["Name(J Kana)","Name(E)"],errors='ignore')
    df_classdata = df_classdata.drop(columns=["Name\n(J Kana)","Name(E)"],errors='ignore')
    df_classdata = df_classdata.drop(columns=["Name_J kana","Name_E"],errors='ignore')
    df_classdata = df_classdata.iloc[:,[0,1,3,4,2]]
    print(df_classdata)
    return df_classdata


### 指定された学年/クラスの学生情報を収集する
def get_classstudents(age, class_name, studentpath="studentlist"):
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
        print ("02学生情報データ(studentlist)が取得できません")
        raise ValueError("02学生情報データ(studentlist)が取得できません")
    df_studentlist = pd.concat(df_studentlists)

    #指定されたクラスのデータを引き出す
    df_classdata = df_studentlist.query("(Grade == @age)&(Class == @class_name)")
    if df_classdata.index.size == 0:
        class_name = age + class_name
        df_classdata = df_studentlist.query("(Grade == @age)&(Class == @class_name)")
    df_classdata = df_classdata.drop(columns=["Absence","Ent.year"])
    df_classdata = df_classdata.drop(columns=["Name(J_Kana)","Name(E) "],errors='ignore')
    df_classdata = df_classdata.drop(columns=["Name(J Kana)","Name(E)"],errors='ignore')
    df_classdata = df_classdata.drop(columns=["Name\n(J Kana)","Name(E)"],errors='ignore')
    df_classdata = df_classdata.drop(columns=["Name_J kana","Name_E"],errors='ignore')
    df_classdata = df_classdata.iloc[:,[0,1,3,4,2]]
    #print(df_classdata)
    return df_classdata

### 対象学生内の平均を計算
def calc_mean(df_answerdata,df_classdata):
    #アンケート結果用の数字データをもつ配列
    df_combineds = []
    df_mean = []
    for i, df_answerym in enumerate(df_answerdata):
        (ym,df_answer)=df_answerym
        try:
            df_combined = pd.concat([df_answer,df_classdata], axis=1, join="inner")
            df_combined =df_combined.drop(columns=["Name(J)"],errors='ignore')
            df_combined =df_combined.drop(columns=["Name_J"],errors='ignore')
            df_combined =df_combined.drop(columns=["Sex","Dept. & Course","Grade","Class"],errors='ignore')
            if len(df_combined.index) > 0:
                df_mean = (ym,df_combined.mean())
            df_combineds.append((ym,df_combined))
        except ValueError:
            print('同一学生の重複回答があるようです: ' + ym)
    #print(df_mean)

    #excelファイルとして出力
    #df_output.to_excel("output.xlsx")

    return df_combineds, df_mean

if __name__ == "__main__":
    print(search('2','Z')) 
