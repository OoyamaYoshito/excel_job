# coding: UTF-8

import pandas as pd
import glob
import sys

#データを整形し、学年とクラス名が一致する情報のみを出力する
def search(age, class_name):
    input_files = glob.glob("answersdata/*.xlsx")
    df_answerdata = []
    for i, filename in enumerate(input_files):
        input_book = pd.ExcelFile(filename)
        #sheet_namesメソッドでExcelブック内の各シートの名前をリストで取得できる
        input_sheet_name = input_book.sheet_names

        #シートの１番目をDataflameに変換
        input_sheet_df = input_book.parse(input_sheet_name[0])

        #indexを学籍番号に変更
        #index名に学籍番号（は不要）が含まれるので削除したい
        
        student_number_cell=[s for s in list(input_sheet_df.columns.values) if "学籍番号" in s]
        df_answers = input_sheet_df.set_index(student_number_cell[0])

        #直書きしてるが、表を参照するようにしたい

        df_program=df_answers.ix[:,["27 プログラミングが得意である"]].mean(axis='columns') 
        df_know_1 =df_answers.ix[:,["03 仕様書を読んでプログラムを作成するとき，プログラムの完成形を想像できる","08 類似のプログラムがなくても，一からプログラムを記述できる","11 順を追って論理的にプログラムを記述できる","16 プログラムを読むことで，大まかな動作を想像できる","19 プログラムを読むことにより，そのプログラムがどのような処理・動作を行うのかを把握する","26 プログラムが思ったとおりに動作しないとき，別のやり方を思いつく"]].mean(axis='columns') 
        df_know_2 =df_answers.ix[:,["04 エラーメッセージの英語の内容を理解できる","05 エラーメッセージが表示されたとき，そのエラーメッセージの内容を理解できる"]].mean(axis='columns') 
        df_know_3 =df_answers.ix[:,["20 プログラムが正しく動作しないとき，変数の値などを変更してみて動作を確認する","23 正しく動作しないとき，正しく動作するまでプログラムを少しずつ変更して動かしてみる"]].mean(axis='columns') 
        df_know_4 =df_answers.ix[:,["09 実行文（代入，条件分岐，繰り返し）を使うべき場所で正しく使える","13 複数の実行文（代入，条件分岐，繰り返し）を組み合わせて使える","15 配列を使うべき場所で正しく使える","17 ライブラリ（ライブラリ関数，外部クラスなど）を使うべき場所で正しく使える","18 機能が似ている実行文（代入，条件分岐，繰り返し）の違いを説明できる","21 関数(サブプログラム，ブロック，メソッドなど）を使うべき場所で正しく使える"]].mean(axis='columns') 
        df_atti_1 =df_answers.ix[:,["01 プログラミングを積極的に勉強している","06 難しいプログラミングにも挑戦する","10 講義以外でも，独学でプログラミングを勉強する","12 プログラミングを学習するとき，さまざまなプログラム作成に自分で挑戦する","25 プログラミングの経験を積むために，さまざまプログラムを作成する"]].mean(axis='columns') 
        df_atti_2 =df_answers.ix[:,["07 プログラムが正しく動かないときは，友人に相談する","14 プログラミングを学習するとき，友人と協力する","24 プログラムが正しく動かないときは，TAやチューターに相談する"]].mean(axis='columns') 
        df_atti_3 =df_answers.ix[:,["02 プログラミングを学習するとき，Web上にある情報やリファレンスを利用する","22 プログラムが正しく動かないとき，Web上にある情報やリファレンスを利用する"]].mean(axis='columns') 
        df_answers = pd.concat([df_program,df_know_1,df_know_2,df_know_3,df_know_4,df_atti_1,df_atti_2,df_atti_3],axis=1)

        df_answers.columns=["プログラミング得意度","構想・設計","エラーメッセージ理解","デバッグ","文法知識","積極性","他者活用","Web活用"]
        #df_answers["日時"]=filename.lstrip("answersdata/").rstrip(".xlsx")
        df_answerdata.append(df_answers)

    if not df_answerdata:
        print ("アンケート回答データが取得できません")
        sys.exit()
    df_answer = pd.concat(df_answerdata)

    #学生のリストを読み込む
    input_files =  glob.glob("studentlist/*.xlsx")
    df_studentlists = []
    for i, filename in enumerate(input_files):
        input_book = pd.ExcelFile(filename)

        #sheet_namesメソッドでExcelブック内の各シートの名前をリストで取得できる
        input_sheet_name = input_book.sheet_names

        #シートの１番目をDataflameに変換
        input_sheet_df = input_book.parse(input_sheet_name[0])

        #学生情報のdataframeを作成
        df_studentlists.append(input_sheet_df.set_index("ID"))

    df_studentlist = pd.concat(df_studentlists)

    #指定されたクラスのデータを引き出す
    df_classdata = df_studentlist.query("(Grade == @age)&(Class == @class_name)")
    df_classdata = df_classdata.drop(columns=["Name(J_Kana)","Name(E) ","Absence","Ent.year"])
    df_classdata = df_classdata.ix[:,[0,1,3,4,2]]

    #アンケート結果用の数字データをもつ配列
    df_combineds = []
    for i, df_answer in enumerate(df_answerdata):
        df_combined = pd.concat([df_answer,df_classdata], axis=1, join="inner")
        df_combined =df_combined.drop(columns=["Name(J)","Sex","Dept. & Course","Grade","Class"])
        df_combineds.append(df_combined)

    #excelファイルとして出力
    #df_output.to_excel("output.xlsx")

    return df_classdata, df_combineds

if __name__ == "__main__":
    print(search('2','Z')) 