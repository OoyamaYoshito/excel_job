import pandas as pd
import sys

args = sys.argv
age = args[1]
clas = args[2]


input_book = pd.ExcelFile('answersdata.xlsx')
 
#sheet_namesメソッドでExcelブック内の各シートの名前をリストで取得できる
input_sheet_name = input_book.sheet_names

#シートの１番目をDataflameに変換
input_sheet_df = input_book.parse(input_sheet_name[0])

#indexを学籍番号に変更
#index名に学籍番号（は不要）が含まれるので削除したい
df_answer = input_sheet_df.set_index("学籍番号（は不要）")

#直書きしてるが、表を参照するようにしたい
#特定の数字で始まったものを出力するようにしたい

df_know_1=df_answer.ix[:,["03 仕様書を読んでプログラムを作成するとき，プログラムの完成形を想像できる","08 類似のプログラムがなくても，一からプログラムを記述できる","11 順を追って論理的にプログラムを記述できる","16 プログラムを読むことで，大まかな動作を想像できる","19 プログラムを読むことにより，そのプログラムがどのような処理・動作を行うのかを把握する","26 プログラムが思ったとおりに動作しないとき，別のやり方を思いつく"]].mean(axis='columns') 
df_know_2=df_answer.ix[:,["04 エラーメッセージの英語の内容を理解できる","05 エラーメッセージが表示されたとき，そのエラーメッセージの内容を理解できる"]].mean(axis='columns') 
df_know_3=df_answer.ix[:,["20 プログラムが正しく動作しないとき，変数の値などを変更してみて動作を確認する","23 正しく動作しないとき，正しく動作するまでプログラムを少しずつ変更して動かしてみる"]].mean(axis='columns') 
df_know_4=df_answer.ix[:,["09 実行文（代入，条件分岐，繰り返し）を使うべき場所で正しく使える","13 複数の実行文（代入，条件分岐，繰り返し）を組み合わせて使える","15 配列を使うべき場所で正しく使える","17 ライブラリ（ライブラリ関数，外部クラスなど）を使うべき場所で正しく使える","18 機能が似ている実行文（代入，条件分岐，繰り返し）の違いを説明できる","21 関数(サブプログラム，ブロック，メソッドなど）を使うべき場所で正しく使える"]].mean(axis='columns') 
df_atti_1=df_answer.ix[:,["01 プログラミングを積極的に勉強している","06 難しいプログラミングにも挑戦する","10 講義以外でも，独学でプログラミングを勉強する","12 プログラミングを学習するとき，さまざまなプログラム作成に自分で挑戦する","25 プログラミングの経験を積むために，さまざまプログラムを作成する"]].mean(axis='columns') 
df_atti_2=df_answer.ix[:,["07 プログラムが正しく動かないときは，友人に相談する","14 プログラミングを学習するとき，友人と協力する","24 プログラムが正しく動かないときは，TAやチューターに相談する"]].mean(axis='columns') 
df_atti_3=df_answer.ix[:,["02 プログラミングを学習するとき，Web上にある情報やリファレンスを利用する","22 プログラムが正しく動かないとき，Web上にある情報やリファレンスを利用する"]].mean(axis='columns') 
df_answer = pd.concat([df_know_1,df_know_2,df_know_3,df_know_4,df_atti_1,df_atti_2,df_atti_3],axis=1)

df_answer.columns=["構想・設計","エラーメッセージ理解","デバッグ","文法知識","積極性","他者活用","Web活用"]

input_book = pd.ExcelFile('studentlist.xlsx')

#sheet_namesメソッドでExcelブック内の各シートの名前をリストで取得できる
input_sheet_name = input_book.sheet_names

#シートの１番目をDataflameに変換
input_sheet_df = input_book.parse(input_sheet_name[0])

#学生情報のdataframeを作成
df_studentlist = input_sheet_df.set_index("ID").drop_duplicates()

#回答結果と結合
df_combined = pd.concat([df_answer,df_studentlist], axis=1, join="inner")

df_output = df_combined.query("(Grade == @age)&(Class == @clas)")

print(df_output)
#excelファイルとして出力
df_output.to_excel("output.xlsx")