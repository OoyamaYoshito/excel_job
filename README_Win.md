# excel_jobツールを使い始めるまで[Windows編]


## Python3をインストール
- 以下の辺りからダウンロードしてインストール
  https://www.python.org/downloads/windows/
- インストールが完了したら、コマンドプロンプトを起動して、以下のコマンド実行で想定したバージョンのPythonが使えるようになったか確認
  `python -V`
- うまくいっていない場合は PATH の設定を確認(PATHの中にPythonぽいのが入っているか)

## Pythonのパッケージをインストール
- コマンドプロンプトで以下のコマンドを実行し、インストール済みのパッケージを確認
  `pip list`
- 下記のパッケージがすべて揃うまで、`pip install パッケージ名` を続ける

  pandas
  xlrd
  openpyxl
  Pillow
  numpy
  matplotlib

## Excelデータの準備
- 未来大事務局の学生名簿Excelをstudentlistフォルダの中に置く
	- 置くファイルは一つだけ、ファイル名は何でもよい
- アンケート回答Excel(複数可)をanswerdataフォルダの中に置く
	- ファイル名はアンケート実施年月を含む 202004.xlsx などのようにする

## 実行
- コマンドプロンプトで以下のように実行
  `python output.py 学年 クラス`
  (学年は1か2、クラスはAからZのどれか一文字)
- 同じフォルダに output.xlsx が生成されるので、中身を確認

### テンプレの確認
- 各シートの元データは template.xlsx に入っています。説明文などの軽微な変更はここで

### グラフの文字が化ける場合...
- グラフ生成に用いているパッケージから

