# excel_job
エクセルファイルから学生のアンケート情報を出力するプログラムです

# 環境
- Python 3.6.7
- pandas 0.23.4
- xlrd 1.1.0
- openpyxl 2.5.11
- Pillow 5.4.1
- numpy 1.15.4
- matplotlib 3.0.2

# インストール方法
[PythonインストールWindows](https://www.python.jp/install/windows/install_py3.html)

[PythonインストールMac](https://qiita.com/ms-rock/items/6e4498a5963f3d9c4a67)

`$ pip install pandas`

`$ pip3 install xlrd`

`$ pip3 install openpyxl`

`$ pip3 install Pillow`

`$ pip3 install numpy`

`$ pip3 install matplotlib`

# 使い方
インストール方法に従って、Python3と必要なライブラリをインストールする

answersdata/ studentlist/　ディレクトリを作成する

studentlist/に学生情報が入ったxlsxファイルを1つだけ入れる（ファイル名は問わない）

answersdata/にアンケートの回答結果を入れる（ファイル名は201907.xlsx,202004.xlsxなど実施時期にする）

IPAexフォントの中から[IPAexゴシック](https://ipafont.ipa.go.jp/)をインストール

環境ごとのインストール方法についてはIPAexフォントの[インストール方法のページ](https://ipafont.ipa.go.jp/node72#jp)を参照すること
