# keiba-csv-excel

v1.0 集計してエクセル出力
v1.1 集計データをグラフ出力して画像に保存

競馬のCSVデータを読み込んで整形・集計し、Excel、グラフに出力するPythonスクリプト。

## 使い方
1. 仮想環境を作成・有効化
   ```bash
   python -m venv venv
   source venv/bin/activate   # Mac/Linux
   venv\Scripts\activate      # Windows

2．必要なライブラリをインストール
pip install -r requirements.txt

3．実行
python main.py


サンプルデータ
sample_data/ にサンプルのCSVを置いてあるので、そのまま動かして試せます。

出力
output.xlsx が作成されます。
