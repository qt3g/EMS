# Excel Splitter

このプログラムは、指定されたExcelファイルを月ごとに分割し、各月の収入と支出の合計を計算して新しいExcelファイルとして保存します。

## 必要条件

- Python 3.x
- pandas
- openpyxl
- tkinter

## インストール

必要なパッケージをインストールするには、以下のコマンドを実行してください：

```sh
pip install pandas openpyxl
```

## 使い方

1. プログラムを実行します。
2. 「令和何年」の入力欄に令和の年を入力します（例：令和3年の場合は「3」と入力）。
3. 「メニュー」から「ディレクトリを指定」を選択し、出力先ディレクトリを設定します。
4. 「Excelファイルを選択して月ごとに分割」ボタンをクリックし、分割したいExcelファイルを選択します。
5. 各月ごとに分割されたファイルが指定した出力先ディレクトリに保存されます。
6. 「ログをクリア」ボタンでログをクリアできます。
7. 「メニュー」から「ファイル名フォーマットを設定」を選択し、出力ファイルの名前フォーマットを指定できます（例：「{month}月_{year}.xlsx」）。

## ファイル構成

- `EMS.py`: メインプログラムファイル

## ライセンス

このプロジェクトはMITライセンスの下でライセンスされています。詳細については、`LICENSE`ファイルを参照してください。
