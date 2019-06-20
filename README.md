Report working hours
==

# 概要

このツールは作業実績報告書等を添付してメール送信することができます。

報告を時短するために作成しました。

## 動作に必要なもの

本体プログラムはPythonでコーディングしているため、Pythonの実行環境が必要です。

+ [Python](https://www.python.org/)

WindowsアプリケーションのAPIを使用するため、以下のライブラリが必要です。

+ [pywin32](https://github.com/mhammond/pywin32)

リアルタイムにキーボード入力を処理するため、以下のライブラリが必要です。

+ [readchar](https://pypi.org/project/readchar/)

## プログラムのダウンロード
`git clone https://github.com/interceptor128/report_working_hours.git`  
でレポジトリをクローン(ローカルPCにダウンロード)できます。

## 使い方

宛先(To,Cc,Bcc)の設定をテキストファイルに記載してください。
設定例は`sample.conf`にあります。

以下のコマンド入力でメール送信画面を表示またはメール送信します。
```command:sample
python report_working_hours.py sample.conf 資料.xls
```
 
設定ファイルと添付ファイルはpyファイルと同じディレクトリ（フォルダ）に入れて指定してください。

必要に応じて、ソース内の件名、メール本文を編集してください。  
宛先の名前、差出人の名前等

## 課題
+ To Cc Bcc には1アドレスしか指定できないので改善したい
+ 件名、メール本文の変更はソースコード内にハードコーディングしているためカスタマイズできない
    + 許容範囲内？