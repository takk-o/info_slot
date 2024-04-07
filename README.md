# info_slot

## Overview
起動時に引数に指定したサイト(https://ana-slo.com/*)よりスロット情報を収集し、Excelファイルに出力する。

## Requirements
- Python 3.12.2
- requests 2.31.0
- bs4 0.0.2
- openpyxl 3.1.2

## Usage
1. 引数に情報収集したいサイトのURLをセットしてinfo_slot.pyを起動
1. log'フォルダのlog.txtに'エラー'または'正常終了'の情報が出力される
1. 'output'フォルダにinfo_slot.xlsxが作成される

## Executable File Creation Procedure
1. pip install pyinstaller
1. pyinstaller info_slot.py --onefile --noconsole

## Automatic Execution Environment Construction
1. dist/info_slotモジュール格納ディレクトリに、info_slot.commandを配置し、実行権限を付与
1. crontab.txtの内容を参考に、crontabに起動時刻を設定

## Author
- takk-o
- Mail : ynurmj5e@gmail.com
