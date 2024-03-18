# info_slot

## Overview
指定サイト(https://x.gd/wbhxu)より、スロット情報を収集し、Excelファイルに出力する。

## Requirements
- Python 3.11.8
- requests 2.31.0
- bs4 0.0.2
- openpyxl 3.1.2

## Usage
1. info_slot.pyを起動
1. 'output'フォルダにinfo_slot.xlsxが作成される

## Automatic Execution Environment Construction
1. dist/info_slotモジュール格納ディレクトリに、info_slot.commandを配置し、実行権限を付与
1. crontab.txtの内容を参考に、crontabに起動時刻を設定

## Author
- takk-o
- Mail : ynurmj5e@gmail.com
