import requests
from bs4 import BeautifulSoup
from pathlib import Path
import openpyxl

# excelブック/シートの準備
wb = openpyxl.Workbook()
ws = wb.active
ws.title = 'info_slot'
ws.cell(row=1, column=1, value='設置機種')

# 一次元配列を二次元配列に変換
def convert_1d_to_2d(l, cols):
    return [l[i:i + cols] for i in range(0, len(l), cols)]

tbl = dict()

url = 'https://ana-slo.com/2024-03-14-%E3%83%91%E3%83%BC%E3%83%AB%E3%82%B7%E3%83%A7%E3%83%83%E3%83%97%E3%81%A8%E3%82%82%E3%81%88%E7%A8%B2%E6%AF%9B%E9%95%B7%E6%B2%BC%E5%BA%97-data/'
soup = BeautifulSoup(requests.get(url).content, 'html.parser')

title_tags = soup.select('h4[id^=section]')
table_tags = soup.select('div[id^=tab01_]')
header_tags = table_tags[0].select('th')
detail_tags = table_tags[0].select('td')

counter = 1

for num in range(len(title_tags)):
    tbl['title'] = title_tags[num].text
    tbl['header'] = [header_tag.text for header_tag in header_tags]
    tbl['detail'] = convert_1d_to_2d([detail_tag.text for detail_tag in detail_tags], 8)
    # pprint(tbl)
    ws.cell(row=counter, column=1, value=tbl['title'])
    ws.cell(row=counter, column=2, value='台番号')
    ws.cell(row=counter, column=3, value='G数')
    ws.cell(row=counter, column=4, value='BB')
    ws.cell(row=counter, column=5, value='RB')
    ws.cell(row=counter, column=6, value='差枚')
    counter += 1
    for row in tbl['detail']:
        ws.cell(row=counter, column=2, value=row[0])
        ws.cell(row=counter, column=3, value=row[1])
        ws.cell(row=counter, column=4, value=row[3])
        ws.cell(row=counter, column=5, value=row[4])
        ws.cell(row=counter, column=6, value=row[2])
        counter += 1
    counter += 1

# フォルダーを作成
folder = Path('output')
folder.mkdir(exist_ok=True)

# excelファイルに出力
excel_path = folder.joinpath('info_slot.xlsx')
wb.save(excel_path)
wb.close()