import requests
import xlwings as xw
from xlwings.utils import rgb_to_int
import datetime
from bs4 import BeautifulSoup



wb = xw.Book()
sheet = wb.sheets[0]
sheet['A1'].value = "IMDB Most Popular TV Shows"
sheet['B1'].value = datetime.date.today()
title = ['名稱', '評分']
sheet['A2'].expand().value = title


source = requests.get('https://www.imdb.com/chart/tvmeter?sort=us,desc&mode=simple&page=1')
soup = BeautifulSoup(source.text, 'html.parser')
tvs = soup.find('tbody', class_="lister-list").find_all('tr')

row = 3  # 從第 3 行開始寫入數據

for tv in tvs:
    title = tv.find('td', class_="titleColumn").a.text
    rating = float(tv.find('td', class_="ratingColumn imdbRating").text)
    sheet.range("A" + str(row)).value = [title, rating]
    row += 1

for cell in sheet['B3'].expand('down'):
    if cell.value >= 8.0:
        cell.api.Font.Color = rgb_to_int((255,0,0))
    else:
        cell.api.Font.Color = rgb_to_int((0,0,0))

for i, row in enumerate(sheet.range('A3:B3').expand('down').rows):
    if i % 2 == 0:
        # 设置灰色填充
        row.color = (192, 192, 192)
    else:
        # 设置白色填充
        row.color = (255, 255, 255)

sheet['A1'].expand().api.HorizontalAlignment = 3
sheet.autofit()
wb.save('imdb.xlsx')