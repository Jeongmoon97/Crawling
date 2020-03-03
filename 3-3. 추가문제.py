# 1. If 문을 이용해서 멜론 사이트에서 1위부터 50위 사이의 방탄소년단의 음악을 찾고 저장하기
# - 이 때, 순위, 노래 제목, 가수명 형태로 엑셀 파일에 저장 sheet1)
# https://www.melon.com/chart/index.htm
# - 가장 마지막 줄에 방탄소년단 음악이 총 몇 곡인지 표시 하기 (sheet1)
# 2. 방탄소년단 앨범 커버 이미지 저장
# bts 폴더를 만들고 저장하기 (이미지 제목은 곡제목으로 저장)

import requests
from bs4 import BeautifulSoup
import openpyxl
from urllib.request import urlretrieve

try:
    wb = openpyxl.load_workbook('bts.xlsx')
    sheet = wb.active
    print('불러오기 완료')
except:
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.append(['순위', '노래제목', '가수'])
    print('새로운 파일을 만들었습니다.')

raw= requests.get("https://www.melon.com/chart/index.htm",
                  headers={'User-Agent': 'Morzilla/5.0'})
html = BeautifulSoup(raw.text,'html.parser')
songs = html.select("tr.lst50")
count = 0

for s in songs:
    rank = s.select_one('div.wrap.t_center span.rank').text
    title = s.select_one('div.ellipsis.rank01 a').text
    singer = s.select_one('div.ellipsis.rank02 > a').text
    info = s.select_one('div:nth-of-type(5) a')
    if '방탄소년단' not in singer:
        continue
    count += 1
    sheet.append([rank, title, singer])
    # 왜 오류???
    url = 'https://www.melon.com/song/detail.htm?songId='+ info.attrs['href'][36:44]
    raw_each = requests.get(url, headers={'User-Agent': 'Morzilla/5.0'})
    html_each = BeautifulSoup(raw_each.text, 'html.parser')
    poster = html_each.select_one('div.thumb img')
    poster_src = poster.attrs['src']
    urlretrieve(poster_src, 'bts/'+title[:2]+'.png')

sheet.append(['방탄소년단 노래',count])
wb.save('bts.xlsx')