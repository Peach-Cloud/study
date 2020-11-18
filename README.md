# 구글 스프레드시트 견적서 자동화

#-*-coding:utf-8-*-
#pip install gspread, pip install --upgrade oauth2client <- 파이참 Terminal에서 설치 후 사용 가능
import gspread
from oauth2client.service_account import ServiceAccountCredentials
scope = ['https://spreadsheets.google.com/feeds']

#현재 스프레드시트의 키 파일의 경로 (구글 드라이브 api를 사용 허가 받아야 함)
json_file_name = '/Users/macname/Downloads/user_key.json'
credentials = ServiceAccountCredentials.from_json_keyfile_name(json_file_name, scope)
gc = gspread.authorize(credentials)

#스프레드시트의 url (링크가 있는 모든 이에게 공개로 전환해야하는 것 주의)
spreadsheet_url = '스프레드시트의 url'
#open doc
doc = gc.open_by_url(spreadsheet_url)
#open sheet
worksheet = doc.worksheet('시1')


#insert 회사 이름
worksheet.update_acell('B4', '회사 이름')
#insert 견적서 보낸 날
worksheet.update_acell('B5', '날짜')
#insert 수신인
worksheet.update_acell('B6', '수신인')
#insert 견적내용
worksheet.update_acell('C14', '견적 내용')
