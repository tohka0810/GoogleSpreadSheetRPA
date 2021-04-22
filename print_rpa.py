import chromedriver_binary
import pyautogui as pgui
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

########
#開始番号
start_no = '1342'
#終了番号
end_no = '1353'

########

#jsonファイルを使って認証情報を取得
SCOPES = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
SERVICE_ACCOUNT_FILE = 'GCPの鍵ファイル名を記載'

credentials = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_ACCOUNT_FILE, SCOPES)

#認証情報を使ってスプレッドシートの操作権を取得
gs = gspread.authorize(credentials)
 
 #共有したスプレッドシートのキーを使ってシートの情報を取得
SPREADSHEET_KEY = 'スプレッドシートのID'
worksheet = gs.open_by_key(SPREADSHEET_KEY).worksheet('印刷したいシート名')
 
signInId = "編集権限のあるGoogleアカウント"
signInPs = "パスワード"

driver = webdriver.Chrome()
#要素がロードされるまでの待ち時間を10秒に設定
wait = WebDriverWait(driver, 10)
driver.implicitly_wait(10)
driver.get("スプレッドシートのフルパス")

driver.find_element_by_id("identifierId").send_keys(signInId)
pgui.press('Enter')

driver.find_element_by_name("password").send_keys(signInPs)
pgui.press('Enter')

print('login success.')

time.sleep(10)

driver.get("スプレッドシートのフルパス")

print('pagechange success.')

time.sleep(5)


while start_no != str(int(end_no)+1):
	print('start_no = ' +start_no)
	print('end_no = ' +end_no)
	
	worksheet.update_acell('B2',start_no)
	
	pgui.keyDown('command')
	pgui.press('p')
	pgui.keyUp('command')
	
	time.sleep(7)
	
	pgui.press('\t')
	pgui.press('Enter')
	
	time.sleep(7)
	
	pgui.press('Enter')
	
	start_no = str(int(start_no) +1)
	time.sleep(3)

print('end print_rpa.py')
time.sleep(15)
driver.quit()