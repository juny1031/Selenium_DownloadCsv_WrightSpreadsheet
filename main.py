from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time
import datetime
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials 

# Seleniumをあらゆる環境で起動させるChromeオプション
options = Options()
options.add_argument('--disable-gpu')
options.add_argument('--disable-extensions')
options.add_argument('--proxy-server="direct://"')
options.add_argument('--proxy-bypass-list=*')
options.add_argument('--start-maximized')
options.add_argument('--headless')

# ダウンロードフォルダを指定
download_dir = "フォルダのパス"
prefs = {"download.default_directory" : download_dir}
options.add_experimental_option("prefs",prefs)

DRIVER_PATH = 'ドライバーのパス'
driver = webdriver.Chrome(executable_path=DRIVER_PATH, chrome_options=options)

# headlessモードでもダウンロードをできるようにする
driver.command_executor._commands["send_command"] = (
    "POST",
    '/session/$sessionId/chromium/send_command'
)
params = {
    'cmd': 'Page.setDownloadBehavior',
    'params': {
        'behavior': 'allow',
        'downloadPath': download_dir
    }
}
driver.execute("send_command", params=params)

# ログインページにアクセス
driver.get("アクセスするURL")

time.sleep(10)

# メールアドレスを記入
id = driver.find_element_by_name("email")
id.send_keys("メールアドレス")
time.sleep(5)
next_button = driver.find_element_by_xpath('XPath')
next_button.click()

time.sleep(5)

# パスワードを記入
password = driver.find_element_by_name("password")
password.send_keys("パスワード")
time.sleep(5)
next_button =driver.find_element_by_xpath('XPath')
next_button.click()
time.sleep(5)

# ダウンロード先のURLへアクセス
driver.get("アクセスするURL")

time.sleep(5)

# ダウンロード対象の条件を指定
calender_button = driver.find_element_by_xpath('XPath')
calender_button.click()
time.sleep(1)

from_date = driver.find_element_by_name('name')
from_date.clear()
from_date.send_keys("集計対象開始日")
time.sleep(1)

today = datetime.datetime.now()
yesterday = today - datetime.timedelta(days=1)
target_day = yesterday.strftime('%Y%m%d')
to_date = driver.find_element_by_name('name')
to_date.clear()
to_date.send_keys(str(target_day))
time.sleep(1)
apply_button = driver.find_element_by_xpath('XPath')
apply_button.click()
time.sleep(1)

group_button = driver.find_element_by_xpath('XPath')
group_button.click()
time.sleep(1)

select_button = driver.find_element_by_xpath('XPath')
select_button.click()
time.sleep(1)

job_button = driver.find_element_by_xpath('XPath')
job_button.click()
time.sleep(1)

select_22biz = driver.find_element_by_xpath('XPath')
select_22biz.click()
time.sleep(1)

download_button = driver.find_element_by_xpath('XPath')
download_button.click()
time.sleep(5)
driver.quit()

download_file_name = download_dir + "パス" + target_day +".csv"

# ダウンロードしたcsvから使う列を指定してdfにする
read_df = pd.read_csv(download_file_name ,encoding='utf-16',sep='\t', skipinitialspace=True,header=0, usecols=[4,5,9,10,16,17,20,64,65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,81,82])
# NaNをnullに変換
output_df = read_df.fillna("")

# GoogleAPIと連携
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name("秘密鍵.jsonのパス", scope)
gc = gspread.authorize(credentials)
ss = gc.open_by_key('スプレッドシートID')
sh = ss.worksheet('シート名')

# update_cellsをする際の範囲指定の英語列の判定
def toAlpha(num):
    if num<=26:
        return chr(64+num)
    elif num%26==0:
        return toAlpha(num//26-1)+chr(90)
    else:
        return toAlpha(num//26)+chr(64+num%26)

col_lastnum = len(output_df.columns) # DataFrameの列数
row_lastnum = len(output_df.index)   # DataFrameの行数

# スプレッドシートに書き込み
cell_list = sh.range('A1:'+toAlpha(col_lastnum)+str(row_lastnum+1))
for cell in cell_list:
    if cell.row == 1:
        val = output_df.columns[cell.col-1]
    else:
        val = output_df.iloc[cell.row-2][cell.col-1]
    cell.value = val
sh.update_cells(cell_list)
