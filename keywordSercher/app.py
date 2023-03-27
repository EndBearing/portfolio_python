import json
import os
import re
import shutil
import sys
import zipfile
from pip._vendor import requests
import certifi
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome import service as cs
from selenium.common.exceptions import SessionNotCreatedException
from bs4 import BeautifulSoup
from time import sleep
import datetime



# インストール済みライブラリ
# selenium
# requests
# BeautifulSoup4
# lxml-xml


# print(certifi.where())

# =====================================
# ChromeWebdriverの自動インストール・更新 
# =====================================
webdriver_url   = "https://chromedriver.storage.googleapis.com/" #ウェブドライバーページ
file_name       = "chromedriver_win32.zip" #Windows用のファイル名
chrome_service = cs.Service(executable_path=os.getcwd() + "\\webdriver\\chromedriver.exe")
# エラー出力ファイルオープン
error_file = open("log\error.text", encoding="utf-8", mode="w")

# ChromeWebdriverファイルのパス指定
try:
    if os.path.isfile(os.getcwd() + "\\webdriver\\chromedriver.exe")== False:
        raise FileNotFoundError()

    options = webdriver.ChromeOptions() 
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    options.add_argument('--headless')  #ヘッドレス化
    driver = webdriver.Chrome(options=options,service=chrome_service)
    driver.close()
    driver.quit()
except (FileNotFoundError,SessionNotCreatedException) as e:
    if type(e) == SessionNotCreatedException:
        error_file.write('[WebDriverファイルが古い可能性があります。最新バージョンのダウンロードします。:::'+ datetime.datetime.now()+']\n')
    elif type(e)==FileNotFoundError:
        error_file.write('[WebDriverファイルが存在しません。ダウンロードを開始開始します。:::'+ datetime.datetime.now()+']\n')
    else:
        error_file.write('[不明な例外です  :::'+ e +':::'+ datetime.datetime.now()+']\n')
        exit()

    response = requests.get(webdriver_url)
    soup = BeautifulSoup(response.text,"lxml-xml")

    if not os.path.exists(os.getcwd() + "\\webdriver\\tmp\\"):
        os.makedirs(os.getcwd() + "\\webdriver\\tmp\\")

    success_flg = False
    version_arr ={}
    cnt = 0
    for version in reversed(soup.find_all("Key")):
        cnt += 1
        if file_name in version.text:
            version_name = re.compile(r'/.*').sub("",version.text)
            version_name = re.compile(r'\..*').sub("", version_name).zfill(5)+str(cnt).zfill(5)

            version_arr[version_name] = version.text

    for version in sorted(version_arr.items(),reverse=True):
        zip_source = requests.get(webdriver_url+version[1])
        # print("起動テストを開始します\t"+webdriver_url+version[1])

        # ダウンロードしたZIPファイルの書き出し
        with open(os.getcwd() + "\\webdriver\\tmp\\" + file_name, "wb") as file:
            for chunk in zip_source.iter_content():
                file.write(chunk)

        # ZIPファイルの解凍
        with zipfile.ZipFile(os.getcwd() + "\\webdriver\\tmp\\" + file_name) as file:
            file.extractall(os.getcwd() + "\\webdriver\\tmp\\")

        try:
            driver = webdriver.Chrome(executable_path=os.getcwd() + "\\webdriver\\tmp\\chromedriver.exe")
            if(os.path.isfile(os.getcwd() + "\\webdriver\\chromedriver.exe")):
                os.remove(os.getcwd() + "\\webdriver\\chromedriver.exe")
            shutil.move(os.getcwd() + "\\webdriver\\tmp\\chromedriver.exe",os.getcwd() + "\\webdriver\\chromedriver.exe")

            # print("正常に起動しました。WebDriverを上書きします。")

            shutil.rmtree(os.getcwd() + "\\webdriver\\tmp\\")
            success_flg = True
            break
        except SessionNotCreatedException as e:
            error_file.write("[起動中にエラーが発生しました。:::" + webdriver_url+version[1] + ':::' + datetime.datetime.now()+']\n')
    #print(version_arr)
    if not success_flg:
        error_file.write("[WebDriverファイルの上書き中に例外が発生しました。:::" + datetime.datetime.now()+']\n')

error_file.close()

# =====================================
# 以下、本処理
# =====================================

file = open("config.json", encoding="utf-8", mode="r")
json_content = json.load(file)
file.close

options = webdriver.ChromeOptions() 
options.add_experimental_option('excludeSwitches', ['enable-logging'])
#これがエラーをなくすコードです。ブラウザ制御コメントを非表示化しています
options.add_argument('--headless')  #ヘッドレス化
options.add_argument('--window-size=1920,1080')
options.add_argument('--user-agent="Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36"')
# ヘッドレス化するとuser-agentが空になるのでoutlookdでブラウザにはじかれていた。user-agentを指定して（偽装して）黙らせる。

driver = webdriver.Chrome(options=options,service=chrome_service)
driver.get('http://webmail.otani.ac.jp/')
sleep(5)

# 認証ページ処理
login_id = driver.find_element(By.XPATH, "//input[@id='userNameInput']") 
login_pw = driver.find_element(By.XPATH, "//input[@id='passwordInput']")
login_id.send_keys(json_content['config']['ID']) 
login_pw.send_keys(json_content['config']['PW'])
login_btn = driver.find_element(By.XPATH, ".//span[@id='submitButton']") 
login_btn.click()

# 10秒待機 （ページ読み込み）
sleep(10)

# Aythenticator　ページ
# アプリに認証通知が飛ぶので承認されるまで1:30は待機。
#カレントページのURLを取得
cur_url = driver.current_url
cnt = 0
while cur_url != driver.current_url :
    sleep(15)
    # print(cnt)
    cnt += 1
    if cnt > 6:
        break

# 承認されなかったら、プログラム終了
if cur_url != driver.current_url:
    driver.close()
    driver.quit()
    sys.exit()


# メール読み切り待ち
sleep(30)

from_list = [] #送信者
title_list = [] #件名
day_list = [] #日付

#メール件名リスト作成 
title_list = driver.find_elements(By.XPATH,'//div[contains(@class, "S2NDX")]/div[2]/div/span')
# メール送信元リスト
from_list = driver.find_elements(By.XPATH,'//div[contains(@class, "S2NDX")]/div[1]/div/span')
# メール受信日取得 
day_list = driver.find_elements(By.XPATH,'//div[contains(@class, "S2NDX")]/div[2]/span')

# ループして文字起こし+選定
detect_list = [[],[],[]]
for i in range(len(title_list)):
     text = title_list[i].get_attribute('innerHTML')
     where_from = from_list[i].get_attribute('innerHTML')
     date = day_list[i].get_attribute('innerHTML')

     isAppend = False
     for i in json_content['keyword']:

        # キーワードが指定されていなかったらスキップ
        if json_content['keyword'][i] == '':
            continue

        # 選定処理, 重複判定
        if json_content['keyword'][i] in text and isAppend == False:
            detect_list[0].append(text)
            detect_list[1].append(where_from)
            detect_list[2].append(date)
            isAppend = True

# 辞書型に格納
data = {
    'title' : detect_list[0],
    'from' : detect_list[1],
    'date' : detect_list[2]
}

# ドライバー閉じる
driver.close()

# =====================================
# 以下、LINE通知処理
# =====================================

#LINE通知用の関数
def line_notify(message):
    try:
        line_notify_token = json_content['config']['token']
        line_notify_api = 'https://notify-api.line.me/api/notify'
        payload = {'message':  message,
                    'stickerPackageId': '8522',
                    'stickerId': '16581283'
                    }
        headers = {'Authorization': 'Bearer ' + line_notify_token}
        requests.post(line_notify_api, data=payload, headers=headers,verify=certifi.where())
    except requests.exceptions.RequestException as e:
        # エラー出力ファイルオープン
        error_file = open("log\error.text", encoding="utf-8", mode="w")
        error_file.write('[line_notifyでエラー  ::'+ type(e) +':'+ e +':::'+ datetime.datetime.now()+']\n')
        error_file.close()


# 通知メッセージ作成(最大1000字) 
if data['title'] == [] and data['from'] == [] and data['date'] == []:
    m = 'キーワードと一致するメールはありませんでした。' 
else:
    content =""
    post_num = len(data['title'])
    if post_num != 0 :
        for num in range(post_num):
            content = content + "=======================\n \
受信日時： " + data['date'][num] +  "\n \
送信者: "+ data['from'][num] + "\n \
<件名> \n" + data['title'][num] + "\n \
\n \
\
"
    m ="キーワードに関連したメールが"+ str(post_num)+ "件見つかりました！！\n" + content

#実行プログラム本体
line_notify(m)


driver.quit()
sys.exit()