import os
import pandas as pd
import json
# ブラウザを自動操作するためseleniumをimport
from selenium import webdriver
# seleniumでヘッドレスモードを指定するためにimport
from selenium.webdriver.chrome.options import Options
# seleniumでEnterキーを送信する際に使用するのでimport
from selenium.webdriver.common.keys import Keys
import time
import chromedriver_binary #add
from webdriver_manager.chrome import ChromeDriverManager
from janome.tokenizer import Tokenizer

# urlリスト
def df_bloglist_func():
    PWD = os.getcwd() # 現在のパス
    df_bloglist = pd.read_excel(f'{PWD}/blog_list.xlsx', engine='openpyxl')
    df_bloglist.dropna(how='any', axis=0, inplace=True)

    return df_bloglist

def google_driver(): # googleドライブの設定（return driver)
    # seleniumで自動操作するブラウザはGoogleChrome(Optionsオブジェクトを作成)
    options = Options()

    # ヘッドレスモードを有効にする
    options.add_argument('--headless')
    # options.headless = True

    # Seleniumでダウンロードするデフォルトディレクトリ
    PWD = os.getcwd() # 現在のパス
    download_directory = PWD

    # クロームオプションの設定
    options = webdriver.ChromeOptions()

    # デフォルトのダウンロードディレクトリの指定
    prefs = {"download.default_directory" : download_directory}  

    # オプションを指定してクロームドライバーの起動
    options.add_experimental_option("prefs", prefs)

    # ChromeのWebDriverオブジェクトを作成
    driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
    
    return driver

def get_Url_Title(google_driver, url, acount, password):
    ''''
    wordpressへのログイン
    ブログタイトルとURKの取得
    '''
    # ドライバーの設定
    driver = google_driver()

    # wordpressログイン
    driver.get(f'{url}/wp-admin/')
    acount = acount
    password = password

    #管理画面への遷移
    driver.get(f'{url}/wp-admin')
    time.sleep(3)
    # wordpressのログイン
    driver.find_element_by_xpath("//*[@id='user_login']").send_keys(acount)
    driver.find_element_by_xpath("//*[@id='user_pass']").send_keys(password)
    login = driver.find_element_by_xpath("//*[@id='wp-submit']").click()
    time.sleep(3)

    # 投稿一覧へ遷移→公開済へ遷移
    driver.find_element_by_xpath("//*[@id='menu-posts']/a").click()
    driver.find_element_by_xpath("//ul[@class='subsubsub']/li[@class='publish']").click()
    
    try:
        # ページ数の取得（整数型へ）
        pages = int(driver.find_element_by_xpath("//*[@class='tablenav-pages']//span[@class='total-pages']").text)
    except:
        pages = 1
    print(url,pages)
    print(30*'=')
    elements_list = []

    # 各ページ
    for page in range(pages):
        print(page+1)
        elements = driver.find_elements_by_xpath("//div[@class='row-actions']/span[@class='view']/a")

        for element in elements:
            element_dic = {}
            element_dic['url'] = element.get_attribute('href')
            element_dic['title'] = element.get_attribute('aria-label').replace('を表示','').replace('“','').replace('”','')
            elements_list.append(element_dic)
        try:
            driver.find_element_by_xpath("//form[@id='posts-filter']/div[1]/div[3]//a[@class='next-page button']").click()
        except:
            pass
        time.sleep(1)
    driver.close() # ブラウザを閉じる
    df = pd.DataFrame(elements_list)  
    return df


# 品詞分解
def title_Tokenizer(df):
    t = Tokenizer() 
    text_list = []
    for text in df['title']:
        token_list = []
        for token in t.tokenize(text):
            if (token.part_of_speech.split(',')[0] == '名詞') & (token.part_of_speech.split(',')[1] in ['一般', '固有名詞']):
                token_list.append(token.surface)
        text_list.append(token_list)
    df_ = pd.DataFrame(text_list)
    df =pd.concat([df,df_],axis=1)

    return df

# Excelへの出力
def output_excel(filenme, df):
    # outputフォルダ作成
    PWD = os.getcwd()
    file_make = f'{PWD}/output'
    if not os.path.exists(file_make):
        os.mkdir(file_make)

    df.to_excel(f'{file_make}/{filenme}.xlsx',index=False, encoding='cp932')
    return df

if __name__ =='__main__':
    # urlのリストを取得
    df_bloglist = df_bloglist_func()
    print(df_bloglist)
    for i in range(len(df_bloglist)):
        url = df_bloglist['URL'][i] 
        filenme = df_bloglist['ブログタイトル'][i]
        acount = df_bloglist['acount'][i]
        password = df_bloglist['password'][i]
    
        # 初期設定（ページ数）
        df = pd.DataFrame()
        df = get_Url_Title(google_driver, url, acount, password)

        # 品詞分解(名詞のみ取得)
        df = title_Tokenizer(df)

        # Excelファイル出力
        output_excel(filenme, df)