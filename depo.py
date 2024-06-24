"""import requests
from bs4 import BeautifulSoup
import pandas as pd
from io import StringIO
import os
from selenium import webdriver

def scrape_to_excel(stock_symbol):
    url = f'https://fintables.com/sirketler/ASTOR/finansal-tablolar/bilanco'
    headers = {"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36"}
    driver = webdriver.Chrome()  # veya kullandığınız tarayıcıya uygun sürücü
    driver.get(url)

    # JavaScript'in yüklenmesini bekle
    driver.implicitly_wait(10)  # gerekirse süreyi ayarlayın

    html_icerik = driver.page_source
    soup = BeautifulSoup(html_icerik, 'html.parser')
    table = soup.find_all("table",attrs={"class":"w-full"})
    print(table)

    # HTML metnini StringIO nesnesine sarın ve read_html'e geçirin.
    html_string = str(table)
    html_io = StringIO(html_string)
    dfs = pd.read_html(html_io)

    # Özel bir konum belirleyin.
    save_path = 'C:\\Users\\KBmonscer\\Downloads\\Borsa'
    excel_filename = f"{stock_symbol}_data.xlsx"
    full_path = os.path.join(save_path, excel_filename)

    # Excel dosyasını belirtilen konuma kaydedin.
    dfs.to_excel(full_path, index=False)
    print(f"{full_path} başarıyla kaydedildi.")



# Kullanıcıdan hisse senedi sembolünü alın.
stock_symbol = "ASTOR"#input("Lütfen hisse senedi sembolünü giriniz: ")

# Fonksiyonu çağırın.
scrape_to_excel(stock_symbol)"""
"""
#stock_symbol = input("Lütfen hisse senedi sembolünü giriniz: ")
url = f'https://tr.investing.com/equities/turk-hava-yollari-balance-sheet'
headers = {"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36"}
response = requests.get(url)
soup = BeautifulSoup(response.text, 'html.parser')
filtre = soup.find( id ="rrtable")
print(filtre)
from selenium import webdriver
from bs4 import BeautifulSoup
from io import StringIO
import pandas as pd
import os

def scrape_to_excel(stock_symbol):
    url = f'https://fintables.com/sirketler/{stock_symbol}/finansal-tablolar/bilanco'
    driver = webdriver.Chrome()  # veya kullandığınız tarayıcıya uygun sürücü
    driver.get(url)
    # JavaScript'in yüklenmesini bekle
    driver.implicitly_wait(10)  # gerekirse süreyi ayarlayın
    html_icerik = driver.page_source
    soup = BeautifulSoup(html_icerik, 'html.parser')
    table = soup.find_all("table", attrs={"class": "w-full"})

    print(table)
    # HTML metnini StringIO nesnesine sarın ve read_html'e geçirin.
    html_string = str(table)
    html_io = StringIO(html_string)
    dfs = pd.read_html(html_io)

    # Özel bir konum belirleyin.
    save_path = 'C:\\Users\\KBmonscer\\Downloads\\Borsa'

    # Tüm tabloları kaydedin.
    for i, df in enumerate(dfs):
        excel_filename = f"{stock_symbol}_table_{i}_data.xlsx"
        full_path = os.path.join(save_path, excel_filename)
        df.to_excel(full_path, index=False)
        print(f"{full_path} başarıyla kaydedildi.")


# Kullanıcıdan hisse senedi sembolünü alın.
stock_symbol = "ASTOR"  # input("Lütfen hisse senedi sembolünü giriniz: ")

# Fonksiyonu çağırın.
scrape_to_excel(stock_symbol)
"""