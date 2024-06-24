import os
import pandas as pd
from io import StringIO
from selenium import webdriver
from bs4 import BeautifulSoup


def scrape_to_excel(stock_symbol, url):
    driver = webdriver.Chrome()
    driver.get(url)
    driver.implicitly_wait(20)  # gerekirse süreyi ayarlayın
    html_icerik = driver.page_source
    soup = BeautifulSoup(html_icerik, 'html.parser')
    tables = soup.find_all("table", attrs={"class": "w-full"})

    # İlk tablonun ilk satırından sütun başlıklarını al
    first_table = tables[0]
    headers = first_table.find_all('th')
    column_mapping = {i: header.get_text() for i, header in enumerate(headers)}

    # Tüm tabloları birleştirmek için bir listeye ekleyin
    all_dfs = []
    for table in tables:
        for span in table.find_all("span", class_="absolute left-full font-bold text-[10px] pl-1 "
                                                  "text-shared-danger-solid-01 opacity-0 group-hover:opacity-100"):
            span.decompose()
        for span in table.find_all("span", class_="absolute left-full font-bold text-[10px] pl-1 "
                                                  "text-shared-success-solid-01 opacity-0 group-hover:opacity-100"):
            span.decompose()
        for span in table.find_all("span", class_="absolute left-full font-bold text-[10px] pl-1 text-foreground-01 "
                                                  "opacity-0 group-hover:opacity-100"):
            span.decompose()

        html_string = str(table)
        html_io = StringIO(html_string)
        df = pd.read_html(html_io)[0]

        # Sütunları yeniden adlandır
        if not df.empty:
            df.columns = [column_mapping.get(i, f"Unknown_{i}") for i in range(df.shape[1])]
        all_dfs.append(df)

    # Tüm tabloları alt alta ekle
    combined_df = pd.concat(all_dfs, axis=0, ignore_index=True)

    return combined_df


# Özel bir konum belirleyin.
save_path = 'Borsa'

stock_symbol = "ENJSA"
urls = [
    "https://fintables.com/sirketler/ENJSA/finansal-tablolar/bilanco",
    "https://fintables.com/sirketler/ENJSA/finansal-tablolar/gelir-tablosu",
    "https://fintables.com/sirketler/ENJSA/finansal-tablolar/nakit-akim-tablosu"
]

# Tüm verileri birleştirilen DataFrame'de topla
all_data_df = pd.DataFrame()

for index, url in enumerate(urls):
    df = scrape_to_excel(stock_symbol, url)
    all_data_df = pd.concat([all_data_df, df], ignore_index=True)

# Birleştirilen verileri Excel dosyasına kaydet
excel_filename = f"{stock_symbol}_combined_data.xlsx"
full_path = os.path.join(save_path, excel_filename)
all_data_df.to_excel(full_path, index=False)
print(f"{full_path} başarıyla kaydedildi.")
