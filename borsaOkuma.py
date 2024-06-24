import pandas as pd

# Kaydedilen Excel dosyasının yolunu belirleyin
excel_path = 'Borsa\\ASTOR_combined_data.xlsx'

# Yeni Excel dosyasının adını belirleyin
new_excel_path = 'Borsa\\ASTOR_selected_rows.xlsx'

# Excel dosyasını oku
df = pd.read_excel(excel_path)
"""
# İstenen hücre aralıklarını bir listeye koy
cell_ranges = ['A12:Z12', 'A26:Z26', 'A42:Z42', 'A55:Z55', 'A58:Z58',
               'A72:Z72', 'A80:Z80', 'A83:Z83', 'A84:Z84', 'A101:Z101',
               'A110:Z110', 'A111:Z111', 'A111:Z111', 'A155:Z155', 'A158:Z158',
               'A135:Z135', 'A142:Z142', 'A137:Z137', 'A138:Z138', 'A154:Z154']

# Belirtilen hücre aralıklarındaki satırları içeren yeni bir DataFrame oluştur
selected_rows = []

for cell_range in cell_ranges:
    start_cell, end_cell = cell_range.split(':')
    start_row = int(start_cell[1:]) - 1  # Excel'de satır numaraları 1'den başlar, pandas'ta ise 0'dan

    # Belirtilen satırı listeye ekle
    selected_rows.append(df.iloc[start_row:start_row+1])

# Listeyi yeni bir DataFrame'e dönüştür ve kaydet
selected_rows_df = pd.concat(selected_rows)
selected_rows_df.to_excel(new_excel_path, index=False)
print(f"{new_excel_path} başarıyla kaydedildi.")"""
def convert_to_float(value):
    if isinstance(value, str):
        if 'Bin TRY' in value:
            return 0.0  # 'Bin TRY' olanları 0.0 olarak ayarla
        return float(value.replace('.', '').replace(',', '.'))
    return value

# Tüm veri çerçevesi için dönüştürme işlemi
for col in df.columns[1:]:  # 'Bilanço Kalemleri' kolonunu atlıyoruz
    df[col] = df[col].apply(convert_to_float)

# Bilanço kalemlerini uygun şekilde seçin
def get_value(df, kalem):
    value = df.loc[df['Bilanço Kalemleri'] == kalem, '2024/3']
    return value.values[0] if not value.empty else 0.0

toplam_varliklar = get_value(df, 'Toplam Varlıklar')
kaynaklar = get_value(df, 'Toplam Kaynaklar')
uzun_vadeli_borclar = get_value(df, 'Toplam Uzun Vadeli Yükümlülükler')
finansal_borclar = get_value(df, 'Finansal Borçlar')+get_value(df, 'Finansal Borçlar2')
donen_varliklar = get_value(df, 'Toplam Dönen Varlıklar')
kisa_vadeli_borclar = get_value(df, 'Toplam Kısa Vadeli Yükümlülükler')
stoklar = get_value(df, 'Stoklar')
nakit_ve_nakit_benzerleri = get_value(df, 'Nakit ve Nakit Benzerleri')
toplam_borclar=kisa_vadeli_borclar+uzun_vadeli_borclar
# Hesaplamalarda NaN kontrolü ve hata ayıklama
def safe_divide(a, b):
    return a / b if b != 0 else float('nan')

# Kaldıraç Oranı (Debt to Equity Ratio)
kaldirac_orani = 100*safe_divide(toplam_borclar, kaynaklar)

# Finansal Borç Oranı (Faiz Yükü) (Financial Debt Ratio)
finansal_borc_orani = 100*safe_divide(finansal_borclar, toplam_varliklar)

# Cari Oran (Current Ratio)
cari_oran = safe_divide(donen_varliklar, kisa_vadeli_borclar)

# Likidite Oranı (Quick Ratio)
likidite_orani = safe_divide(donen_varliklar - stoklar, kisa_vadeli_borclar)

# Nakit Oranı (Cash Ratio)
nakit_orani = safe_divide(nakit_ve_nakit_benzerleri, kisa_vadeli_borclar)

# Oranları yazdır
print(f"Kaldıraç Oranı: {kaldirac_orani}")
print(f"Finansal Borç Oranı (Faiz Yükü): {finansal_borc_orani}")
print(f"Cari Oran: {cari_oran}")
print(f"Likidite Oranı: {likidite_orani}")
print(f"Nakit Oranı: {nakit_orani}")

new_data = {
    'Bilanço Kalemleri': [
        'Kaldıraç Oranı',
        'Finansal Borç Oranı (Faiz Yükü)',
        'Cari Oran',
        'Likidite Oranı',
        'Nakit Oranı'
    ],
    '2024/3': [
        kaldirac_orani,
        finansal_borc_orani,
        cari_oran,
        likidite_orani,
        nakit_orani
    ]
}
#------------------------------------------
# Gerekli verileri çekme
net_kar = get_value(df, 'Brüt Kar (Zarar)')
faaliyet_kar = get_value(df, 'Faaliyet Karı (Zararı)')
toplam_gelir = get_value(df, 'Satış Gelirleri')
faiz_giderleri = get_value(df, 'Ödenen Faiz')
dolasimdaki_hisse_sayisi = get_value(df, 'Ödenmiş Sermaye')
toplam_varliklar = get_value(df, 'Toplam Kaynaklar')

# Hesaplamalarda NaN kontrolü ve hata ayıklama
def safe_divide(a, b):
    return a / b if b != 0 else float('nan')

# Hisse Başına Kar (Earnings per Share, EPS)
hisse_basina_kar = safe_divide(net_kar, dolasimdaki_hisse_sayisi)

# Faize Karşı Kazanç (Interest Coverage Ratio)
faize_karsi_kazanc = safe_divide(faaliyet_kar, faiz_giderleri) * 100

# Faaliyet Kar Marjı (Operating Profit Margin)
faaliyet_kar_marji = safe_divide(faaliyet_kar, toplam_gelir) * 100

# Satış Karlılığı (Net Profit Margin)
satis_karliligi = safe_divide(net_kar, toplam_gelir) * 100

# Şirketin Kar Oluşturma Verimliliği (Return on Assets, ROA)
roa = safe_divide(net_kar, toplam_varliklar) * 100

# Oranları yazdır
print(f"Hisse Başına Kar: {hisse_basina_kar}")
print(f"Faize Karşı Kazanç: {faize_karsi_kazanc}")
print(f"Faaliyet Kar Marjı: {faaliyet_kar_marji}")
print(f"Satış Karlılığı: {satis_karliligi}")
print(f"Şirketin Kar Oluşturma Verimliliği: {roa}")

# Yeni verileri ekleme
new_data = {
    'Bilanço Kalemleri': [
        'Hisse Başına Kar',
        'Faize Karşı Kazanç (%)',
        'Faaliyet Kar Marjı (%)',
        'Satış Karlılığı (%)',
        'Şirketin Kar Oluşturma Verimliliği (%)'
    ],
    '2024/3': [
        hisse_basina_kar,
        faize_karsi_kazanc,
        faaliyet_kar_marji,
        satis_karliligi,
        roa
    ]
}
# Yeni satırları DataFrame'e ekleme
new_df = pd.DataFrame(new_data)
df = pd.concat([df, new_df], ignore_index=True)

# Güncellenmiş veri çerçevesini Excel dosyasına yaz
df.to_excel(new_excel_path, index=False)