import csv
import os
import pandas as pd
from zipfile import ZipFile
import locale

#MKK dosyalarını zip'ten çıkarmak
zip_folder = "Q:/_HiSenetl/_PARYA/MKK/MKK_INDIRILEN_DOSYALAR/165/MKK KOM"
extract_folder = "./dosyalar"
for file in os.listdir(zip_folder):
    with ZipFile(os.path.join(zip_folder, file), 'r') as zip_ref:
        zip_ref.extractall(extract_folder)
csv_folder = "./dosyalar/"

#MKK Dosyaları birleştirmek
combined_df = pd.concat([pd.read_csv(os.path.join(csv_folder, file), sep=';', thousands=".", decimal=",", low_memory=False) for file in os.listdir(csv_folder)], ignore_index=True)
combined_df.columns=["DONEM","UYE_KODU","HESAP_NO","KOMISYON_TURU","ACIKLAMA","MATRAH","KOMISYON", "BSMV/KDV"]

# 'KOMISYON' and 'BSMV/KDV' sutunlarını stringe donusturme, virgul-nokto ve numeric duzenlemesı
combined_df['KOMISYON'] = combined_df['KOMISYON'].astype(str).str.replace(',', '.').astype(float)
combined_df['BSMV/KDV'] = combined_df['BSMV/KDV'].astype(str).str.replace(',', '.').astype(float)

# TIB HESAPLARI ICIN KOMISYON HESAPLAMA
tib_komisyon = combined_df[(combined_df['HESAP_NO'].str.len() >= 10) & (combined_df['HESAP_NO'].str.contains('B|AKHATAP'))]

# 'KOMISYON' ve 'BSMV/KDV' degerlerini toplama
tib_komisyon['KOMISYON_BSMV_KDV'] = tib_komisyon['KOMISYON'] + tib_komisyon['BSMV/KDV']

# 'KOMISYON_TURU'ne gore filtreleme ve 'KOMISYON_BSMV_KDV' toplama
tib_komisyon_sum = tib_komisyon.groupby('KOMISYON_TURU')['KOMISYON_BSMV_KDV'].sum()

# Format duzenleme
tib_komisyon_sum_formatted = tib_komisyon_sum.apply(lambda x: locale.format_string("%.2f", x, grouping=True))




# HB HESAPLARI ICIN KOMISYON HESAPLAMA
hb_komisyon = combined_df[combined_df['HESAP_NO'].str.len() == 8]

# 'KOMISYON' ve 'BSMV/KDV' degerlerini toplama
hb_komisyon['KOMISYON_BSMV_KDV'] = hb_komisyon['KOMISYON'] + hb_komisyon['BSMV/KDV']

# 'KOMISYON_TURU'ne gore filtreleme ve 'KOMISYON_BSMV_KDV' toplama
hb_komisyon_sum = hb_komisyon.groupby('KOMISYON_TURU')['KOMISYON_BSMV_KDV'].sum()

# Format duzenleme
hb_komisyon_sum_formatted = hb_komisyon_sum.apply(lambda x: locale.format_string("%.2f", x, grouping=True))



# IYM HESAPLARI ICIN KOMISYON HESAPLAMA
filtered = combined_df[combined_df['HESAP_NO'].str.len() < 8]
iym_komisyon=(~filtered['HESAP_NO'].str.contains('B|AKHATAP'))

#'KOMISYON' ve 'BSMV/KDV' degerlerini toplama
iym_komisyon['KOMISYON_BSMV_KDV'] = iym_komisyon['KOMISYON'] + iym_komisyon['BSMV/KDV']


# 'KOMISYON_TURU'ne gore filtreleme ve 'KOMISYON_BSMV_KDV' toplama
iym_komisyon_sum = iym_komisyon.groupby('KOMISYON_TURU')['KOMISYON_BSMV_KDV'].sum()

# Format duzenleme
iym_komisyon_sum_formatted = iym_komisyon_sum.apply(lambda x: locale.format_string("%.2f", x, grouping=True))

#Nihai dosyayı olusturma
# Degerleri tek bir dataframede bir araya getirmek
combined_results = pd.concat([tib_komisyon_sum, hb_komisyon_sum, iym_komisyon_sum], axis=1)
combined_results.columns = ['TIB_KOMISYON_SUM', 'HB_KOMISYON_SUM', 'IYM_KOMISYON_SUM'] #sutun isimlerini getirmek
combined_results.fillna(0, inplace=True) # Bos satırlara 0 yazma

# Excel dosyası olusturma
excel_file_path = "Q:/_HiSenetl/_PARYA/MKK/MKK_INDIRILEN_DOSYALAR/165/MKK_Komisyon.xlsx"
combined_results.to_excel(excel_file_path)


