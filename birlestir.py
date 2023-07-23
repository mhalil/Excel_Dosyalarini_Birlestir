"""
Aynı klasörde bulunan excel dosyalarını birleştiren python kodu.
Tüm dosya içeriği birleştirildikten sonsa "TUMU.xlsx" adında yeni bir dosyaya kaydedilecek.
"""
import pandas as pd
import os

# miktar = int(input("Kaç adet dosya içeriğini birleştirmek istiyorsunuz ?"))
# ilk_dosya_numarasi = int(input("İlk dosya numarasını yazın. Dosya numaraları sıralı olmalıdır: "))

dosyalar = os.listdir()     # "birlestir.py" dosyasının bulunduğu dizindeki tüm dosya isimlerini, uzantıları ile birlikte al, "dosyalar" isimli değişkene ata.
 
dosya_isimleri= []      # ".xlsx" ya da ".xls" uzantılı dosyaların toplanacağı boş liste oluştur.

for i in dosyalar:      # Dizindeki tüm dosya isimlerini kontrol et, ".xlsx" ya da ".xls" uzantılı dosyaları "dosya_isimleri" isimli listeye ekle.
    if ((i[-5:] == ".xlsx") or (i[-4:] == ".xls")):     # dosya uzantılarını kontrol et.
        dosya_isimleri.append(i)

def VeriCercevesi(dosya_adi):   # Belirtilen dosya adına göre, dosya içeriğini DataFrame'e çeviren fonksiyon.
    return pd.read_excel(dosya_adi)

"""
Veri çerçevesi oluşturuken, dosya ismini de bir sütuna yazdır.
"""

df = pd.DataFrame()

for dosya in dosya_isimleri:
    df_dosya = VeriCercevesi(dosya)
    print(df_dosya.head())
    # df.join(df_dosya)


# print(df.describe())
