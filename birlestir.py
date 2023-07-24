"""
Aynı klasörde bulunan excel dosyalarını birleştiren python kodu.
Tüm dosya içeriği birleştirildikten sonsa "TUMU.xlsx" adında yeni bir dosyaya kaydedilecek.
Python Kodunun çalışması için bilgisayarınızda "Pandas", "openpyxl" ve "xlrd." kütüphanelerinin / modüllerinin yüklü olması gerekir.
"""
import pandas as pd
import os

os.remove("TUMU.xlsx")      # Klasör içindeki "TUMU.xlsx" dosyasını siler. Eğer bu dosya mevcut değilse bu satırı çalıştırmayın ya da silin. 

dosyalar = os.listdir()     # "birlestir.py" dosyasının bulunduğu dizindeki tüm dosya isimlerini, uzantıları ile birlikte al, "dosyalar" isimli değişkene ata.
dosyalar.sort()             # dosya isimlerini sırala

dosya_isimleri= []      # ".xlsx" ya da ".xls" uzantılı dosyaların toplanacağı boş liste oluştur.

for i in dosyalar:      # Dizindeki tüm dosya isimlerini kontrol et, ".xlsx", ".xls" ya da ".ods" uzantılı dosyaları "dosya_isimleri" isimli listeye ekle.
    if ((i[-5:] == ".xlsx") or (i[-4:] == ".xls") or (i[-4:] == ".ods") ):     # dosya uzantılarını kontrol et.
        dosya_isimleri.append(i)

def VeriCercevesi(dosya_adi):   # Belirtilen dosya adına göre, dosya içeriğini DataFrame'e çeviren fonksiyon.
    return pd.read_excel(dosya_adi)

baslik = VeriCercevesi(dosya_isimleri[0]).columns       # ilk dosyanın basligi

def VeriCercevesiBasliksiz(dosya_adi):   # Belirtilen dosya adına göre, dosya içeriğini Başlıksız DataFrame'e çeviren fonksiyon.
    return pd.read_excel(dosya_adi, header = None, names = baslik, skiprows=[0])  # dosya adındaki başlık silinecek, klasördeki ilk dosyanın başlığı tüm diğer dosyalara başlık olarak eklenecek.

data_frame = pd.DataFrame()     # Verileri toplayacağımız boş bir veri çerçevesi oluşturuyoruz.

for dosya in dosya_isimleri:

    if dosya == dosya_isimleri[0]:
        df = VeriCercevesi(dosya)
        df["Dosya_Adi"] = dosya
        data_frame = df

    else:
        df_dosya = VeriCercevesiBasliksiz(dosya)
        df_dosya["Dosya_Adi"] = dosya
        data_frame = pd.concat([data_frame, df_dosya])

# data_frame.dropna(axis=1, inplace=True)     # Kayıp veri barındıran Sütunlar siliniyor ve dosya güncelleniyor / üzerine yazılıyor. Bu kod, sütun içerisinde 1 adet kayıp veri olsa bile o sütunu komple siliyor.

data_frame.to_excel("TUMU.xlsx")    # Tüm dosyalar birleştirildikten sonra sonuç "TUMU.xlsx" ismi ile kaydedilir.