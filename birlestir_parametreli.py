"""
Aynı klasörde bulunan excel dosyalarını birleştiren python kodu.
Tüm dosya içeriği birleştirildikten sonsa "TUMU.xlsx" adında yeni bir dosyaya kaydedilecek.
Python Kodunun çalışması için bilgisayarınızda "Pandas", "openpyxl" ve "xlrd." kütüphanelerinin / modüllerinin yüklü olması gerekir.
"""
import pandas as pd
from glob import glob
import os

dosyalar = glob("*.xls*")     		# Python dosyasının bulunduğu dizindeki (klasördeki) tüm excel dosya isimlerini, uzantıları ile birlikte al, "dosyalar" isimli listede topla.
dosyalar.sort()             		# dosyalar listesindeki öğeleri (dosya isimlerini) alfabetik olarak sırala.

if "TUMU.xlsx" in dosyalar:         # Klasör içinde "TUMU.xlsx" dosyasının olup olmadığını kontrol et, varsa aşağıdaki kodları çalıştır.
    os.remove("TUMU.xlsx")          # Klasör içindeki "TUMU.xlsx" isimli dosyayı sil.
    dosyalar.remove("TUMU.xlsx")    # "TUMU.xlsx" isimli öğeyi "dosyalar" listesinden çıkar.

# # print("Klasördeki dosya isimleri:", dosyalar)

##### PARAMETRELER #####
sayfa_adi = 0						# "None" değeri tüm sayfa verilerini kullanır.
sayfa_adi_girdi = input("Sayfa Adı Belirtin (ya da ENTER tuşu ile geçin. 'None' ibaresi 'tüm sayfalar'ı ifade eder): ")
if sayfa_adi_girdi:
	sayfa_adi = sayfa_adi_girdi


satir_atla = None
satir_atla_girdi = input("Atlanacak satır sayısı belirtin (İlk satır BAŞLIK satırı ise, ENTER tuşu ile geçin): ")
if satir_atla_girdi.isdigit():
	satir_atla = int(satir_atla_girdi)
# # print(satir_atla)


sutunlar_girdi = input("Sütun aralığı belirtin. (Ör. A:D) (ya da tüm veri barındıran sütunları kullanmak için ENTER tuşu ile geçin): ")
sutunlar = None
if ":" in sutunlar_girdi:
	sutunlar = sutunlar_girdi.split(":")[0] + ":" + sutunlar_girdi.split(":")[1]
else:
	print("Sütun bildirimi hatalı.")
# # print(sutunlar)
#####

##### BASLIK BELİRLEME FONKSİYONU
baslik = list()

def baslik_belirt():
	df = pd.read_excel(dosyalar[0],				# İlk dosyanın basligini al. Diğer dosyalardaki başlıklar silinecek, bu başlık eklenecek.
					sheet_name = sayfa_adi,
					skiprows = 0 if satir_atla == None else satir_atla - 1,
					usecols = sutunlar)
	return df.columns

baslik = baslik_belirt()
print("Kullanılacak tablo başlığı:", baslik)
#####

#####
def VeriCercevesiBasliksiz(dosya_adi):			# Belirtilen dosya adına göre, dosya içeriğini Başlıksız DataFrame'e çeviren fonksiyon.
	df = pd.read_excel(dosya_adi,
						sheet_name = sayfa_adi,
						skiprows = 1 if satir_atla == None else satir_atla,
						usecols = sutunlar,
						header = None,			# Dosya adındaki başlık silinecek,
						names = baslik)			# klasördeki ilk dosyanın başlığı tüm diğer dosyalara başlık olarak eklenecek.
	return df
#####

#####
def tumunu_calistir():
	for dosya in dosyalar:
		if dosya == dosyalar[0]:
			data_frame = VeriCercevesiBasliksiz(dosya)
			data_frame["Dosya_Adi"] = dosya

		else:
			df_gecici = VeriCercevesiBasliksiz(dosya)
			df_gecici["Dosya_Adi"] = dosya
			data_frame = pd.concat([data_frame, df_gecici])

	print(data_frame)
	data_frame.to_excel("TUMU.xlsx")    # Tüm dosyalar birleştirildikten sonra sonuç "TUMU.xlsx" ismi ile kaydedilir.
#####

tumunu_calistir()
