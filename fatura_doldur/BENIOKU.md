# Sipariş Föyünden ETA Faturası

Hasel Plastik günlük sipariş föyünün fotoğrafını okur, seçtiğiniz firmanın satırlarını
fiyat listesine göre fiyatlandırır ve ETA V8 SQL'de **3-SATIŞ** ekranına yazar.
**Program kaydetmez (F2'ye basmaz).** Kontrol edip kaydetmek sizde.

## Kurulum (bir kez)
1. https://www.python.org adresinden Python 3.11+ kurun ("Add to PATH" kutusunu işaretleyin).
2. Bu klasörü bilgisayara kopyalayın (ör. `C:\FaturaDoldur`).
3. `config.json` dosyasını Not Defteri ile açıp şunları doldurun:
   - `api_key`: Anthropic API anahtarınız (https://console.anthropic.com).
   - `urunler > SEDEFLI > stok_kodu`: Sedefli bantlı OPP torbanın ETA stok kodu.
   - **Fiyatlar**, iki yoldan biri:
     - `excel > dosya`: `Hasel Fiyat Listesi TL.xlsx` dosyasının tam yolu
       (ör. `"F:\\Fiyat\\Hasel Fiyat Listesi TL.xlsx"`), `kur_sayfa`/`kur_hucre`
       (dolar kurunun hücresi), her tablo için `sayfa` adı ve `sol_ust_hucre`
       (en 11-15 / boy 10 kesişimindeki ilk USD fiyat hücresi). Program 3 satır × 6
       sütun okur.
     - veya `fiyat_tablolari_usd` içine USD fiyatları elle yazın
       (satırlar: en 11-15, 16-22, 23+; sütunlar: boy 10, 12, 15-17, 20, 22x23, 25+).
4. `BASLAT.bat`'a çift tıklayın (ilk açılışta gerekli paketleri kurar).

## ETA kalibrasyonu (bir kez, ekran düzeni değişirse tekrar)
ETA'da 3-SATIŞ ekranını açıp **büyütün**, programda **⚙ ETA Kalibrasyonu**'na basın.
Her adımda fareyi söylenen hücrenin üzerine götürüp bekleyin (5 sn):
1. satır STOK KODU, AÇIKLAMA, MİKTAR, FİYAT ve 2. satır STOK KODU.

## Günlük kullanım
1. **📷 Fotoğraf Seç** → föy fotoğrafı.
2. **Firma** kutusuna firma adını (bir kısmı yeter) yazın, **🔍 Fotoğrafı Oku**.
3. Kur ve varsa "X cent indirim"i girin, **↻ Hesapla**.
4. Tabloyu kontrol edin. **Sarı** satırlarda uyarı var; hücreye çift tıklayıp düzeltin.
   Altta Mal Hizmet Toplamı / KDV / Genel Toplam görünür.
5. ETA'da 3-SATIŞ ekranını açın, cariyi seçin, **▶ ETA'ya Doldur**.
   Acil durdurmak için fareyi ekranın **sol üst köşesine** götürün.
6. ETA'daki toplamları programdakiyle karşılaştırın, doğruysa F2 ile kaydedin.

## Kurallar
- Faturaya KG sütunundaki gerçek kilo girer; KG yoksa miktar boş bırakılır, fiyat yine yazılır.
- Bopp → OPP02, Sed → sedefli kod. Açıklamaya ebat `13*20` biçiminde yazılır.
- Fiyat = (USD liste fiyatı − indirim cent/100) × kur, 2 haneye yuvarlanır.
- Grup sınırları (`en_gruplari`, `boy_gruplari`) `config.json`'dan değiştirilebilir.
