# Sipariş Föyünden ETA Faturası (Windows)

Föy fotoğrafını okur → ETA 3-SATIŞ ekranını doldurur → **onayınızı bekler** →
kaydeder + e-Fatura/e-Arşiv gönderir + PDF alır → PDF'i WhatsApp'tan seçtiğiniz kişiye
atar → 2 kopya yazdırır.

## 1. Kurulum (bir kez)
1. Python 3.11+ kurun (python.org, kurulumda **"Add to PATH"** işaretli).
2. **SumatraPDF** kurun (sessiz ve doğru kopya sayısıyla yazdırma için; ücretsiz).
3. WhatsApp Masaüstü kurulu ve oturum açık olsun.
4. Klasörü kopyalayın (ör. `C:\FaturaDoldur`), `config.json`'ı Not Defteri ile açın:
   - `api_key`: Anthropic API anahtarı (fotoğraf okuma için).
   - `urunler > SEDEFLI > stok_kodu`: Sedefli bantlı OPP torbanın ETA stok kodu.
   - `excel`: `Hasel Fiyat Listesi TL.xlsx` yolu, kur hücresi, her tablonun sayfa adı ve
     sol üst fiyat hücresi (en 11-15 / boy 10). Ya da `fiyat_tablolari_usd`'ye elle yazın.
   - `pdf_klasoru`: PDF'lerin kaydedileceği klasör.
   - `yazdir > kopya`: 2. `yazici` boşsa varsayılan yazıcı.
   - `whatsapp > mesaj`: PDF ile giden yazı.
5. `BASLAT.bat`'a çift tıklayın.

## 2. ETA kalibrasyonu (bir kez)
ETA'da boş 3-SATIŞ ekranını büyütün → **⚙ ETA Kalibrasyonu**. Fareyi sırayla
CARİ KODU alanı, 1. satır STOK KODU / AÇIKLAMA / MİKTAR / FİYAT, 2. satır STOK KODU
üzerinde 5'er saniye tutun.

## 3. Kaydet/Gönder makrosu (bir kez)
Program kaydetme + e-belge gönderme + PDF alma adımlarını sizi bir kez izleyerek öğrenir.
Doldurulmuş gerçek bir faturada **⏺ Kaydet/Gönder Makrosu**'na basın ve her zamanki gibi:
- F2 ile kaydedin, e-Fatura/e-Arşiv'i gönderin, PDF'i kaydetme penceresini açın,
- dosya adı yerine **F9**'a basın, sonra Kaydet'e tıklayın,
- süresi değişken bir bekleme varsa (GİB yanıtı vb.) orada **F11**'e basın,
- bitince **F10**.
Ekranlar her faturada aynı yerde açılmalı; ETA penceresini büyütülmüş kullanın.

## 4. Günlük kullanım
1. **📷 Fotoğraf Seç**, **Firma** seçin/yazın (telefon ve cari kodu rehberden gelir;
   yeni firmada bir kez yazın, otomatik hatırlanır), **🔍 Fotoğrafı Oku**.
2. Kur / indirim → **↻ Hesapla**. Sarı satırları çift tıklayıp düzeltin.
3. ETA'da boş 3-SATIŞ ekranı açıkken **▶ ETA'ya Doldur**.
4. Açılan **Onay** penceresindeki toplamları ETA ile karşılaştırın:
   - **✔ ONAYLA** → kaydet, gönder, PDF, WhatsApp, 2 kopya yazdır.
   - **✖ Vazgeç** → hiçbir şey kaydedilmez; ETA'da düzeltip kendiniz devam edebilirsiniz.
5. Bir şey ters giderse PDF'i ETA'dan elle kaydedip
   **📤 Hazır PDF'i Gönder + Yazdır** ile gönderin.

**Acil durdurma:** fareyi ekranın **sol üst köşesine** götürün.

## Kurallar
- Miktar = föydeki KG; KG yoksa boş bırakılır, fiyat yine yazılır.
- Bopp → OPP02, Sed → sedefli kod; açıklama `13*20`.
- Fiyat = (USD liste − indirim cent/100) × kur, 2 hane.
