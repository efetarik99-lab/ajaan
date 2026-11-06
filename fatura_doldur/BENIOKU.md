# Sipariş Föyünden ETA Faturası (Windows)

Föy fotoğrafını Claude veya ChatGPT'ye okutup yazısını yapıştırırsınız → ETA 3-SATIŞ ekranını doldurur → **onayınızı bekler** →
kaydeder + e-Fatura/e-Arşiv gönderir + PDF alır → PDF'i WhatsApp'tan seçtiğiniz kişiye
atar → 2 kopya yazdırır.

## 1. Kurulum (bir kez)
1. Zip'e sağ tıklayın → **Tümünü ayıkla** → `C:\FaturaDoldur` klasörüne çıkarın.
2. Klasördeki **KUR.bat**'a çift tıklayın. Python, gerekli paketler ve SumatraPDF
   kendiliğinden kurulur, masaüstüne **Fatura Doldur** simgesi gelir.
   (Windows "korumalı" uyarısı verirse: *Ek bilgi* → *Yine de çalıştır*.)
3. Program ilk açılışta sedefli ürünün stok kodunu ve fiyat listesi Excel'ini sorar.
4. WhatsApp Masaüstü kurulu ve oturumu açık olsun.

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
1. **📋 İstemi Kopyala** → Claude veya ChatGPT'yi açın, föy fotoğrafını ekleyin,
   Ctrl+V ile istemi yapıştırıp gönderin. Gelen cevabı (| ile ayrılmış satırlar) kopyalayın.
   (İstemi bir kez Claude Projesi / ChatGPT özel talimatı olarak kaydederseniz sadece
   fotoğrafı atmanız yeter.)
2. **📝 Yazıyı Yapıştır** → Tamam. **Firma** seçin/yazın (telefon ve cari kodu rehberden
   gelir; yeni firmada bir kez yazın, otomatik hatırlanır).
3. Kur / indirim → **↻ Hesapla** (firmayı değiştirince de buna basın). Sarı satırları çift tıklayıp düzeltin.
4. ETA'da boş 3-SATIŞ ekranı açıkken **▶ ETA'ya Doldur**.
5. Açılan **Onay** penceresindeki toplamları ETA ile karşılaştırın:
   - **✔ ONAYLA** → kaydet, gönder, PDF, WhatsApp, 2 kopya yazdır.
   - **✖ Vazgeç** → hiçbir şey kaydedilmez; ETA'da düzeltip kendiniz devam edebilirsiniz.
6. Bir şey ters giderse PDF'i ETA'dan elle kaydedip
   **📤 Hazır PDF'i Gönder + Yazdır** ile gönderin.

**Acil durdurma:** fareyi ekranın **sol üst köşesine** götürün.

## Kurallar
- Miktar = föydeki KG; KG yoksa boş bırakılır, fiyat yine yazılır.
- Bopp → OPP02, Sed → sedefli kod; açıklama `13*20`.
- Fiyat = (USD liste − indirim cent/100) × kur, 2 hane.
