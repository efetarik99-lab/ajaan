"""Hasel Plastik – Sipariş föyü fotoğrafından ETA satış faturası.

Akış: Fotoğraf seç -> Firma seç -> Oku -> tabloyu kontrol et -> ETA'ya Doldur
      -> ETA'da kontrol et -> ONAYLA -> (kaydet + e-belge gönder + PDF) makrosu
      -> PDF WhatsApp'tan gider -> yazdırılır.
Onay vermeden hiçbir şey kaydedilmez/gönderilmez.
"""
import datetime as dt
import json
import os
import re
import sys
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import cekirdek as ck

KLASOR = os.path.dirname(os.path.abspath(sys.argv[0]))
AYAR_YOLU = os.path.join(KLASOR, "config.json")
REHBER_YOLU = os.path.join(KLASOR, "rehber.json")


def tl(x):
    if x is None:
        return ""
    return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def rehber_yukle():
    if os.path.exists(REHBER_YOLU):
        with open(REHBER_YOLU, encoding="utf-8") as f:
            return json.load(f)
    return {}


class Uygulama(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Hasel – Sipariş Föyünden ETA Faturası")
        self.geometry("1150x740")
        self.ayar = ck.ayarlari_yukle(AYAR_YOLU)
        self.rehber = rehber_yukle()
        self.ham_secili = []
        self.satirlar = []
        self.foto = None
        self._arayuz()
        self._excel_yukle()

    # ================================================================ arayüz
    def _arayuz(self):
        ust = ttk.Frame(self, padding=8)
        ust.pack(fill="x")
        ttk.Button(ust, text="📷 Fotoğraf Seç", command=self.foto_sec).grid(row=0, column=0)
        self.foto_lbl = ttk.Label(ust, text="(fotoğraf seçilmedi)", width=34)
        self.foto_lbl.grid(row=0, column=1, sticky="w", padx=6)
        ttk.Label(ust, text="Firma:").grid(row=0, column=2, sticky="e")
        self.firma = tk.StringVar()
        self.firma_cb = ttk.Combobox(ust, textvariable=self.firma, width=24,
                                     values=sorted(self.rehber))
        self.firma_cb.grid(row=0, column=3, padx=4)
        self.firma_cb.bind("<<ComboboxSelected>>", self.firma_secildi)
        self.firma_cb.bind("<FocusOut>", self.firma_secildi)
        self.oku_btn = ttk.Button(ust, text="🔍 Fotoğrafı Oku", command=self.oku)
        self.oku_btn.grid(row=0, column=4, padx=4)

        ttk.Label(ust, text="WhatsApp tel:").grid(row=0, column=5, sticky="e")
        self.tel = tk.StringVar()
        ttk.Entry(ust, textvariable=self.tel, width=16).grid(row=0, column=6, padx=4)
        ttk.Label(ust, text="Cari kodu:").grid(row=0, column=7, sticky="e")
        self.cari = tk.StringVar()
        ttk.Entry(ust, textvariable=self.cari, width=12).grid(row=0, column=8, padx=4)

        ttk.Label(ust, text="Kur (₺/$):").grid(row=1, column=0, sticky="e", pady=6)
        self.kur = tk.StringVar(value=str(self.ayar.get("kur") or ""))
        ttk.Entry(ust, textvariable=self.kur, width=10).grid(row=1, column=1, sticky="w")
        ttk.Label(ust, text="İndirim (cent):").grid(row=1, column=2, sticky="e")
        self.cent = tk.StringVar(value=str(self.ayar.get("cent_indirim") or 0))
        ttk.Entry(ust, textvariable=self.cent, width=8).grid(row=1, column=3, sticky="w", padx=4)
        ttk.Button(ust, text="↻ Hesapla", command=self.yenile).grid(row=1, column=4, padx=4)

        orta = ttk.Frame(self, padding=(8, 0))
        orta.pack(fill="both", expand=True)
        kolonlar = ("ebat", "cins", "stok", "kg", "fiyat", "tutar", "uyari")
        basliklar = ("Açıklama (ebat)", "Cins", "Stok Kodu", "KG", "₺/kg", "Tutar", "Uyarı")
        genislik = (110, 80, 140, 80, 80, 110, 450)
        self.tablo = ttk.Treeview(orta, columns=kolonlar, show="headings", height=18)
        for k, b, g in zip(kolonlar, basliklar, genislik):
            self.tablo.heading(k, text=b)
            self.tablo.column(k, width=g, anchor="w" if k in ("uyari", "stok") else "center")
        self.tablo.tag_configure("uyari", background="#fff3cd")
        self.tablo.pack(side="left", fill="both", expand=True)
        self.tablo.bind("<Double-1>", self.hucre_duzenle)
        sb = ttk.Scrollbar(orta, orient="vertical", command=self.tablo.yview)
        sb.pack(side="right", fill="y")
        self.tablo.configure(yscrollcommand=sb.set)

        alt = ttk.Frame(self, padding=8)
        alt.pack(fill="x")
        ttk.Button(alt, text="➕ Satır", command=self.satir_ekle).pack(side="left")
        ttk.Button(alt, text="➖ Sil", command=self.satir_sil).pack(side="left", padx=4)
        self.toplam_lbl = ttk.Label(alt, text="", font=("Segoe UI", 10, "bold"))
        self.toplam_lbl.pack(side="left", padx=16)
        self.doldur_btn = ttk.Button(alt, text="▶ ETA'ya Doldur", command=self.doldur)
        self.doldur_btn.pack(side="right")

        ayar = ttk.Frame(self, padding=(8, 0, 8, 6))
        ayar.pack(fill="x")
        ttk.Label(ayar, text="Kurulum:").pack(side="left")
        ttk.Button(ayar, text="⚙ ETA Kalibrasyonu", command=self.kalibrasyon).pack(side="left", padx=4)
        ttk.Button(ayar, text="⏺ Kaydet/Gönder Makrosu", command=self.makro_kaydet).pack(side="left", padx=4)
        ttk.Button(ayar, text="📤 Hazır PDF'i Gönder + Yazdır",
                   command=self.hazir_pdf).pack(side="right")

        self.durum = tk.StringVar(value="Hazır. Hücreleri çift tıklayarak düzeltebilirsiniz.")
        ttk.Label(self, textvariable=self.durum, relief="sunken", anchor="w",
                  padding=4).pack(fill="x", side="bottom")

    def durum_yaz(self, m):
        self.after(0, self.durum.set, m)

    # ================================================================ veri
    def _excel_yukle(self):
        self.tablolar = self.ayar.get("fiyat_tablolari_usd", {})
        try:
            ex_tablo, ex_kur = ck.excelden_fiyatlari_oku(self.ayar)
        except Exception as e:  # noqa: BLE001
            messagebox.showwarning("Excel", f"Fiyat listesi okunamadı:\n{e}")
            return
        if ex_tablo:
            self.tablolar = {**self.tablolar, **ex_tablo}
            self.durum.set("Fiyatlar Excel'den okundu.")
        if ex_kur and not self.kur.get():
            self.kur.set(str(ex_kur))

    def firma_secildi(self, _=None):
        k = self.rehber.get(self.firma.get())
        if k:
            self.tel.set(k.get("telefon", ""))
            self.cari.set(k.get("cari_kodu", ""))

    def rehbere_yaz(self):
        f = self.firma.get().strip()
        if not f:
            return
        self.rehber[f] = {"telefon": self.tel.get().strip(), "cari_kodu": self.cari.get().strip()}
        with open(REHBER_YOLU, "w", encoding="utf-8") as fp:
            json.dump(self.rehber, fp, ensure_ascii=False, indent=2)
        self.firma_cb["values"] = sorted(self.rehber)

    def foto_sec(self):
        yol = filedialog.askopenfilename(
            title="Sipariş föyü fotoğrafı",
            filetypes=[("Resim", "*.jpg *.jpeg *.png *.webp *.bmp"), ("Hepsi", "*.*")])
        if yol:
            self.foto = yol
            self.foto_lbl.config(text=os.path.basename(yol))

    def oku(self):
        if not self.foto:
            messagebox.showinfo("Fotoğraf", "Önce fotoğraf seçin.")
            return
        self.oku_btn.state(["disabled"])
        self.durum.set("Fotoğraf okunuyor (20-60 sn sürebilir)...")

        def is_():
            try:
                ham = ck.fotograftan_oku(self.foto, self.ayar.get("api_key"))
                self.after(0, self._okundu, ham, None)
            except Exception as e:  # noqa: BLE001
                self.after(0, self._okundu, None, e)

        threading.Thread(target=is_, daemon=True).start()

    def _okundu(self, ham, hata):
        self.oku_btn.state(["!disabled"])
        if hata:
            self.durum.set("Okuma başarısız.")
            messagebox.showerror("Hata", str(hata))
            return
        firmalar = sorted({h["firma"] for h in ham if h.get("firma")})
        secili = ck.firma_filtrele(ham, self.firma.get())
        if not secili:
            messagebox.showwarning("Firma", "Bu firma föyde bulunamadı. Föydeki firmalar:\n\n"
                                   + "\n".join(firmalar))
            self.durum.set(f"{len(ham)} satır okundu, firma eşleşmedi.")
            return
        self.ham_secili = secili
        self.yenile()
        self.durum.set(f"{len(ham)} satır okundu, {len(secili)} satır seçildi. "
                       "Sarı satırları kontrol edin.")

    def yenile(self):
        if not self.ham_secili:
            return
        kur = ck.sayi(self.kur.get())
        if not kur:
            messagebox.showwarning("Kur", "Dolar kurunu girin.")
            return
        cent = ck.sayi(self.cent.get()) or 0.0
        self.satirlar = ck.satirlari_hazirla(self.ham_secili, self.ayar, kur, cent,
                                             self.tablolar, "liste")
        self.ciz()

    def toplamlar(self):
        ara = round(sum(s.tutar or 0 for s in self.satirlar), 2)
        kdv = round(ara * self.ayar.get("kdv_orani", 20) / 100, 2)
        return ara, kdv, round(ara + kdv, 2)

    def ciz(self):
        self.tablo.delete(*self.tablo.get_children())
        for i, s in enumerate(self.satirlar):
            uyari = list(s.uyari)
            if s.fiyat is None and not uyari:
                uyari.append("fiyat yok")
            kg = "" if s.kg is None else tl(s.kg).rstrip("0").rstrip(",")
            self.tablo.insert("", "end", iid=str(i), tags=("uyari",) if uyari else (),
                              values=(s.ebat, s.cins, s.stok_kodu, kg, tl(s.fiyat),
                                      tl(s.tutar), " | ".join(uyari)))
        ara, kdv, gen = self.toplamlar()
        self.toplam_lbl.config(text=f"Mal Hizmet: {tl(ara)} ₺   KDV: {tl(kdv)} ₺   "
                                    f"Genel Toplam: {tl(gen)} ₺")

    def hucre_duzenle(self, olay):
        iid = self.tablo.identify_row(olay.y)
        kol = self.tablo.identify_column(olay.x)
        alanlar = {"#1": "ebat", "#2": "cins", "#3": "stok_kodu", "#4": "kg", "#5": "fiyat"}
        if not iid or kol not in alanlar:
            return
        x, y, w, h = self.tablo.bbox(iid, kol)
        s = self.satirlar[int(iid)]
        alan = alanlar[kol]
        d = getattr(s, alan)
        e = ttk.Entry(self.tablo)
        e.insert(0, "" if d is None else (str(d).replace(".", ",") if alan in ("kg", "fiyat")
                                          else str(d)))
        e.place(x=x, y=y, width=w, height=h)
        e.focus_set()

        def bitir(_=None):
            v = e.get()
            if alan in ("kg", "fiyat"):
                setattr(s, alan, ck.sayi(v))
            elif alan == "ebat":
                s.ebat = ck.ebat_metni(v)
            else:
                setattr(s, alan, v.strip())
            s.uyari = []
            e.destroy()
            self.ciz()

        e.bind("<Return>", bitir)
        e.bind("<FocusOut>", bitir)
        e.bind("<Escape>", lambda _: e.destroy())

    def satir_ekle(self):
        tip = next(iter(self.ayar["urunler"]))
        self.satirlar.append(ck.Satir(firma=self.firma.get(), ebat="", cins=tip,
                                      stok_kodu=self.ayar["urunler"][tip]["stok_kodu"]))
        self.ciz()

    def satir_sil(self):
        for iid in sorted(self.tablo.selection(), key=int, reverse=True):
            del self.satirlar[int(iid)]
        self.ciz()

    # ================================================================ kurulum
    def kalibrasyon(self):
        import pyautogui

        adimlar = [
            ("cari_kodu", "CARİ KODU alanı (kullanmıyorsanız boş bir yere)"),
            ("stok_kodu", "1. satırın STOK KODU hücresi"),
            ("aciklama", "1. satırın AÇIKLAMA hücresi"),
            ("miktar", "1. satırın MİKTAR hücresi"),
            ("fiyat", "1. satırın FİYAT hücresi"),
            ("ikinci", "2. satırın STOK KODU hücresi"),
        ]
        if not messagebox.askokcancel(
                "Kalibrasyon",
                "ETA'da boş bir 3-SATIŞ ekranı açıp BÜYÜTÜN.\n\n"
                "Her adımda fareyi söylenen yerin ÜZERİNE getirip bekleyin (tıklamayın);\n"
                "5 saniye sonra konum kaydedilir. Adımlar alttaki durum çubuğunda yazar."):
            return
        eta = self.ayar["eta"]
        for anahtar, tarif in adimlar:
            for kalan in range(5, 0, -1):
                self.durum.set(f"Fareyi götürün → {tarif}  ({kalan})")
                self.update()
                time.sleep(1)
            x, y = pyautogui.position()
            if anahtar == "ikinci":
                eta["satir_yuksekligi"] = y - eta["stok_kodu"][1]
            else:
                eta[anahtar] = [x, y]
        ck.ayarlari_kaydet(AYAR_YOLU, self.ayar)
        self.durum.set("Kalibrasyon kaydedildi.")
        messagebox.showinfo("Kalibrasyon", "Konumlar kaydedildi.")

    def makro_kaydet(self):
        import otomasyon as ot

        if not messagebox.askokcancel(
                "Kaydet/Gönder Makrosu",
                "Bu, faturayı KAYDETME + e-belge GÖNDERME + PDF KAYDETME adımlarını\n"
                "bir kez sizi izleyerek öğrenir. Gerçek bir faturayla yapın.\n\n"
                "1) ETA'da doldurulmuş bir fatura ekranı açık olsun.\n"
                "2) Tamam'a basınca kayıt başlar. Her zamanki gibi:\n"
                "   F2 ile kaydet → e-Fatura/e-Arşiv gönder → PDF kaydet.\n"
                "3) PDF 'Farklı kaydet' penceresinde dosya adı yazmak yerine F9'a basın,\n"
                "   sonra Kaydet'e tıklayın (program her seferinde kendi dosya adını yazar).\n"
                "4) Uzun sürebilen bir yerde (GİB onayı vb.) F11'e basın: program orada\n"
                "   durup 'Devam?' diye sorar.\n"
                "5) Bitince F10'a basın.\n\n"
                "Not: Bu pencereye/programa tıklamayın; ETA penceresi aynı yerde olsun."):
            return

        def bitti(adimlar):
            self.after(0, self._makro_bitti, adimlar)

        self.iconify()
        self._kaydedici = ot.MakroKaydedici(bitti)
        self._kaydedici.baslat()

    def _makro_bitti(self, adimlar):
        self.deiconify()
        if not any(a["tip"] == "pdf_yolu" for a in adimlar):
            messagebox.showwarning("Makro", "Kayıtta F9 (PDF yolu) yok. PDF bulunamayacağı için "
                                   "makro kaydedilmedi. Tekrar deneyin.")
            return
        self.ayar["makro"] = adimlar
        ck.ayarlari_kaydet(AYAR_YOLU, self.ayar)
        messagebox.showinfo("Makro", f"{len(adimlar)} adım kaydedildi.")

    # ================================================================ ETA + onay
    def doldur(self):
        eta = self.ayar["eta"]
        if not eta.get("satir_yuksekligi"):
            messagebox.showwarning("Kalibrasyon", "Önce ⚙ ETA Kalibrasyonu yapın.")
            return
        if not self.satirlar:
            return
        eksik = [s.ebat or "(boş)" for s in self.satirlar if not s.stok_kodu or s.fiyat is None]
        if eksik:
            messagebox.showwarning("Eksik", "Stok kodu/fiyatı eksik satırlar:\n" + ", ".join(eksik))
            return
        if not messagebox.askokcancel(
                "ETA'ya Doldur",
                f"{len(self.satirlar)} satır yazılacak.\n\n"
                "• ETA'da BOŞ bir 3-SATIŞ ekranı açık ve büyütülmüş olmalı.\n"
                "• Cari kodu boşsa cariyi önceden siz seçin.\n"
                "• Durdurmak için fareyi ekranın SOL ÜST köşesine götürün."):
            return
        self.rehbere_yaz()
        self.iconify()
        threading.Thread(target=self._doldur_is, daemon=True).start()

    def _doldur_is(self):
        import otomasyon as ot
        import pyautogui

        try:
            ot.eta_doldur(self.satirlar, self.cari.get().strip(), self.ayar, self.durum_yaz)
            self.after(0, self.onay_penceresi)
        except pyautogui.FailSafeException:
            self.after(0, self._hata, "Durduruldu (fare sol üst köşeye götürüldü).")
        except Exception as e:  # noqa: BLE001
            self.after(0, self._hata, f"Hata: {e}")

    def _hata(self, m):
        self.deiconify()
        self.durum.set(m)
        messagebox.showwarning("Durdu", m)

    def onay_penceresi(self):
        ara, kdv, gen = self.toplamlar()
        p = tk.Toplevel(self)
        p.title("Onay")
        p.attributes("-topmost", True)
        p.geometry("+20+20")
        metin = (f"ETA'yı kontrol edin.\n\nFirma: {self.firma.get()}\n"
                 f"Satır: {len(self.satirlar)}\n"
                 f"Mal Hizmet: {tl(ara)} ₺\nKDV: {tl(kdv)} ₺\nGenel Toplam: {tl(gen)} ₺\n\n"
                 f"WhatsApp: {self.tel.get() or '(YOK!)'}\n"
                 f"Yazdırma: {self.ayar.get('yazdir', {}).get('kopya', 2)} kopya")
        ttk.Label(p, text=metin, padding=12, justify="left").pack()
        f = ttk.Frame(p, padding=8)
        f.pack()

        def onayla():
            if not self.ayar.get("makro"):
                messagebox.showwarning("Makro", "Önce ⏺ Kaydet/Gönder Makrosu kaydedin.", parent=p)
                return
            if not self.tel.get().strip():
                messagebox.showwarning("Telefon", "WhatsApp numarası boş.", parent=p)
                return
            p.destroy()
            self.rehbere_yaz()
            threading.Thread(target=self._gonder_is, daemon=True).start()

        def iptal():
            p.destroy()
            self.deiconify()
            self.durum.set("Onay verilmedi; ETA'daki fatura kaydedilmedi.")

        ttk.Button(f, text="✔ ONAYLA: Kaydet, Gönder, WhatsApp, Yazdır",
                   command=onayla).pack(side="left", padx=4)
        ttk.Button(f, text="✖ Vazgeç", command=iptal).pack(side="left", padx=4)

    def _pdf_yolu(self):
        klasor = self.ayar.get("pdf_klasoru") or os.path.join(KLASOR, "Faturalar")
        os.makedirs(klasor, exist_ok=True)
        ad = re.sub(r'[\\/:*?"<>|]', "", self.firma.get().strip() or "fatura")
        return os.path.join(klasor, f"{ad}_{dt.datetime.now():%Y%m%d_%H%M%S}.pdf")

    def _sor(self):
        """Makro içindeki F11 noktasında kullanıcıya sorar (iş parçacığından)."""
        olay, sonuc = threading.Event(), {}

        def goster():
            sonuc["ok"] = messagebox.askyesno(
                "Devam?", "ETA'da işlem bitti mi? (örn. GİB gönderimi)\n\n"
                "Evet = devam et, Hayır = durdur")
            olay.set()

        self.after(0, goster)
        olay.wait()
        return sonuc["ok"]

    def _gonder_is(self):
        import gonder
        import otomasyon as ot
        import pyautogui

        pdf = self._pdf_yolu()
        try:
            self.durum_yaz("Kaydediliyor / gönderiliyor...")
            ot.makro_oynat(self.ayar["makro"], pdf, self.ayar.get("makro_hiz_carpani", 1.5),
                           self._sor, self.durum_yaz)
            self.durum_yaz("PDF bekleniyor...")
            if not gonder.dosya_bekle(pdf, self.ayar.get("pdf_bekle_sn", 180)):
                raise RuntimeError(f"PDF oluşmadı: {pdf}\nETA'yı kontrol edin; PDF'i elle "
                                   "kaydedip '📤 Hazır PDF'i Gönder + Yazdır' kullanın.")
            self._pdf_dagit(pdf)
        except pyautogui.FailSafeException:
            self.after(0, self._hata, "Durduruldu (fare sol üst köşe). ETA'daki durumu kontrol edin!")
        except Exception as e:  # noqa: BLE001
            self.after(0, self._hata, str(e))

    def _pdf_dagit(self, pdf):
        import gonder

        self.durum_yaz("WhatsApp'tan gönderiliyor...")
        gonder.whatsapp_gonder(self.tel.get(), pdf, self.ayar)
        self.durum_yaz("Yazdırılıyor...")
        gonder.yazdir(pdf, self.ayar)
        m = f"Tamamlandı. PDF: {pdf}\nWhatsApp'ta gittiğini kontrol edin."
        self.after(0, lambda: (self.deiconify(), self.durum.set(m),
                               messagebox.showinfo("Bitti", m)))

    def hazir_pdf(self):
        pdf = filedialog.askopenfilename(title="Fatura PDF", filetypes=[("PDF", "*.pdf")])
        if not pdf:
            return
        if not self.tel.get().strip():
            messagebox.showwarning("Telefon", "Önce WhatsApp numarasını girin.")
            return
        self.rehbere_yaz()
        self.iconify()

        def is_():
            try:
                self._pdf_dagit(pdf)
            except Exception as e:  # noqa: BLE001
                self.after(0, self._hata, str(e))

        threading.Thread(target=is_, daemon=True).start()


if __name__ == "__main__":
    Uygulama().mainloop()
