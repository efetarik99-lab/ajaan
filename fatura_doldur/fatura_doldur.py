"""Hasel Plastik – Sipariş föyü fotoğrafından ETA satış faturası doldurma.

Akış: Fotoğraf seç -> Firma yaz -> "Fotoğrafı Oku" -> tabloyu kontrol et/düzelt
      -> ETA'da 3-SATIŞ ekranını aç -> "ETA'ya Doldur".
Program ASLA kaydetmez (F2'ye basmaz). Kontrol edip kaydetmek sizde.
"""
import os
import sys
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import cekirdek as ck

KLASOR = os.path.dirname(os.path.abspath(sys.argv[0]))
AYAR_YOLU = os.path.join(KLASOR, "config.json")


def tl(x):
    if x is None:
        return ""
    return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


class Uygulama(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Hasel – Sipariş Föyünden ETA Faturası")
        self.geometry("1100x720")
        self.ayar = ck.ayarlari_yukle(AYAR_YOLU)
        self.ham = []          # Claude'dan gelen tüm satırlar
        self.satirlar = []     # seçili firmanın hazırlanmış satırları
        self.foto = None
        self._arayuz()
        self._excel_yukle()

    # ------------------------------------------------------------ arayüz
    def _arayuz(self):
        ust = ttk.Frame(self, padding=8)
        ust.pack(fill="x")
        ttk.Button(ust, text="📷 Fotoğraf Seç", command=self.foto_sec).grid(row=0, column=0)
        self.foto_lbl = ttk.Label(ust, text="(fotoğraf seçilmedi)", width=40)
        self.foto_lbl.grid(row=0, column=1, sticky="w", padx=6)
        ttk.Label(ust, text="Firma:").grid(row=0, column=2)
        self.firma = tk.StringVar()
        ttk.Entry(ust, textvariable=self.firma, width=22).grid(row=0, column=3, padx=4)
        self.oku_btn = ttk.Button(ust, text="🔍 Fotoğrafı Oku", command=self.oku)
        self.oku_btn.grid(row=0, column=4, padx=4)

        ttk.Label(ust, text="Kur (₺/$):").grid(row=1, column=0, sticky="e", pady=6)
        self.kur = tk.StringVar(value=str(self.ayar.get("kur") or ""))
        ttk.Entry(ust, textvariable=self.kur, width=10).grid(row=1, column=1, sticky="w")
        ttk.Label(ust, text="İndirim (cent):").grid(row=1, column=2, sticky="e")
        self.cent = tk.StringVar(value=str(self.ayar.get("cent_indirim") or 0))
        ttk.Entry(ust, textvariable=self.cent, width=8).grid(row=1, column=3, sticky="w", padx=4)
        self.kaynak = tk.StringVar(value=self.ayar.get("fiyat_kaynagi", "liste"))
        ttk.Radiobutton(ust, text="Fiyat listesinden", variable=self.kaynak,
                        value="liste", command=self.yenile).grid(row=1, column=4, sticky="w")
        ttk.Radiobutton(ust, text="Föydeki fiyat", variable=self.kaynak,
                        value="foy", command=self.yenile).grid(row=1, column=5, sticky="w")
        ttk.Button(ust, text="↻ Hesapla", command=self.yenile).grid(row=1, column=6, padx=4)

        orta = ttk.Frame(self, padding=(8, 0))
        orta.pack(fill="both", expand=True)
        kolonlar = ("ebat", "cins", "stok", "kg", "fiyat", "tutar", "uyari")
        basliklar = ("Açıklama (ebat)", "Cins", "Stok Kodu", "KG", "₺/kg", "Tutar", "Uyarı")
        genislik = (110, 80, 140, 80, 80, 110, 420)
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
        ttk.Button(alt, text="➖ Seçili Satırı Sil", command=self.satir_sil).pack(side="left", padx=4)
        self.toplam_lbl = ttk.Label(alt, text="", font=("Segoe UI", 10, "bold"))
        self.toplam_lbl.pack(side="left", padx=20)
        ttk.Button(alt, text="⚙ ETA Kalibrasyonu", command=self.kalibrasyon).pack(side="right")
        self.doldur_btn = ttk.Button(alt, text="▶ ETA'ya Doldur", command=self.doldur)
        self.doldur_btn.pack(side="right", padx=6)

        self.durum = tk.StringVar(value="Hazır. Çift tıklayarak hücreleri düzeltebilirsiniz.")
        ttk.Label(self, textvariable=self.durum, relief="sunken", anchor="w",
                  padding=4).pack(fill="x", side="bottom")

    # ------------------------------------------------------------ veri
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

    def foto_sec(self):
        yol = filedialog.askopenfilename(
            title="Sipariş föyü fotoğrafı",
            filetypes=[("Resim", "*.jpg *.jpeg *.png *.webp *.heic *.bmp"), ("Hepsi", "*.*")])
        if yol:
            self.foto = yol
            self.foto_lbl.config(text=os.path.basename(yol))

    def oku(self):
        if not self.foto:
            messagebox.showinfo("Fotoğraf", "Önce fotoğraf seçin.")
            return
        self.oku_btn.state(["disabled"])
        self.durum.set("Fotoğraf okunuyor (20-60 sn sürebilir)...")

        def is_parcacigi():
            try:
                ham = ck.fotograftan_oku(self.foto, self.ayar.get("api_key"))
                self.after(0, self._okundu, ham, None)
            except Exception as e:  # noqa: BLE001
                self.after(0, self._okundu, None, e)

        threading.Thread(target=is_parcacigi, daemon=True).start()

    def _okundu(self, ham, hata):
        self.oku_btn.state(["!disabled"])
        if hata:
            self.durum.set("Okuma başarısız.")
            messagebox.showerror("Hata", str(hata))
            return
        self.ham = ham
        firmalar = sorted({h["firma"] for h in ham if h.get("firma")})
        secili = ck.firma_filtrele(ham, self.firma.get())
        if not secili:
            messagebox.showwarning(
                "Firma", "Bu firma föyde bulunamadı. Föydeki firmalar:\n\n" + "\n".join(firmalar))
            self.durum.set(f"{len(ham)} satır okundu, firma eşleşmedi.")
            return
        self.ham_secili = secili
        self.yenile()
        self.durum.set(f"{len(ham)} satır okundu, {len(secili)} satır seçildi. "
                       "Sarı satırları mutlaka kontrol edin.")

    def _kur_cent(self):
        kur = ck.sayi(self.kur.get())
        cent = ck.sayi(self.cent.get()) or 0.0
        return kur, cent

    def yenile(self):
        if not getattr(self, "ham_secili", None):
            return
        kur, cent = self._kur_cent()
        if self.kaynak.get() == "liste" and not kur:
            messagebox.showwarning("Kur", "Fiyat listesinden hesap için kuru girin.")
            return
        self.satirlar = ck.satirlari_hazirla(self.ham_secili, self.ayar, kur or 0, cent,
                                             self.tablolar, self.kaynak.get())
        self.ciz()

    def ciz(self):
        self.tablo.delete(*self.tablo.get_children())
        toplam = 0.0
        for i, s in enumerate(self.satirlar):
            uyari = list(s.uyari)
            if s.fiyat is None and not any("fiyat" in u or "tablo" in u for u in uyari):
                uyari.append("fiyat yok")
            self.tablo.insert("", "end", iid=str(i), tags=("uyari",) if uyari else (),
                              values=(s.ebat, s.cins, s.stok_kodu,
                                      "" if s.kg is None else tl(s.kg).rstrip("0").rstrip(","),
                                      tl(s.fiyat), tl(s.tutar), " | ".join(uyari)))
            toplam += s.tutar or 0
        kdv = round(toplam * self.ayar.get("kdv_orani", 20) / 100, 2)
        self.toplam_lbl.config(
            text=f"Mal Hizmet Toplamı: {tl(toplam)} ₺   KDV: {tl(kdv)} ₺   "
                 f"Genel Toplam: {tl(toplam + kdv)} ₺")

    # ------------------------------------------------------------ düzenleme
    def hucre_duzenle(self, olay):
        iid = self.tablo.identify_row(olay.y)
        kol = self.tablo.identify_column(olay.x)
        if not iid or kol not in ("#1", "#2", "#3", "#4", "#5"):
            return
        x, y, w, h = self.tablo.bbox(iid, kol)
        s = self.satirlar[int(iid)]
        alan = {"#1": "ebat", "#2": "cins", "#3": "stok_kodu", "#4": "kg", "#5": "fiyat"}[kol]
        deger = getattr(s, alan)
        e = ttk.Entry(self.tablo)
        e.insert(0, "" if deger is None else str(deger).replace(".", ",")
                 if alan in ("kg", "fiyat") else str(deger))
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
            s.uyari = [u for u in s.uyari if not u.startswith("okuma şüpheli")
                       and not (alan == "kg" and u.startswith("KG yok"))
                       and not (alan == "fiyat")]
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

    # ------------------------------------------------------------ ETA
    def kalibrasyon(self):
        import pyautogui

        adimlar = [
            ("stok_kodu", "1. satırın STOK KODU hücresi"),
            ("aciklama", "1. satırın AÇIKLAMA hücresi"),
            ("miktar", "1. satırın MİKTAR hücresi"),
            ("fiyat", "1. satırın FİYAT hücresi"),
            ("ikinci", "2. satırın STOK KODU hücresi (satır yüksekliği için)"),
        ]
        if not messagebox.askokcancel(
                "Kalibrasyon",
                "ETA'da 3-SATIŞ ekranını açın ve büyütün (tam ekran).\n\n"
                "Her adımda fareyi söylenen hücrenin ÜZERİNE getirip bekleyin;\n"
                "5 saniye sonra konum kaydedilir. Tıklamayın.\n\nBaşlayalım mı?"):
            return
        eta = self.ayar["eta"]
        for anahtar, tarif in adimlar:
            for kalan in range(5, 0, -1):
                self.durum.set(f"Fareyi götürün: {tarif}  ({kalan})")
                self.update()
                time.sleep(1)
            x, y = pyautogui.position()
            if anahtar == "ikinci":
                eta["satir_yuksekligi"] = y - eta["stok_kodu"][1]
            else:
                eta[anahtar] = [x, y]
        ck.ayarlari_kaydet(AYAR_YOLU, self.ayar)
        self.durum.set(f"Kalibrasyon kaydedildi: {eta}")
        messagebox.showinfo("Kalibrasyon", "Konumlar kaydedildi.")

    def _sayi_yaz(self, x):
        if x is None:
            return ""
        t = f"{x:.2f}".rstrip("0").rstrip(".")
        return t.replace(".", self.ayar.get("ondalik_ayirac", ","))

    def doldur(self):
        eta = self.ayar["eta"]
        if not eta["satir_yuksekligi"] or eta["stok_kodu"] == [0, 0]:
            messagebox.showwarning("Kalibrasyon", "Önce ⚙ ETA Kalibrasyonu yapın.")
            return
        if not self.satirlar:
            return
        eksik = [s.ebat for s in self.satirlar if not s.stok_kodu or s.fiyat is None]
        if eksik:
            messagebox.showwarning("Eksik", "Stok kodu/fiyatı eksik satırlar var:\n"
                                   + ", ".join(eksik))
            return
        if not messagebox.askokcancel(
                "ETA'ya Doldur",
                f"{len(self.satirlar)} satır ETA'ya yazılacak.\n\n"
                "• ETA'da 3-SATIŞ ekranı açık, cari seçili ve satırlar boş olmalı.\n"
                "• Pencere kalibrasyondaki konumda/boyutta olmalı.\n"
                "• Durdurmak için fareyi ekranın SOL ÜST köşesine götürün.\n"
                "• Program kaydetmez; bitince kontrol edip F2'ye siz basın."):
            return
        self.iconify()
        threading.Thread(target=self._doldur_calis, daemon=True).start()

    def _doldur_calis(self):
        import pyautogui
        import pyperclip

        eta = self.ayar["eta"]
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = eta.get("tus_arasi_sn", 0.25)

        def yapistir(metin):
            pyperclip.copy(metin)
            pyautogui.hotkey("ctrl", "v")

        try:
            time.sleep(eta.get("baslamadan_bekle_sn", 3))
            for i, s in enumerate(self.satirlar):
                sira = min(i, eta.get("gorunur_satir", 15) - 1)
                dy = sira * eta["satir_yuksekligi"]
                pyautogui.click(eta["stok_kodu"][0], eta["stok_kodu"][1] + dy)
                yapistir(s.stok_kodu)
                pyautogui.press("enter")
                pyautogui.press("enter")
                time.sleep(0.4)
                pyautogui.click(eta["aciklama"][0], eta["aciklama"][1] + dy)
                yapistir(s.ebat)
                pyautogui.press("enter")
                pyautogui.click(eta["miktar"][0], eta["miktar"][1] + dy)
                if s.kg is not None:
                    pyautogui.write(self._sayi_yaz(s.kg))
                pyautogui.press("enter")
                pyautogui.click(eta["fiyat"][0], eta["fiyat"][1] + dy)
                pyautogui.write(self._sayi_yaz(s.fiyat))
                pyautogui.press("enter")
                time.sleep(0.4)
            mesaj = ("Doldurma bitti. ETA'da satırları ve toplamları kontrol edin, "
                     "doğruysa F2 ile siz kaydedin.")
            hata = False
        except pyautogui.FailSafeException:
            mesaj, hata = "Durduruldu (fare sol üst köşeye götürüldü).", True
        except Exception as e:  # noqa: BLE001
            mesaj, hata = f"Hata: {e}", True
        self.after(0, self._doldurma_bitti, mesaj, hata)

    def _doldurma_bitti(self, mesaj, hata):
        self.deiconify()
        self.durum.set(mesaj)
        (messagebox.showwarning if hata else messagebox.showinfo)("ETA", mesaj)


if __name__ == "__main__":
    Uygulama().mainloop()
