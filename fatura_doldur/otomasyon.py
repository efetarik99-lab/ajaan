"""Windows klavye/fare otomasyonu: ETA doldurma, makro kaydet/oynat."""
import ctypes
import sys
import time

import pyautogui

pyautogui.FAILSAFE = True  # fareyi sol üst köşeye götürmek her şeyi durdurur

# ------------------------------------------------------------------ yazma
# Klavye düzeninden bağımsız yazmak için Unicode SendInput kullanılır
# (Türkçe Q klavyede pyautogui ',' yerine 'ö' basabiliyor).
if sys.platform == "win32":
    from ctypes import wintypes

    _ULONG_PTR = ctypes.POINTER(ctypes.c_ulong)

    class _KI(ctypes.Structure):
        _fields_ = [("wVk", wintypes.WORD), ("wScan", wintypes.WORD),
                    ("dwFlags", wintypes.DWORD), ("time", wintypes.DWORD),
                    ("dwExtraInfo", _ULONG_PTR)]

    class _MI(ctypes.Structure):
        _fields_ = [("dx", wintypes.LONG), ("dy", wintypes.LONG),
                    ("mouseData", wintypes.DWORD), ("dwFlags", wintypes.DWORD),
                    ("time", wintypes.DWORD), ("dwExtraInfo", _ULONG_PTR)]

    class _U(ctypes.Union):
        _fields_ = [("ki", _KI), ("mi", _MI)]

    class _INPUT(ctypes.Structure):
        _fields_ = [("type", wintypes.DWORD), ("u", _U)]

    def _unicode_tus(ch, yukari):
        inp = _INPUT(type=1)
        inp.u.ki = _KI(0, ord(ch), 0x0004 | (0x0002 if yukari else 0), 0, None)
        ctypes.windll.user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(_INPUT))


def yaz(metin: str, aralik: float = 0.02):
    """Metni aktif alana yazar."""
    if sys.platform != "win32":
        pyautogui.write(metin, interval=aralik)
        return
    for ch in metin:
        _unicode_tus(ch, False)
        _unicode_tus(ch, True)
        time.sleep(aralik)


def tus(*tuslar):
    if len(tuslar) == 1:
        pyautogui.press(tuslar[0])
    else:
        pyautogui.hotkey(*tuslar)


def tikla(x, y, sag=False):
    pyautogui.click(x, y, button="right" if sag else "left")


# ------------------------------------------------------------------ ETA doldurma
def sayi_metni(x, ayirac=","):
    if x is None:
        return ""
    return f"{x:.2f}".rstrip("0").rstrip(".").replace(".", ayirac)


def eta_doldur(satirlar, cari_kodu, ayar, durum=lambda m: None):
    eta = ayar["eta"]
    pyautogui.PAUSE = eta.get("tus_arasi_sn", 0.25)
    ayirac = ayar.get("ondalik_ayirac", ",")
    time.sleep(eta.get("baslamadan_bekle_sn", 3))

    if cari_kodu and eta.get("cari_kodu") and eta["cari_kodu"] != [0, 0]:
        durum("Cari seçiliyor...")
        tikla(*eta["cari_kodu"])
        yaz(cari_kodu)
        tus("enter")
        time.sleep(1.0)

    for i, s in enumerate(satirlar):
        durum(f"Satır {i + 1}/{len(satirlar)}: {s.ebat}")
        sira = min(i, eta.get("gorunur_satir", 15) - 1)
        dy = sira * eta["satir_yuksekligi"]
        tikla(eta["stok_kodu"][0], eta["stok_kodu"][1] + dy)
        yaz(s.stok_kodu)
        tus("enter")
        tus("enter")
        time.sleep(0.4)
        tikla(eta["aciklama"][0], eta["aciklama"][1] + dy)
        yaz(s.ebat)
        tus("enter")
        tikla(eta["miktar"][0], eta["miktar"][1] + dy)
        if s.kg is not None:
            yaz(sayi_metni(s.kg, ayirac))
        tus("enter")
        tikla(eta["fiyat"][0], eta["fiyat"][1] + dy)
        yaz(sayi_metni(s.fiyat, ayirac))
        tus("enter")
        time.sleep(0.4)


# ------------------------------------------------------------------ makro
DURDUR_TUSU = "f10"   # kaydı bitir
PDF_TUSU = "f9"       # kaydetme penceresinde: PDF dosya yolunu buraya yaz
BEKLE_TUSU = "f11"    # oynatırken burada durup kullanıcıya "Devam?" sor


class MakroKaydedici:
    """Kullanıcının ETA'da yaptığı tıklama ve tuşları kaydeder."""

    def __init__(self, bitince):
        self.adimlar = []
        self.bitince = bitince
        self._son = time.time()
        self._basili = set()
        self._fare = self._klavye = None

    def _gecen(self):
        simdi = time.time()
        dt, self._son = simdi - self._son, simdi
        return round(dt, 2)

    def _ekle(self, adim):
        adim["bekle"] = self._gecen()
        self.adimlar.append(adim)

    def baslat(self):
        from pynput import keyboard, mouse

        self._k = keyboard

        def tikladi(x, y, dugme, basildi):
            if basildi:
                self._ekle({"tip": "tikla", "x": x, "y": y,
                            "sag": dugme == mouse.Button.right})

        def basti(key):
            ad = self._ad(key)
            if ad in ("ctrl", "alt", "shift"):
                self._basili.add(ad)
                return
            if ad == DURDUR_TUSU:
                self.durdur()
                return False
            if ad == PDF_TUSU:
                self._ekle({"tip": "pdf_yolu"})
                return
            if ad == BEKLE_TUSU:
                self._ekle({"tip": "sor"})
                return
            mod = sorted(self._basili & {"ctrl", "alt"})
            if mod:
                self._ekle({"tip": "tus", "tuslar": mod + [ad]})
            elif len(ad) == 1:
                son = self.adimlar[-1] if self.adimlar else None
                if son and son["tip"] == "yaz":
                    son["metin"] += ad
                    self._son = time.time()
                else:
                    self._ekle({"tip": "yaz", "metin": ad})
            else:
                self._ekle({"tip": "tus", "tuslar": [ad]})

        def birakti(key):
            self._basili.discard(self._ad(key))

        self._fare = mouse.Listener(on_click=tikladi)
        self._klavye = keyboard.Listener(on_press=basti, on_release=birakti)
        self._fare.start()
        self._klavye.start()

    def _ad(self, key):
        k = self._k.Key
        for mod, adlar in (("ctrl", (k.ctrl, k.ctrl_l, k.ctrl_r)),
                           ("alt", (k.alt, k.alt_l, k.alt_r, getattr(k, "alt_gr", None))),
                           ("shift", (k.shift, k.shift_l, k.shift_r))):
            if key in adlar:
                return mod
        if isinstance(key, k):
            return key.name
        if self._basili & {"ctrl", "alt"} and getattr(key, "vk", None):
            return chr(key.vk).lower()
        return key.char or ""

    def durdur(self):
        if self._fare:
            self._fare.stop()
        if self._klavye:
            self._klavye.stop()
        self.bitince(self.adimlar)


def makro_oynat(adimlar, pdf_yolu, carpan=1.5, sor=lambda: True, durum=lambda m: None):
    """Kaydedilmiş adımları tekrarlar. sor() False dönerse iptal edilir."""
    pyautogui.PAUSE = 0.1
    for i, a in enumerate(adimlar):
        time.sleep(min(max(a.get("bekle", 0.5), 0.3), 10) * carpan)
        durum(f"Makro adım {i + 1}/{len(adimlar)}: {a['tip']}")
        if a["tip"] == "tikla":
            tikla(a["x"], a["y"], a.get("sag", False))
        elif a["tip"] == "yaz":
            yaz(a["metin"])
        elif a["tip"] == "tus":
            tus(*a["tuslar"])
        elif a["tip"] == "pdf_yolu":
            tus("ctrl", "a")
            yaz(pdf_yolu)
        elif a["tip"] == "sor":
            if not sor():
                raise RuntimeError("Kullanıcı makroyu durdurdu.")
