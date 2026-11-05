"""PDF'i WhatsApp Masaüstü'nden gönderme ve yazdırma."""
import os
import re
import subprocess
import time

import otomasyon as ot


def telefon_duzelt(tel: str) -> str:
    """'0532 123 45 67' -> '905321234567'."""
    t = re.sub(r"\D", "", tel or "")
    if t.startswith("0"):
        t = "90" + t[1:]
    elif len(t) == 10:
        t = "90" + t
    return t


def dosya_bekle(yol, sure=180):
    """Dosya oluşup boyutu sabitlenene kadar bekler."""
    bitis = time.time() + sure
    onceki = -1
    while time.time() < bitis:
        if os.path.exists(yol):
            boyut = os.path.getsize(yol)
            if boyut > 0 and boyut == onceki:
                return True
            onceki = boyut
        time.sleep(1)
    return False


def _dosyayi_panoya_kopyala(yol):
    komut = f"Set-Clipboard -Path '{yol.replace(chr(39), chr(39) * 2)}'"
    subprocess.run(["powershell", "-NoProfile", "-Command", komut],
                   check=True, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))


def whatsapp_gonder(tel, pdf, ayar):
    wa = ayar.get("whatsapp", {})
    numara = telefon_duzelt(tel)
    if len(numara) < 11:
        raise ValueError(f"Telefon numarası geçersiz: {tel}")
    os.startfile(f"whatsapp://send?phone={numara}")  # sohbeti açar
    time.sleep(wa.get("bekle_sn", 6))
    _dosyayi_panoya_kopyala(pdf)
    ot.tus("ctrl", "v")
    time.sleep(wa.get("onizleme_bekle_sn", 3))
    if wa.get("mesaj"):
        ot.yaz(wa["mesaj"])
    ot.tus("enter")
    time.sleep(2)


def _sumatra_bul(ayar):
    adaylar = [
        ayar.get("yazdir", {}).get("sumatra_yolu", ""),
        os.path.expandvars(r"%LOCALAPPDATA%\SumatraPDF\SumatraPDF.exe"),
        r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
        r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
    ]
    return next((a for a in adaylar if a and os.path.exists(a)), None)


def yazdir(pdf, ayar):
    y = ayar.get("yazdir", {})
    kopya = int(y.get("kopya", 2))
    yazici = y.get("yazici", "")
    sumatra = _sumatra_bul(ayar)
    if sumatra:
        hedef = ["-print-to", yazici] if yazici else ["-print-to-default"]
        subprocess.run([sumatra, *hedef, "-print-settings", f"{kopya}x", "-silent", pdf],
                       check=True)
    else:  # varsayılan PDF programıyla yazdır
        for _ in range(kopya):
            os.startfile(pdf, "print")
            time.sleep(5)
