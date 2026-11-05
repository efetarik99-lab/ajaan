"""Fotoğraftan okuma, fiyat hesabı ve eşleştirme mantığı (GUI'den bağımsız)."""
import base64
import io
import json
import os
import re
from dataclasses import dataclass, field

MODEL = "claude-opus-5-5"

# ---------------------------------------------------------------- yardımcılar
_TR = str.maketrans("İIıŞşĞğÜüÖöÇç", "IIISSGGUUOOCC")


def normalize(s: str) -> str:
    """Türkçe karakterleri sadeleştirip büyük harfe çevirir (karşılaştırma için)."""
    return (s or "").translate(_TR).upper().strip()


def sayi(s) -> float | None:
    """'12,5' / '12.5' / '1.250,5' gibi metni sayıya çevirir; boşsa None."""
    if s is None:
        return None
    if isinstance(s, (int, float)):
        return float(s)
    t = str(s).strip().replace(" ", "")
    if not t:
        return None
    if "," in t and "." in t:
        t = t.replace(".", "").replace(",", ".")
    else:
        t = t.replace(",", ".")
    m = re.search(r"-?\d+(\.\d+)?", t)
    return float(m.group()) if m else None


def ebat_ayir(ebat: str) -> tuple[int, int] | None:
    """'13*20', '13x20', '13 X 20+3' -> (13, 20)."""
    nums = re.findall(r"\d+(?:[.,]\d+)?", ebat or "")
    if len(nums) < 2:
        return None
    return round(sayi(nums[0])), round(sayi(nums[1]))


def ebat_metni(ebat: str) -> str:
    e = ebat_ayir(ebat)
    return f"{e[0]}*{e[1]}" if e else (ebat or "").strip()


# ---------------------------------------------------------------- ayarlar
def ayarlari_yukle(yol: str) -> dict:
    with open(yol, encoding="utf-8") as f:
        return json.load(f)


def ayarlari_kaydet(yol: str, ayar: dict) -> None:
    with open(yol, "w", encoding="utf-8") as f:
        json.dump(ayar, f, ensure_ascii=False, indent=2)


def urun_tipi(cins: str, ayar: dict) -> str | None:
    """Föydeki CİNS yazısını ürün tipine (BANTLI / SEDEFLI ...) çevirir."""
    c = normalize(cins)
    for tip, bilgi in ayar["urunler"].items():
        for anahtar in bilgi["anahtar_kelimeler"]:
            if normalize(anahtar) in c:
                return tip
    return None


def _grup_bul(deger: int, gruplar: list[dict]) -> int | None:
    for i, g in enumerate(gruplar):
        if g["min"] <= deger <= g["max"]:
            return i
    return None


def excelden_fiyatlari_oku(ayar: dict) -> tuple[dict, float | None]:
    """Ayardaki 'excel' bölümüne göre fiyat tablolarını ve kuru Excel'den okur."""
    ex = ayar.get("excel") or {}
    dosya = ex.get("dosya")
    if not dosya or not os.path.exists(dosya):
        return {}, None
    import openpyxl

    wb = openpyxl.load_workbook(dosya, data_only=True)
    kur = None
    if ex.get("kur_sayfa") and ex.get("kur_hucre"):
        kur = sayi(wb[ex["kur_sayfa"]][ex["kur_hucre"]].value)
    n_en = len(ayar["en_gruplari"])
    n_boy = len(ayar["boy_gruplari"])
    tablolar = {}
    for tip, t in (ex.get("tablolar") or {}).items():
        ws = wb[t["sayfa"]]
        bas = ws[t["sol_ust_hucre"]]
        tablo = []
        for r in range(n_en):
            satir = []
            for c in range(n_boy):
                v = ws.cell(row=bas.row + r, column=bas.column + c).value
                satir.append(sayi(v))
            tablo.append(satir)
        tablolar[tip] = tablo
    return tablolar, kur


def birim_fiyat(tip: str, ebat: str, ayar: dict, kur: float, cent_indirim: float,
                tablolar: dict) -> tuple[float | None, str]:
    """USD tablo fiyatı × kur − indirim. (fiyat, açıklama/uyarı) döndürür."""
    e = ebat_ayir(ebat)
    if not e:
        return None, "ebat okunamadı"
    en, boy = e
    ri = _grup_bul(en, ayar["en_gruplari"])
    ci = _grup_bul(boy, ayar["boy_gruplari"])
    if ri is None or ci is None:
        return None, f"{en}*{boy} fiyat tablosunda yok"
    tablo = tablolar.get(tip)
    if not tablo:
        return None, f"{tip} fiyat tablosu yok"
    usd = tablo[ri][ci]
    if not usd:
        return None, f"{tip} tablosunda {ayar['en_gruplari'][ri]['ad']} / " \
                     f"{ayar['boy_gruplari'][ci]['ad']} boş"
    fiyat = round((usd - cent_indirim / 100.0) * kur, 2)
    return fiyat, f"{usd} USD"


# ---------------------------------------------------------------- satır modeli
@dataclass
class Satir:
    firma: str
    ebat: str
    cins: str
    siparis: str = ""
    kg: float | None = None
    foy_fiyat: float | None = None
    stok_kodu: str = ""
    fiyat: float | None = None
    uyari: list = field(default_factory=list)

    @property
    def tutar(self) -> float | None:
        if self.kg is None or self.fiyat is None:
            return None
        return round(self.kg * self.fiyat, 2)


def firma_filtrele(satirlar: list[dict], firma: str) -> list[dict]:
    if not firma.strip():
        return satirlar
    f = normalize(firma)
    return [s for s in satirlar if f in normalize(s.get("firma", ""))]


def satirlari_hazirla(ham: list[dict], ayar: dict, kur: float, cent: float,
                      tablolar: dict, fiyat_kaynagi: str = "liste") -> list[Satir]:
    sonuc = []
    for h in ham:
        s = Satir(
            firma=h.get("firma", ""),
            ebat=ebat_metni(h.get("ebat", "")),
            cins=h.get("cins", ""),
            siparis=h.get("siparis_miktari", ""),
            kg=sayi(h.get("kg")),
            foy_fiyat=sayi(h.get("fiyat")),
        )
        if h.get("supheli"):
            s.uyari.append("okuma şüpheli: " + (h.get("not") or "kontrol et"))
        tip = urun_tipi(s.cins, ayar)
        if tip is None:
            s.uyari.append(f"cins tanınmadı: {s.cins}")
        else:
            s.stok_kodu = ayar["urunler"][tip]["stok_kodu"]
            if fiyat_kaynagi == "foy" and s.foy_fiyat:
                s.fiyat = s.foy_fiyat
            else:
                s.fiyat, bilgi = birim_fiyat(tip, s.ebat, ayar, kur, cent, tablolar)
                if s.fiyat is None:
                    s.uyari.append(bilgi)
        if s.kg is None:
            s.uyari.append("KG yok (boş girilecek)")
        sonuc.append(s)
    return sonuc


# ---------------------------------------------------------------- Claude ile okuma
_SEMA = {
    "type": "object",
    "properties": {
        "satirlar": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "firma": {"type": "string"},
                    "ebat": {"type": "string"},
                    "cins": {"type": "string"},
                    "siparis_miktari": {"type": "string"},
                    "kg": {"type": "string"},
                    "fiyat": {"type": "string"},
                    "supheli": {"type": "boolean"},
                    "not": {"type": "string"},
                },
                "required": ["firma", "ebat", "cins", "siparis_miktari", "kg",
                             "fiyat", "supheli", "not"],
                "additionalProperties": False,
            },
        }
    },
    "required": ["satirlar"],
    "additionalProperties": False,
}

_TALIMAT = """Bu fotoğraf "HASEL PLASTİK Günlük Sipariş – İmalat Dökümü" föyü.
Sütunlar: FİRMA, EBAT, CİNS, SİP. MİK., KG, FİYAT. Yazılar el yazısı olabilir.

Tablodaki HER satırı oku ve şu kurallarla döndür:
- firma: Firma adı. Firma sadece ilk satıra yazılıp altındakiler boş / tırnak (") ile
  geçilmişse, aynı firma adını o satırlara da yaz.
- ebat: "en*boy" biçiminde (ör. "13*20"). Ek yazı varsa (ör. "+3") olduğu gibi bırak.
- cins: Yazıldığı gibi (ör. "Bopp", "Sed").
- siparis_miktari, kg, fiyat: Yazıldığı gibi rakam; boşsa "" (boş metin).
  KG sütunundaki rakamı sipariş miktarıyla karıştırma.
- supheli: Herhangi bir rakamı/yazıyı tam okuyamadıysan true, değilse false.
- not: supheli ise hangi hücrenin neden belirsiz olduğunu kısaca yaz, değilse "".
Tahmin uydurma; okuyamadığın yeri supheli=true ile işaretle."""


def _resmi_hazirla(yol: str) -> tuple[str, str]:
    """Resmi gerekirse küçültüp (API sınırı 5 MB) base64 JPEG olarak döndürür."""
    from PIL import Image, ImageOps

    img = Image.open(yol)
    img = ImageOps.exif_transpose(img).convert("RGB")
    img.thumbnail((2400, 2400))
    for kalite in (90, 80, 70, 60):
        buf = io.BytesIO()
        img.save(buf, "JPEG", quality=kalite)
        if buf.tell() < 4_500_000:
            break
    return base64.standard_b64encode(buf.getvalue()).decode(), "image/jpeg"


def fotograftan_oku(yol: str, api_key: str | None = None) -> list[dict]:
    import anthropic

    client = anthropic.Anthropic(api_key=api_key or None)
    veri, tur = _resmi_hazirla(yol)
    yanit = client.messages.create(
        model=MODEL,
        max_tokens=16000,
        output_config={"effort": "high",
                       "format": {"type": "json_schema", "schema": _SEMA}},
        # Güvenlik filtresi nadiren reddederse sunucu otomatik yedek modele geçer.
        extra_headers={"anthropic-beta": "server-side-fallback-2026-07-01"},
        extra_body={"fallbacks": "default"},
        messages=[{
            "role": "user",
            "content": [
                {"type": "image",
                 "source": {"type": "base64", "media_type": tur, "data": veri}},
                {"type": "text", "text": _TALIMAT},
            ],
        }],
    )
    if yanit.stop_reason == "refusal":
        raise RuntimeError("Model isteği reddetti, fotoğrafı tekrar deneyin.")
    if yanit.stop_reason == "max_tokens":
        raise RuntimeError("Yanıt yarıda kesildi (çok satır?).")
    metin = next(b.text for b in yanit.content if b.type == "text")
    return json.loads(metin)["satirlar"]


# ---------------------------------------------------------------- yapıştırılan metinden okuma
ISTEM = """Bu fotoğraf HASEL PLASTİK "Günlük Sipariş – İmalat Dökümü" föyü.
Sütunlar: FİRMA, EBAT, CİNS, SİP. MİK., KG, FİYAT.

Tablodaki HER satırı, aşağıdaki biçimde, her satır bir satıra gelecek şekilde yaz.
Başka hiçbir açıklama, başlık veya tablo çizgisi ekleme; kod bloğu içinde ver:

FİRMA | EBAT | CİNS | SİP.MİK | KG | FİYAT | NOT

Kurallar:
- Firma sadece ilk satıra yazılıp altı boş/tırnak (") ise firma adını her satıra yaz.
- EBAT "13*20" biçiminde olsun.
- Boş hücreyi boş bırak (iki | arası boş).
- KG ile sipariş miktarını karıştırma.
- Emin olmadığın hücre varsa sonuna ? koy ve NOT'a nedenini yaz. Tahmin uydurma.

Örnek:
ABC AMBALAJ | 13*20 | Bopp | 5000 | 12,5 |  |
ABC AMBALAJ | 25*35 | Sed | 2000 | 8? |  | KG 8 mi 9 mu belirsiz
"""


def metinden_satirlar(metin: str) -> list[dict]:
    """ISTEM biçimindeki (veya markdown tablo) yazıyı satır listesine çevirir."""
    sonuc, onceki_firma = [], ""
    for ham in (metin or "").splitlines():
        satir = ham.strip().strip("`").strip()
        if "|" not in satir:
            continue
        hucre = [h.strip() for h in satir.strip("|").split("|")]
        if all(re.fullmatch(r":?-{2,}:?", h) or not h for h in hucre):
            continue  # markdown ayırıcı çizgisi
        if normalize(hucre[1] if len(hucre) > 1 else "") in ("EBAT", "OLCU"):
            continue  # başlık
        hucre += [""] * (7 - len(hucre))
        firma, ebat, cins, sip, kg, fiyat, notu = hucre[:7]
        if firma in ("", '"', "''", "〃", "-") :
            firma = onceki_firma
        onceki_firma = firma
        supheli = any("?" in h for h in hucre[:6]) or bool(notu)
        sonuc.append({
            "firma": firma, "ebat": ebat.replace("?", ""), "cins": cins.replace("?", ""),
            "siparis_miktari": sip.replace("?", ""), "kg": kg.replace("?", ""),
            "fiyat": fiyat.replace("?", ""), "supheli": supheli, "not": notu,
        })
    return [s for s in sonuc if s["ebat"]]
