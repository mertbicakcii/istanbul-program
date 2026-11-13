#!/usr/bin/env python3
"""
İstanbul Haftalık Maç & Etkinlik Programı
=========================================
Belirli bir hafta (Pazartesi–Pazar) içinde İstanbul'da oynanacak yerel ve
uluslararası maçları ve bilet satışındaki etkinlikleri resmi kaynaklardan çeker,
tek bir Excel dosyasına yazar.

Kaynaklar
  Futbol    : TFF (tff.org) lig fikstürleri + maç/stat detay sayfaları
              UEFA (match.uefa.com) – UCL/UEL/UECL, kadınlar, milli maçlar, gençler
  Basketbol : EuroLeague & EuroCup resmi API (api-live.euroleague.net)
  Voleybol  : TVF genel fikstür Excel'i (yurt içi ligler + CEV kupaları + milli)
  Etkinlik  : Biletix İstanbul etkinlik araması

Kullanım
  python istanbul_hafta.py                 # önümüzdeki hafta
  python istanbul_hafta.py --hafta 2026-11-23   # o tarihi içeren hafta
  python istanbul_hafta.py --cikti program.xlsx
"""
import argparse
import concurrent.futures as cf
import datetime as dt
import io
import json
import re
import sys
import time
from pathlib import Path
from zoneinfo import ZoneInfo

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

TZ = ZoneInfo("Europe/Istanbul")
UTC = dt.timezone.utc
UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/130.0 Safari/537.36")
BASE = Path(__file__).resolve().parent
CACHE_DIR = BASE / ".cache"
GUNLER = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]

S = requests.Session()
S.headers.update({"User-Agent": UA, "Accept-Language": "tr-TR,tr;q=0.9,en;q=0.8"})


def get(url, **kw):
    for i in range(3):
        try:
            r = S.get(url, timeout=40, **kw)
            r.raise_for_status()
            return r
        except requests.RequestException:
            if i == 2:
                raise
            time.sleep(1.5 * (i + 1))


def tr_upper(s):
    return (s or "").replace("i", "İ").replace("ı", "I").upper()


def is_istanbul(*texts):
    return any("İSTANBUL" in tr_upper(t) or "ISTANBUL" in tr_upper(t) for t in texts if t)


class Cache:
    """Basit JSON önbelleği (maç detayı / stat şehri gibi az değişen veriler)."""

    def __init__(self, name):
        CACHE_DIR.mkdir(exist_ok=True)
        self.path = CACHE_DIR / f"{name}.json"
        try:
            self.data = json.loads(self.path.read_text("utf-8"))
        except Exception:
            self.data = {}

    def save(self):
        self.path.write_text(json.dumps(self.data, ensure_ascii=False), "utf-8")


def row(brans, organizasyon, tarih, saat, mac, salon, kaynak, not_=""):
    return {"Tarih": tarih, "Gün": GUNLER[tarih.weekday()], "Saat": saat, "Branş": brans,
            "Organizasyon": organizasyon, "Maç": mac, "Salon / Stat": salon,
            "Not": not_, "Kaynak": kaynak}


# ---------------------------------------------------------------- UEFA ---
def uefa(bas, bit):
    out, offset = [], 0
    while True:
        r = get("https://match.uefa.com/v5/matches", params={
            "fromDate": bas.isoformat(), "toDate": bit.isoformat(),
            "limit": 100, "offset": offset, "order": "ASC"}).json()
        for m in r:
            st = m.get("stadium") or {}
            city = ((st.get("city") or {}).get("translations") or {}).get("name", {}).get("EN", "")
            if not is_istanbul(city):
                continue
            k = dt.datetime.fromisoformat(m["kickOffTime"]["dateTime"].replace("Z", "+00:00")).astimezone(TZ)
            stad = ((st.get("translations") or {}).get("name") or {}).get("EN") or st.get("name", "")
            home = (m.get("homeTeam") or {}).get("internationalName", "?")
            away = (m.get("awayTeam") or {}).get("internationalName", "?")
            out.append(row("Futbol", m["competition"]["metaData"]["name"], k.date(),
                           k.strftime("%H:%M"), f"{home} - {away}", stad, "UEFA"))
        if len(r) < 100:
            break
        offset += 100
    return out


# ----------------------------------------------------------- EuroLeague ---
def euroleague(bas, bit):
    out = []
    sezon = bas.year if bas.month >= 7 else bas.year - 1
    for code, ad in (("E", "EuroLeague"), ("U", "EuroCup")):
        data = get(f"https://api-live.euroleague.net/v2/competitions/{code}/seasons/{code}{sezon}/games").json()["data"]
        for g in data:
            d = dt.datetime.fromisoformat(g["utcDate"].replace("Z", "+00:00")).astimezone(TZ)
            if not (bas <= d.date() <= bit):
                continue
            v = g.get("venue") or {}
            adres = v.get("address") or ""
            ev_ist = is_istanbul(g["local"]["club"]["name"]) and not g.get("isNeutralVenue")
            # Adreste İstanbul geçiyorsa ya da ev sahibi İstanbul takımı kendi salonunda oynuyorsa
            # (ör. Turkcell salonunun adresinde şehir yazmıyor)
            if not (is_istanbul(adres) or (ev_ist and (not adres or "TURK" in tr_upper(adres)))):
                continue
            saat = d.strftime("%H:%M") if g.get("confirmedHour", True) else "Belirsiz"
            out.append(row("Basketbol", f"{ad} – {g.get('roundName', '')}", d.date(), saat,
                           f"{g['local']['club']['name']} - {g['road']['club']['name']}",
                           v.get("name", ""), "EuroLeague API"))
    return out


# ------------------------------------------------------------------ TVF ---
TVF_LIG = {
    "VSL": "Sultanlar Ligi", "SGEL": "Efeler Ligi", "ACH1": "1. Lig", "UP2": "2. Lig",
    "ASKV": "Kupa Voley", "ASŞK": "Şampiyonlar Kupası", "CL": "CEV Şampiyonlar Ligi",
    "CH": "CEV Kupası", "CC": "CEV Challenge Kupası", "CWC": "Kulüpler Dünya Şampiyonası",
    "EC": "Avrupa Şampiyonası", "WC": "Dünya Şampiyonası", "BC": "Balkan Kupası",
    "MG": "Akdeniz Oyunları", "BVA CUP": "BVA Kupası",
}


def tvf(bas, bit):
    sezon = bas.year if bas.month >= 8 else bas.year - 1
    url = (f"https://tvf.org.tr/_dosyalar/Lig_Sezon_Arsivi/Genel_Fiksturler/"
           f"{sezon}-{sezon + 1}_genel_fikstur.xlsx")
    ws = load_workbook(io.BytesIO(get(url).content), read_only=True, data_only=True).active
    rows = ws.iter_rows(values_only=True)
    hdr = [str(h or "").strip() for h in next(rows)]
    ix = {h: i for i, h in enumerate(hdr)}

    def c(r, name):
        i = ix.get(name)
        return r[i] if i is not None and i < len(r) else None

    out = []
    for r in rows:
        t = c(r, "TARİH")
        if not isinstance(t, dt.datetime) or not (bas <= t.date() <= bit):
            continue
        if not is_istanbul(c(r, "İL (ÜLKE*)")):
            continue
        saat = c(r, "SAAT")
        saat = saat.strftime("%H:%M") if isinstance(saat, dt.time) else (str(saat) if saat else "Belirsiz")
        lig = str(c(r, "LİG") or "").strip()
        cins = {"E": "Erkekler", "K": "Kadınlar"}.get(str(c(r, "E-K") or "").strip(), "")
        org = " ".join(x for x in (TVF_LIG.get(lig, lig), f"({cins})" if cins else "") if x)
        tur = str(c(r, "TÜR") or "").strip()
        if tur and tur.upper() not in ("LE",):
            org += " – " + " ".join(w[:1] + w[1:].replace("I", "ı").replace("İ", "i").lower()
                                    for w in tur.upper().split())
        not_ = str(c(r, "TV") or "").strip()
        out.append(row("Voleybol", org, t.date(), saat,
                       f"{str(c(r, 'A TAKIMI') or '').strip()} - {str(c(r, 'B TAKIMI') or '').strip()}",
                       str(c(r, "SALON (ŞEHİR*)") or "Belirsiz").strip(), "TVF",
                       f"Yayın: {not_}" if not_ else ""))
    return out


# ------------------------------------------------------------------ TFF ---
TFF_LIGLER = {
    198: "Trendyol Süper Lig", 142: "Trendyol 1. Lig", 976: "Nesine 2. Lig",
    971: "Nesine 3. Lig", 1000: "Kadın Futbol Süper Ligi", 1602: "TFF Kadınlar 1. Ligi",
    1001: "TFF Kadınlar 2. Ligi", 1298: "TFF Kadınlar 3. Ligi", 1750: "U19 PAF Ligi",
    335: "U21, A2 ve PAF Ligi", 900: "Futsal Ligi",
}
TFF = "https://www.tff.org/"


def tff_html(url):
    r = get(url)
    return r.content.decode("cp1254", errors="ignore")


def tff_bloklar(html):
    """Sayfayı 'N.Hafta' bloklarına böler; her blok için [(macId, oynandı_mı)] döner."""
    parts = re.split(r">\s*\d+\s*\.\s*Hafta\s*<", html, flags=re.I)
    bloklar = []
    for p in parts[1:]:
        maclar = [(mid, bool(re.search(r"\d", txt)))
                  for mid, txt in re.findall(r"macId=(\d+)\"[^>]*>([^<]*)</a>", p, flags=re.I)]
        if maclar:
            bloklar.append(maclar)
    return bloklar


def tff_mac(mid, cache):
    if mid in cache.data and cache.data[mid].get("ts", 0) > time.time() - 6 * 3600:
        return cache.data[mid]
    h = tff_html(f"{TFF}Default.aspx?pageID=29&macId={mid}")

    def f(key):
        m = re.search(rf'id="[^"]*{key}"[^>]*>([^<]*)<', h)
        return re.sub(r"\s+", " ", m.group(1)).strip() if m else ""

    sid = re.search(r"stadId=(\d+)", h)
    d = {"ev": f("lnkTakim1"), "dep": f("lnkTakim2"), "org": f("lblOrganizasyonAdi"),
         "stad": f("lnkStad"), "tarih": f("lblTarih"), "stadId": sid.group(1) if sid else "",
         "ts": time.time()}
    cache.data[mid] = d
    return d


def tff_sehir(sid, cache):
    if not sid:
        return ""
    if sid not in cache.data:
        h = tff_html(f"{TFF}Default.aspx?pageId=394&stadId={sid}")
        t = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", h))
        m = re.search(r"Şehir\s*:\s*([^\s:]+(?:\s[^\s:]+)?)\s+(?:Açıklama|Oyun|Adres|$)", t)
        cache.data[sid] = m.group(1).strip() if m else ""
    return cache.data[sid]


def tff_tarih(s):
    m = re.match(r"(\d{2})\.(\d{2})\.(\d{4})(?:\s*-\s*(\d{2}:\d{2}))?", s or "")
    if not m:
        return None, ""
    return dt.date(int(m[3]), int(m[2]), int(m[1])), (m[4] or "Belirsiz")


def tff(bas, bit, log):
    mc, sc = Cache("tff_mac"), Cache("tff_stad")
    out, gorulen = [], set()
    with cf.ThreadPoolExecutor(6) as ex:
        for pid, ad in TFF_LIGLER.items():
            try:
                h = tff_html(f"{TFF}default.aspx?pageID={pid}")
                sayfalar = [h] + [tff_html(f"{TFF}default.aspx?pageID={pid}&grupID={g}")
                                  for g in sorted(set(re.findall(r"grupID=(\d+)", h, flags=re.I)))]
                for sh in sayfalar:
                    for blok in tff_bloklar(sh):
                        ids = [m for m, oyn in blok if not oyn and m not in gorulen]
                        if not ids:
                            continue
                        gorulen.update(ids)
                        detay = list(ex.map(lambda m: tff_mac(m, mc), ids))
                        tarihler = [tff_tarih(d["tarih"])[0] for d in detay]
                        for d in detay:
                            gun, saat = tff_tarih(d["tarih"])
                            if gun and bas <= gun <= bit:
                                sehir = tff_sehir(d["stadId"], sc)
                                if is_istanbul(sehir):
                                    out.append(row("Futbol", d["org"] or ad, gun, saat,
                                                   f"{d['ev']} - {d['dep']}", d["stad"], "TFF"))
                        if any(t and t > bit for t in tarihler):
                            break  # sonraki haftalar hedef haftanın dışında
            except Exception as e:  # tek lig hatası tüm çıktıyı durdurmasın
                log(f"  ! TFF {ad}: {e}")
    mc.save()
    sc.save()
    return out


# -------------------------------------------------------------- Biletix ---
BX_KAT = {"MUSIC": "Müzik", "ART": "Sahne / Sanat", "FAMILY": "Aile", "SPORT": "Spor",
          "OTHER": "Diğer / Festival / Deneyim"}


def biletix(bas, bit):
    a = dt.datetime.combine(bas, dt.time(0), TZ).astimezone(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")
    b = dt.datetime.combine(bit, dt.time(23, 59, 59), TZ).astimezone(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")
    r = get("https://www.biletix.com/solr/tr/select/", params=[
        ("q", "*:*"), ("wt", "json"), ("rows", "2000"), ("start", "0"),
        ("fq", 'city:"İstanbul"'), ("fq", "type:event"), ("fq", f"start:[{a} TO {b}]")]).json()
    out = []
    for d in r["response"]["docs"]:
        s = dt.datetime.fromisoformat(d["start"].replace("Z", "+00:00")).astimezone(TZ)
        e = dt.datetime.fromisoformat(d.get("end", d["start"]).replace("Z", "+00:00")).astimezone(TZ)
        org = ", ".join(d.get("group") or [])
        not_ = f"{e.date().strftime('%d.%m.%Y')} tarihine kadar" if e.date() > s.date() else ""
        out.append({"Tarih": s.date(), "Gün": GUNLER[s.weekday()], "Saat": s.strftime("%H:%M"),
                    "Kategori": BX_KAT.get(d.get("category"), d.get("category", "")),
                    "Etkinlik": d.get("sname") or (d.get("name") or [""])[0],
                    "Organizatör / Seri": org, "Mekan": d.get("svenue", ""), "Not": not_,
                    "Link": f"https://www.biletix.com/etkinlik/{d['id']}/TURKIYE/tr",
                    "Kaynak": "Biletix"})
    return out


# ---------------------------------------------------------------- Excel ---
def yaz(path, maclar, etkinlikler, bas, bit, durum):
    wb = Workbook()
    baslik = Font(bold=True, color="FFFFFF")
    dolgu = PatternFill("solid", fgColor="1F3A5F")

    def sayfa(ws, kayitlar, kolonlar, genislik):
        ws.append(kolonlar)
        for c in ws[1]:
            c.font, c.fill = baslik, dolgu
            c.alignment = Alignment(vertical="center")
        for k in kayitlar:
            ws.append([k.get(c, "") for c in kolonlar])
        for i, w in enumerate(genislik, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        for c in ws["A"][1:]:
            c.number_format = "DD.MM.YYYY"
        ws.freeze_panes = "A2"
        if kayitlar:
            ws.auto_filter.ref = ws.dimensions

    ws = wb.active
    ws.title = "Maçlar"
    sayfa(ws, maclar, ["Tarih", "Gün", "Saat", "Branş", "Organizasyon", "Maç", "Salon / Stat", "Not", "Kaynak"],
          [12, 11, 9, 11, 42, 52, 44, 22, 14])
    sayfa(wb.create_sheet("Etkinlikler"), etkinlikler,
          ["Tarih", "Gün", "Saat", "Kategori", "Etkinlik", "Organizatör / Seri", "Mekan", "Not", "Link", "Kaynak"],
          [12, 11, 9, 22, 46, 34, 36, 24, 40, 10])
    ws = wb.create_sheet("Kaynak Durumu")
    ws.append([f"İstanbul programı: {bas:%d.%m.%Y} – {bit:%d.%m.%Y}"])
    ws.append([f"Oluşturulma: {dt.datetime.now(TZ):%d.%m.%Y %H:%M}"])
    ws.append([])
    ws.append(["Kaynak", "Durum", "Kayıt"])
    for c in ws[4]:
        c.font, c.fill = baslik, dolgu
    for d in durum:
        ws.append(d)
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 70
    wb.save(path)


def hafta_araligi(ref):
    bas = ref - dt.timedelta(days=ref.weekday())
    return bas, bas + dt.timedelta(days=6)


def topla(bas, bit, log=lambda m: None, etkinlik=True):
    """Hafta için (maçlar, etkinlikler, kaynak_durumu) döner."""
    kaynaklar = [("UEFA (match.uefa.com)", lambda: uefa(bas, bit)),
                 ("EuroLeague / EuroCup API", lambda: euroleague(bas, bit)),
                 ("TVF genel fikstür", lambda: tvf(bas, bit)),
                 ("TFF lig fikstürleri", lambda: tff(bas, bit, log))]
    maclar, durum = [], []
    for ad, fn in kaynaklar:
        log(f"→ {ad} ...")
        try:
            r = fn()
            maclar += r
            durum.append([ad, "OK", len(r)])
            log(f"  {len(r)} İstanbul maçı")
        except Exception as e:
            durum.append([ad, f"HATA: {e}", 0])
            log(f"  ! {e}")

    etkinlikler = []
    if etkinlik:
        log("→ Biletix ...")
        try:
            etkinlikler = biletix(bas, bit)
            durum.append(["Biletix etkinlikleri", "OK", len(etkinlikler)])
            log(f"  {len(etkinlikler)} etkinlik")
        except Exception as e:
            durum.append(["Biletix etkinlikleri", f"HATA: {e}", 0])
            log(f"  ! {e}")

    durum += [
        ["TBF (BSL, kadın basketbol ligleri)", "Otomatik erişim engelli (bot koruması) – tbf.org.tr'den elle kontrol edin", ""],
        ["FIBA (BCL, EuroLeague Women, EuroCup)", "Herkese açık API yok – fiba.basketball'dan elle kontrol edin", ""],
        ["Hentbol / buz hokeyi / diğer", "Henüz eklenmedi", ""],
    ]

    # aynı maç iki kaynaktan gelirse tekilleştir
    tek, seen = [], set()
    for m in sorted(maclar, key=lambda x: (x["Tarih"], x["Saat"], x["Branş"])):
        k = (m["Tarih"], m["Saat"], tr_upper(m["Maç"]))
        if k not in seen:
            seen.add(k)
            tek.append(m)
    etkinlikler.sort(key=lambda x: (x["Tarih"], x["Saat"], x["Etkinlik"]))
    return tek, etkinlikler, durum


def main():
    ap = argparse.ArgumentParser(description="İstanbul haftalık maç ve etkinlik programı")
    ap.add_argument("--hafta", help="Hafta içindeki herhangi bir tarih (YYYY-AA-GG). Varsayılan: önümüzdeki hafta")
    ap.add_argument("--cikti", help="Excel dosya adı")
    ap.add_argument("--etkinlik-yok", action="store_true", help="Biletix etkinliklerini atla")
    a = ap.parse_args()

    bugun = dt.datetime.now(TZ).date()
    ref = dt.date.fromisoformat(a.hafta) if a.hafta else bugun + dt.timedelta(days=7 - bugun.weekday())
    bas, bit = hafta_araligi(ref)
    cikti = Path(a.cikti or BASE / f"istanbul_program_{bas:%Y-%m-%d}.xlsx")
    log = lambda m: print(m, file=sys.stderr)
    log(f"Hafta: {bas:%d.%m.%Y} – {bit:%d.%m.%Y}")
    tek, etkinlikler, durum = topla(bas, bit, log, not a.etkinlik_yok)
    yaz(cikti, tek, etkinlikler, bas, bit, durum)
    log(f"\n✓ {len(tek)} maç, {len(etkinlikler)} etkinlik → {cikti}")


if __name__ == "__main__":
    main()
