#!/usr/bin/env python3
"""Bu hafta + gelecek hafta için veriyi toplar, site/ klasörüne index.html ve Excel dosyalarını yazar."""
import datetime as dt
import html
import json
import sys
from pathlib import Path

from istanbul_hafta import TZ, hafta_araligi, topla, yaz

OUT = Path(__file__).resolve().parent / "site"
AYLAR = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos",
         "Eylül", "Ekim", "Kasım", "Aralık"]


def jsonable(rows):
    out = []
    for r in rows:
        r = dict(r)
        r["Tarih"] = r["Tarih"].isoformat()
        out.append(r)
    return out


def etiket(bas, bit):
    if bas.month == bit.month:
        return f"{bas.day}–{bit.day} {AYLAR[bit.month - 1]} {bit.year}"
    return f"{bas.day} {AYLAR[bas.month - 1]} – {bit.day} {AYLAR[bit.month - 1]} {bit.year}"


def main():
    log = lambda m: print(m, file=sys.stderr)
    OUT.mkdir(exist_ok=True)
    bugun = dt.datetime.now(TZ).date()
    haftalar = []
    for ad, ref in (("Bu hafta", bugun), ("Gelecek hafta", bugun + dt.timedelta(days=7))):
        bas, bit = hafta_araligi(ref)
        log(f"=== {ad}: {bas} – {bit}")
        maclar, etk, durum = topla(bas, bit, log)
        dosya = f"istanbul_program_{bas:%Y-%m-%d}.xlsx"
        yaz(OUT / dosya, maclar, etk, bas, bit, durum)
        haftalar.append({"ad": ad, "etiket": etiket(bas, bit), "bas": bas.isoformat(),
                         "excel": dosya, "maclar": jsonable(maclar), "etkinlikler": jsonable(etk),
                         "durum": [[str(x) for x in d] for d in durum]})
    veri = {"guncelleme": dt.datetime.now(TZ).strftime("%d.%m.%Y %H:%M"), "haftalar": haftalar}
    sablon = (Path(__file__).resolve().parent / "sablon.html").read_text("utf-8")
    js = json.dumps(veri, ensure_ascii=False).replace("</", "<\\/")
    (OUT / "index.html").write_text(sablon.replace("/*__VERI__*/null", js), "utf-8")
    log(f"✓ site/index.html yazıldı ({sum(len(h['maclar']) for h in haftalar)} maç)")


if __name__ == "__main__":
    main()
