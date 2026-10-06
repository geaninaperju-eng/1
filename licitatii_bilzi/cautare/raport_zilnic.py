#!/usr/bin/env python3
"""Raport zilnic BILZI: toate anunțurile de lucrări deschise de pe licitatiipublice.ro, cu verdict.

Rulare (din licitatii_bilzi/):  python3 cautare/raport_zilnic.py  [AAAA-LL-ZZ]
Ieșire: rapoarte/licitatii_<data>.csv (verdict, termen, județ, km, valoare, obiect, motiv, link).
Cere variabilele LICITATIIPUBLICE_USER / LICITATIIPUBLICE_PASS. Folosește doar paginile de căutare
(paginile de detaliu au limită zilnică pe cont – deschide-le doar pentru candidații potriviți).
"""
import csv, datetime, json, re, sys, time
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "clienti")); sys.path.insert(0, str(ROOT / "cautare"))
import lp_client
lp_client.PAUSE = 2.5
from lp_client import LP
import parse

AZI = datetime.date.fromisoformat(sys.argv[1]) if len(sys.argv) > 1 else datetime.date.today()
KW = ['acoperis', 'acoperisuri', 'invelitoare', 'sarpanta', 'jgheaburi', 'burlane', 'tigla', 'tinichigerie',
      'hidroizolatie', 'terasa', 'reparatii acoperis', 'pluvial', 'reabilitare cladire', 'reparatii curente',
      'lucrari de reparatii', 'tabla']
LUNI = {'Ian': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'Mai': 5, 'Iun': 6, 'Iul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10,
        'Noi': 11, 'Dec': 12}
KM = {'Buzau': 0, 'Braila': 100, 'Ialomita': 90, 'Prahova': 110, 'Vrancea': 90, 'Galati': 130, 'Bucuresti': 115,
      'Ilfov': 120, 'Calarasi': 160, 'Dambovita': 180, 'Giurgiu': 180, 'Covasna': 130, 'Brasov': 150, 'Bacau': 180,
      'Tulcea': 200, 'Constanta': 230, 'Vaslui': 250, 'Teleorman': 220, 'Arges': 240, 'Neamt': 250, 'Iasi': 290,
      'Harghita': 260}
ROOF = re.compile(r'acoperi|invelit|sarpant|jgheab|burlan|tigl|tinichig|streasin|hidroizola|terasa', re.I)
NOT = re.compile(r'drum|podet|canaliz|autostr|apa potabila|retea|retele|covor|parcare|strazi|bituminoase|trotuar|'
                 r'iluminat|decolmatare|remorcher|chiller|cablu|pod din', re.I)
WORK = re.compile(r'lucrari|reparati|reabilit|construi|moderniz|consolid|executie|eficien|renovare|amenajare|'
                  r'refacere|refunctionalizare|extindere', re.I)
SUPPLY = re.compile(r'furnizare|materiale|achizitie (de )?(echip|mobilier|produse)|servicii|medicament|dotari|'
                    r'software|autovehicul|carburant|alimente', re.I)
EMOJI = {'2': '🟢 POTRIVITĂ', '3': '🟡 DE VERIFICAT / CONDIȚIONATĂ', '4': '🔵 DOAR SUBCONTRACTARE', '5': '⚪ NU'}


def data(s):
    try:
        z, l, a = s.split()[:3]
        return datetime.date(int(a), LUNI[l], int(z))
    except Exception:
        return None


def scaneaza():
    lp, seen, tmp = LP(), {}, ROOT / "rapoarte" / ".r.html"
    for kw in KW:
        h = lp.search(kw, 'achizitii', 0, months=1)
        if 'Am detectat o problema' in h:
            time.sleep(660)
            h = lp.search(kw, 'achizitii', 0, months=1)
        for pg in range(6):
            if pg:
                h = lp.page('achizitii', pg)
            tmp.write_text(h)
            its = parse.items(str(tmp))
            if not its:
                break
            for i in its:
                seen[i['link']] = i
    tmp.unlink(missing_ok=True)
    return list(seen.values())


def clasifica(items):
    rows = []
    for i in items:
        t = data(i['termen'] or '')
        if not t or t < AZI:
            continue
        title, loc = (i['title'] or '').strip(), i['loc'] or ''
        km = min([KM[c.strip()] for c in loc.split(',') if c.strip() in KM], default=None)
        try:
            v = float((i['val'] or '').replace('RON', '').replace('.', '').replace(',', '.').strip())
        except ValueError:
            v = None
        roof = bool(ROOF.search(title)) and not NOT.search(title)
        work = bool(WORK.search(title)) and not SUPPLY.search(title) and not NOT.search(title)
        if not (roof or work):
            continue
        if km is None:
            ver, mot = '5', 'În afara razei de 300 km de Buzău'
        elif roof and (v is None or v <= 150000):
            ver, mot = '2', 'Acoperiș/terasă în zonă, sub pragul de experiență (122.455 lei) – verifică caietul'
        elif roof and v <= 600000:
            ver, mot = '3', 'Acoperiș în zonă – depinde de experiența cerută în fișa de date'
        elif roof:
            ver, mot = '4', 'Acoperiș în zonă, valoare mare – doar subcontractor/asociat'
        elif v is not None and v <= 450200:
            ver, mot = '3', 'Lucrare de construcții în zonă (poate include acoperiș) – verifică lista de cantități'
        else:
            ver, mot = '4', 'Lucrare mare de construcții – subcontractare învelitoare/tinichigerie'
        ch = i['link'].split('ch=')[-1]
        rows.append([ver, t.isoformat(), loc, '' if km is None else km, '' if v is None else f"{v:.0f}", title[:110],
                     mot, 'https://www.licitatiipublice.ro/licitatiipublice/module/achizitii/vizualizare_anunt.jsp?ch='
                     + ch])
    rows.sort(key=lambda r: (r[0], r[1], r[3] if r[3] != '' else 999))
    for r in rows:
        r[0] = EMOJI[r[0]]
    return rows


if __name__ == "__main__":
    rows = clasifica(scaneaza())
    out = ROOT / "rapoarte" / f"licitatii_{AZI.isoformat()}.csv"
    with open(out, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(['Verdict', 'Termen', 'Județ', 'Km de la Buzău (aprox.)', 'Valoare estimată (lei fără TVA)',
                    'Obiectul achiziției', 'De ce', 'Link'])
        w.writerows(rows)
    from collections import Counter
    print(out, len(rows), dict(Counter(r[0] for r in rows)))
