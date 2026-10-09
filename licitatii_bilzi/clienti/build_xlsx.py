"""Faza 3: construiește clienti_potentiali.xlsx (foile Clienti, Achizitii, Concurenti) din date/lista.json + date/detalii.json."""
import json, os, re, datetime as dt
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from detalii import ROOF

HERE = os.path.dirname(os.path.abspath(__file__))
D = os.path.join(HERE, "date")
BASE = "https://www.licitatiipublice.ro/licitatiipublice/module"
TODAY = dt.date.today()

# distanță rutieră aproximativă Buzău -> reședința de județ (km)
DIST = {"Buzau": 20, "Vrancea": 70, "Braila": 100, "Prahova": 100, "Ialomita": 100, "Bucuresti": 120,
        "Ilfov": 115, "Galati": 125, "Covasna": 125, "Brasov": 145, "Dambovita": 165, "Calarasi": 165,
        "Bacau": 170, "Giurgiu": 185, "Harghita": 190, "Tulcea": 200, "Teleorman": 210, "Constanta": 220,
        "Arges": 225, "Neamt": 230, "Vaslui": 240, "Iasi": 300}
# localități importante din județul Buzău / apropiate (km)
DIST_LOC = {"buzau": 0, "ramnicu sarat": 35, "nehoiu": 60, "pogoanele": 40, "patarlagele": 45,
            "focsani": 70, "ploiesti": 100, "braila": 100, "slobozia": 100, "urziceni": 75, "mizil": 35}

LUNI = dict(ian=1, feb=2, mar=3, apr=4, mai=5, iun=6, iul=7, aug=8, sep=9, oct=10, noi=11, nov=11, dec=12)

def norm(s):
    s = (s or "").lower()
    for a, b in zip("ăâîșşțţ", "aaisstt"):
        s = s.replace(a, b)
    return s

def pdate(s):
    if not s:
        return None
    s = s.strip()
    m = re.match(r"(\d{1,2})-(\d{1,2})-(\d{4})", s)
    if m:
        return dt.date(int(m[3]), int(m[2]), int(m[1]))
    m = re.match(r"(\d{1,2}) (\w{3})\w* (\d{4})", s)
    if m and m[2].lower() in LUNI:
        return dt.date(int(m[3]), LUNI[m[2].lower()], int(m[1]))
    return None

def pval(s):
    if not s:
        return None
    s = s.replace("\xa0", " ")
    m = re.search(r"([\d.,]+)\s*(RON|EUR|LEI)?", s, re.I)
    if not m:
        return None
    n = m[1]
    if re.fullmatch(r"\d{1,3}(\.\d{3})+(,\d+)?", n):
        n = n.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"\d{1,3}(,\d{3})+(\.\d+)?", n):
        n = n.replace(",", "")
    else:
        n = n.replace(",", ".")
    try:
        v = float(n)
    except ValueError:
        return None
    if (m[2] or "").upper() == "EUR":
        v *= 5.0
    return round(v, 2)

TIPURI = [("MApN / unitate militară", r"aparari|u\.?m\.? ?\d|unitatea militara|cazarma|statul major|brigada|baza aeriana"),
          ("Poliție / Jandarmerie / IGSU / MAI", r"politi|jandarm|situatii de urgenta|isu |igsu|afacerilor interne|frontier|imigrari|penitenciar"),
          ("Romsilva / silvic", r"romsilva|padurilor|silvic|\bocolul"),
          ("Universitate / cercetare", r"universitat|academi|institutul|cercetare"),
          ("Spital / sănătate", r"spital|sanatate|ambulant|dsp |medic|cantacuzino"),
          ("Școală / grădiniță / liceu", r"scoal|liceu|colegi|gradinit|gimnazial|creș|cresa|palatul copiilor|seminar"),
          ("Cultură / patrimoniu / culte", r"muzeu|biblioteca|teatr|cultur|patrimoniu|biseric|parohi|episcopi|arhiepiscop|manastir|filarmon"),
          ("Asistență socială", r"dgaspc|asistenta sociala|protectia copilului|camin|centrul de ingrijire|batrani"),
          ("Primărie / UAT", r"primari|^comuna|^orasul|^municipiul|^judetul|consiliul (local|judetean)|^uat|sector \d"),
          ("Companie de stat / regie", r"\bs\.a\.?|\bsa\b|\bregia\b|\br\.a\.?|compania|\bcfr\b|hidroelectrica|electrica|transelectrica|\bposta\b|apele romane|aeroport|\bport\b|termoficare|nuclearelectrica|romgaz|conpet|cnair|administratia nationala"),
          ("Finanțe / administrație centrală", r"finantelor|ministerul|agentia|autoritatea|directia|institutia prefectului|casa (de|judeteana)|anaf|oficiul")]

def tip(name):
    n = norm(name)
    for t, rx in TIPURI:
        if re.search(rx, n):
            return t
    if re.search(r"\bs\.?r\.?l\.?|\bsrl\b", n):
        return "Firmă privată (proiect cu finanțare)"
    return "Altă entitate publică"

def judet_of(loc):
    for j in DIST:
        if loc and j.lower() in norm(loc):
            return j
    return None

def dist(judet, localitate):
    l = re.sub(r"^(municipiul|orasul|oras|comuna|sat)\s+", "", norm(localitate).strip()).replace("-", " ")
    if l in DIST_LOC:
        return DIST_LOC[l]
    return DIST.get(judet)

def main():
    lista = json.load(open(os.path.join(D, "lista.json")))["items"]
    det = json.load(open(os.path.join(D, "detalii.json")))
    planuri_l = []
    for k, it in lista.items():
        if it["modul"] != "planuriachizitii":
            continue
        jl = [j for j in DIST if j.lower() in norm(it.get("localizare"))]
        if not jl or not (set(it["cuvinte"]) - {"tabla", "materiale acoperis"} or ROOF.search(it["titlu"] or "")):
            continue
        planuri_l.append(dict(data=pdate(it.get("data_publicare")), data_est=pdate(it.get("data_estimata")),
                              titlu=it["titlu"], judet=jl[0], valoare=pval(it.get("valoare")),
                              cuvinte=", ".join(it["cuvinte"]),
                              link=f"{BASE}/planuriachizitii/vizualizare_anunt.jsp?ch={it['ch']}"))
    plan_jud = defaultdict(list)
    for p_ in planuri_l:
        plan_jud[p_["judet"]].append(p_)
    rows = []
    for k, it in lista.items():
        if it["modul"] == "planuriachizitii":
            continue
        d = det.get(k)
        if not d:
            # detaliu nedescărcat (limită cont): folosim doar datele din listă
            if not ROOF.search(it["titlu"] or ""):
                continue
            d = {"titlu": it["titlu"], "entitati": [], "Tip procedura": "(detaliu nedescărcat)"}
        titlu = d.get("titlu") or it["titlu"]
        text = " ".join(filter(None, [titlu, it["titlu"], d.get("Descriere"), d.get("Obiect"), d.get("cpv")]))
        if re.search(r"tinichigerie auto|mecanica si tinichigerie|caroseri|auto(turism|vehicul)", norm(text)):
            continue
        if not ROOF.search(text) and not re.search(r"4526\d|44112[1-5]|4419\d", d.get("cpv") or ""):
            continue
        jud_loc = [j for j in DIST if j.lower() in norm(it.get("localizare") or d.get("Localizare"))]
        org = next((e for e in d["entitati"] if norm(e["sectiune"]).startswith(("organizator", "autoritate"))),
                   d["entitati"][0] if d["entitati"] else {})
        win = next((e for e in d["entitati"] if e is not org), None)
        judet = org.get("Judet") or (jud_loc[0] if jud_loc else None)
        if not jud_loc and judet_of(judet) is None:
            continue
        val = pval(d.get("Valoare atribuita") or d.get("Pret maximal") or d.get("Valoare estimata")
                   or d.get("Val. estimata") or d.get("Valoare") or it.get("valoare"))
        data = (pdate(it.get("data_atribuire")) if it["modul"] == "atribuiri" else None) or \
               pdate(it.get("data_publicare")) or pdate(d.get("Data publicare"))
        rows.append(dict(key=k, data=data, modul=it["modul"], nr=d.get("Numar anunt"), obiect=titlu,
                         tip_proc=d.get("Tip procedura"), cpv=d.get("cpv"),
                         org=org.get("denumire") or it.get("organizator"), cui=org.get("Cod fiscal"),
                         judet=judet_of(judet) or (jud_loc[0] if jud_loc else judet),
                         localitate=org.get("Localitate"), org_ent=org, valoare=val,
                         termen=it.get("termen") or d.get("Termen limita") or it.get("data_estimata"),
                         win=win, link=f"{BASE}/{it['modul']}/vizualizare_anunt.jsp?ch={it['ch']}"))
    # de-dublare (același anunț apare uneori de 2 ori - loturi): modul+nr+obiect+organizator
    seen, uniq = set(), []
    for r in sorted(rows, key=lambda r: (r["data"] or dt.date(1900, 1, 1)), reverse=True):
        sig = (r["modul"], r["nr"] or r["key"], norm(r["obiect"])[:80], r["cui"] or r["org"],
               (r["win"] or {}).get("Cod fiscal"))
        if sig in seen:
            continue
        seen.add(sig)
        uniq.append(r)
    rows = uniq

    # ---- Clienti (rândurile fără CUI se leagă după denumire de cele cu CUI)
    name2cui = {norm(r["org"]): r["cui"] for r in rows if r["cui"]}
    cl = defaultdict(list)
    for r in rows:
        cl[r["cui"] or name2cui.get(norm(r["org"])) or norm(r["org"])].append(r)
    clienti = []
    for key, rs in cl.items():
        o = next((r["org_ent"] for r in rs if r["org_ent"]), {})
        nume = o.get("denumire") or rs[0]["org"]
        jud = next((r["judet"] for r in rs if r["judet"]), None)
        loc = o.get("Localitate")
        achiz = [r for r in rs if r["modul"] != "planuriachizitii"]
        planuri = [r for r in rs if r["modul"] == "planuriachizitii"]
        vals = [r["valoare"] for r in rs if r["valoare"]]
        last = max(achiz or rs, key=lambda r: r["data"] or dt.date(1900, 1, 1))
        km = dist(jud, loc)
        # scor prioritate
        s = 0.0
        s += 2 if km is not None and km <= 60 else 1.5 if km is not None and km <= 130 else 1 if km is not None and km <= 200 else 0.5
        n = len(rs)
        s += 1.5 if n >= 4 else 1 if n >= 2 else 0.5
        s += 1 if any(20000 <= v <= 150000 for v in vals) else 0.5 if any(150000 < v <= 500000 or 5000 <= v < 20000 for v in vals) else 0
        s += 0.25 if plan_jud.get(jud) else 0
        s += 0.5 if any(r["data"] and (TODAY - r["data"]).days <= 365 for r in rs) else 0
        scor = 5 if s >= 4.5 else 4 if s >= 3.5 else 3 if s >= 2.75 else 2 if s >= 2 else 1
        clienti.append(dict(
            denumire=nume, cui=o.get("Cod fiscal") or "", tip=tip(nume), judet=jud, localitate=loc or "",
            adresa=o.get("Adresa", ""), telefon=o.get("Telefon", ""), email=o.get("Email", ""), web=o.get("Web", ""),
            nr=len(achiz), nr_plan=len(plan_jud.get(jud, [])), total=round(sum(vals), 2) if vals else None,
            ult_data=last["data"], ult_obiect=last["obiect"], ult_val=last["valoare"], ult_modul=last["modul"],
            plan=f"nedeterminat (modul plătit); în județ: {len(plan_jud.get(jud, []))} poziții plan cu acoperiș",
            plan_obiect="; ".join(sorted({p_["titlu"] for p_ in plan_jud.get(jud, [])}))[:300],
            km=km, scor=scor, scor_brut=round(s, 2)))
    clienti.sort(key=lambda c: (-c["scor"], -c["scor_brut"], c["km"] or 999, -(c["total"] or 0)))

    # ---- Concurenti
    co = defaultdict(list)
    for r in rows:
        if r["win"] and r["modul"] in ("cumpararidirecte", "atribuiri"):
            co[r["win"].get("Cod fiscal") or norm(r["win"]["denumire"])].append(r)
    concurenti = []
    for k, rs in co.items():
        w = rs[0]["win"]
        vals = [r["valoare"] for r in rs if r["valoare"]]
        concurenti.append(dict(denumire=w["denumire"], cui=w.get("Cod fiscal", ""), judet=w.get("Judet", ""),
                               localitate=w.get("Localitate", ""), telefon=w.get("Telefon", ""), email=w.get("Email", ""),
                               web=w.get("Web", ""), nr=len(rs), total=round(sum(vals), 2) if vals else None,
                               judete_lucru=", ".join(sorted({r["judet"] for r in rs if r["judet"]})),
                               clienti=", ".join(sorted({r["org"] for r in rs}))[:400]))
    concurenti.sort(key=lambda c: (-c["nr"], -(c["total"] or 0)))

    wb = Workbook()
    def sheet(ws, headers, data, widths):
        ws.append(headers)
        for row in data:
            ws.append(["" if x == "null" else x for x in row])
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1F4E78")
            c.alignment = Alignment(wrap_text=True, vertical="center")
        ws.freeze_panes = "B2"
        ws.auto_filter.ref = ws.dimensions
    ws = wb.active
    ws.title = "Clienti"
    sheet(ws, ["Scor prioritate (1-5)", "Denumire", "CUI", "Tip entitate", "Județ", "Localitate", "Adresă", "Telefon",
               "E-mail", "Web", "Nr. achiziții acoperiș găsite", "Nr. poziții plan cu acoperiș în județ",
               "Valoare totală estimată (lei)", "Ultima achiziție – data", "Ultima achiziție – obiect",
               "Ultima achiziție – valoare (lei)", "Ultima achiziție – modul", "Are plan achiziții cu acoperiș",
               "Planuri în județ – obiect (trunchiat de site)", "Distanța aprox. față de Buzău (km)"],
          [[c["scor"], c["denumire"], c["cui"], c["tip"], c["judet"], c["localitate"], c["adresa"], c["telefon"],
            c["email"], c["web"], c["nr"], c["nr_plan"], c["total"], c["ult_data"], c["ult_obiect"], c["ult_val"],
            c["ult_modul"], c["plan"], c["plan_obiect"], c["km"]] for c in clienti],
          [10, 45, 12, 26, 12, 16, 35, 16, 28, 22, 11, 11, 15, 12, 60, 14, 16, 10, 50, 12])
    ws2 = wb.create_sheet("Achizitii")
    sheet(ws2, ["Data", "Modul", "Nr. anunț", "Tip procedură", "Obiect", "Organizator", "CUI organizator", "Județ",
                "Valoare (lei)", "Termen / data estimată", "Câștigător / ofertant", "CUI câștigător", "CPV", "Link"],
          [[r["data"], r["modul"], r["nr"], r["tip_proc"], r["obiect"], r["org"], r["cui"], r["judet"], r["valoare"],
            r["termen"], (r["win"] or {}).get("denumire"), (r["win"] or {}).get("Cod fiscal"), r["cpv"], r["link"]]
           for r in rows],
          [11, 16, 14, 18, 60, 40, 12, 12, 14, 14, 35, 12, 40, 30])
    ws4 = wb.create_sheet("Planuri_judete")
    sheet(ws4, ["Județ", "Distanța aprox. față de Buzău (km)", "Nr. poziții plan cu acoperiș/învelitori",
                "Valoare totală estimată (lei)", "Data publicării", "Data estimată achiziție", "Obiect (trunchiat)",
                "Valoare (lei)", "Cuvinte cheie", "Link"],
          [[p_["judet"], DIST[p_["judet"]], len(plan_jud[p_["judet"]]),
            round(sum(x["valoare"] or 0 for x in plan_jud[p_["judet"]])), p_["data"], p_["data_est"], p_["titlu"],
            p_["valoare"], p_["cuvinte"], p_["link"]]
           for p_ in sorted(planuri_l, key=lambda x: (DIST[x["judet"]], -(x["valoare"] or 0)))],
          [12, 12, 12, 15, 12, 12, 30, 14, 30, 30])
    ws3 = wb.create_sheet("Concurenti")
    sheet(ws3, ["Denumire", "CUI", "Județ sediu", "Localitate", "Telefon", "E-mail", "Web", "Nr. contracte acoperiș",
                "Valoare totală (lei)", "Județe în care a câștigat", "Clienți"],
          [[c["denumire"], c["cui"], c["judet"], c["localitate"], c["telefon"], c["email"], c["web"], c["nr"],
            c["total"], c["judete_lucru"], c["clienti"]] for c in concurenti],
          [40, 12, 14, 16, 16, 28, 22, 10, 15, 30, 70])
    for w in (ws, ws2, ws4):
        for row in w.iter_rows(min_row=2):
            for c in row:
                if isinstance(c.value, dt.date):
                    c.number_format = "DD.MM.YYYY"
                elif isinstance(c.value, float):
                    c.number_format = "#,##0"
    for row in ws3.iter_rows(min_row=2):
        for c in row:
            if isinstance(c.value, float):
                c.number_format = "#,##0"
    out = os.path.join(HERE, "clienti_potentiali.xlsx")
    wb.save(out)
    json.dump(dict(achizitii=len(rows), clienti=len(clienti), concurenti=len(concurenti),
                   perioada=[str(min(r["data"] for r in rows if r["data"])), str(max(r["data"] for r in rows if r["data"]))],
                   top20=[(c["scor"], c["denumire"], c["judet"], c["nr"], c["nr_plan"], c["total"], c["km"]) for c in clienti[:20]]),
              open(os.path.join(D, "sumar.json"), "w"), ensure_ascii=False, indent=1, default=str)
    print(open(os.path.join(D, "sumar.json")).read())

if __name__ == "__main__":
    main()
