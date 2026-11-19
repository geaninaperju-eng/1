#!/usr/bin/env python3
"""Oferte de FURNIZARE (materiale pentru acoperiș) – dosar .docx + propunere financiară .xlsx cu formule.

Diferă de genereaza_dosare_bilzi.py (lucrări, F3 + analize): aici oferta e un centralizator de prețuri pe produse,
cu transportul repartizat în prețul unitar și profitul aplicat uniform.
Prețurile de cost vin din preturi/resurse.csv (coloana sursa spune de unde).

Rulare: python3 oferta_furnizare_bilzi.py um01151
"""
import csv
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Font

import genereaza_dosare_bilzi as g
from genereaza_dosare_bilzi import B, Y, P, FIRMA, antet, semnatura, table, page_break, bullet, lei

ROOT = Path(__file__).resolve().parent
RES = {r["cod"]: r for r in csv.DictReader(open(ROOT / "preturi" / "resurse.csv", encoding="utf-8"), delimiter=";")}

FURNIZARI = {
    "um01151": {
        "fisier": "ADV1558282_UM01151_Bucuresti_materiale_invelitoare",
        "titlu": "Achiziție bunuri materiale destinate înlocuirii învelitorii din azbociment",
        "autoritate": "UNITATEA MILITARĂ 01151 BUCUREȘTI",
        "adresa": "Str. Ghencea nr. 35-37, sector 6, București, cod 061695",
        "contact": "Mădălina Tudor, tel. 021 421 09 00",
        "email_depunere": "achizitii.um01151@yahoo.com",
        "nr_anunt": "ADV1558282",
        "data_anunt": "18.11.2026",
        "termen": "23.11.2026",
        "ora": "08:00",
        "cpv": "03419000-0 – Cherestea",
        "valoare": 46700.00,
        "procedura": "cumpărare directă",
        "valabilitate": "30 (treizeci)",
        "livrare": "în maximum 10 (zece) zile calendaristice de la comanda fermă, la sediul UM 01151, "
                   "Str. Ghencea nr. 35-37, București, inclusiv descărcarea",
        "profit": 0.15,
        "transport": "F07",
        "decizie": Y("persoanele cu funcții de decizie din UM 01151 – se cer telefonic/e-mail la Mădălina Tudor "
                     "(nu sunt publicate în anunț)"),
        # (nr, denumire din specificația tehnică, cerințe, UM, cantitate, cod resursă, cantitate resursă / UM ofertă,
        #  produs propus – marcă / producător)
        "produse": [
            ("1", "Cherestea (scândură) 5 m", "rășinoase; L = 5 m; grosime 2,5 cm; lățime 15–25 cm", "mc", 12,
             "F01", 1, "Scândură rășinoase netivită/tivită 25 mm, L 5 m – " ),
            ("2", "Dulapi rășinoase 5×10×500 cm", "rășinoase; L = 5 m; lățime 10 cm; grosime 5 cm", "buc", 135,
             "F02", 1, "Dulap rășinoase 50×100×5000 mm, uscat – "),
            ("3", "Cuie construcții", "3×70 mm; pentru lucrări de construcții; ambalaj cutie 5 kg", "kg", 100,
             "F03", 1, "Cuie construcții 3×70 mm, cutie 5 kg – "),
            ("4", "Tablă zincată cutată", "pentru învelitori pe suport continuu; oțel zincat DX51; grosime 0,4 mm; "
             "înălțime cută 10 mm; lățime totală 0,91 m; lățime utilă 0,86 m; lungime 2 m", "mp", 650,
             "F04", 1, "Tablă cutată zincată 0,40 mm, oțel DX51D+Z, profil h=10 mm, 910/860 mm, L 2 m – "),
            ("5", "Șuruburi autoforante 35 mm", "prindere tablă pe lemn; oțel carbon + zinc; L 35 mm; Ø 4,8 mm; "
             "șaibă oțel + cauciuc (EPDM); cutie 250 buc", "cutie", 30,
             "F05", 1, "Șurub autoforant pentru lemn 4,8×35 mm cu șaibă EPDM, zincat, cutie 250 buc – "),
            ("6", "Holșurub 130 mm", "oțel carbon; 6×130 mm; cap înecat; filet pe toată lungimea; vârf ascuțit; "
             "cutie 200 buc", "cutie", 2,
             "F06", 1, "Holșurub 6×130 mm cap înecat, filet total, zincat, cutie 200 buc – "),
        ],
        "marci": ["Total Forest Construct / Depozit Cherestea Chitila (de confirmat)",
                  "Total Forest Construct / Depozit Cherestea Chitila (de confirmat)",
                  "Romcuie / Dedeman (de confirmat)",
                  "BILKA Steel – profil de confirmat (h 10 mm, 910/860 mm)",
                  "Friulsider / Würth (de confirmat)",
                  "Friulsider / Würth (de confirmat)"],
    },
}


def cost(cod):
    return float(RES[cod]["pret_lei"])


def calcul(F):
    """Prețurile unitare (lei fără TVA) – exact aceleași formule ca în foaia „Calcul” din Excel."""
    tot_mat = sum(p[4] * p[6] * cost(p[5]) for p in F["produse"])
    transp = cost(F["transport"])
    rows = []
    for p in F["produse"]:
        c = p[6] * cost(p[5])
        cu_tr = c * (1 + transp / tot_mat)
        pu = round(cu_tr * (1 + F["profit"]), 2)
        rows.append((p, c, cu_tr, pu, round(pu * p[4], 2)))
    return rows, tot_mat, transp


# ------------------------------------------------------------------------------------------------ docx

def genereaza_docx(cheie):
    F = FURNIZARI[cheie]
    rows, _, _ = calcul(F)
    total = sum(r[4] for r in rows)
    doc = g.new_doc()
    opis = ["Scrisoare de înaintare și opis",
            "Formular de ofertă (propunerea financiară)",
            "Anexa la formularul de ofertă – centralizator de prețuri pe produse",
            "Propunerea tehnică – conformitatea cu specificațiile tehnice (marcă / producător)",
            "Fotografii ale produselor, fișe tehnice și certificate de calitate / declarații de conformitate",
            "Declarație privind conflictul de interese (art. 59–60 din Legea nr. 98/2016)",
            "Declarație privind neîncadrarea în art. 164, 165 și 167 din Legea nr. 98/2016",
            "Declarație privind respectarea art. 51 din Legea nr. 98/2016",
            "Declarație de consimțământ privind prelucrarea datelor cu caracter personal",
            "Certificat constatator ONRC (copie)"]

    # 1. scrisoare
    antet(doc)
    P(doc, "Către: ", B(F["autoritate"]), after=0)
    P(doc, F["adresa"], after=0)
    P(doc, "E-mail: ", F["email_depunere"], after=12)
    P(doc, B("SCRISOARE DE ÎNAINTARE"), align="c", size=13, after=10)
    P(doc, f"Ca urmare a anunțului de cumpărare directă {F['nr_anunt']} publicat în SEAP la data de {F['data_anunt']}, "
           f"privind „{F['titlu']}”, cod CPV {F['cpv']}, noi, {FIRMA['denumire']}, vă transmitem alăturat oferta noastră "
           f"de preț, în valoare totală de {lei(total)} lei fără TVA, la care se adaugă TVA, cu transportul și "
           "descărcarea la sediul unității incluse.", align="j")
    P(doc, f"Persoană de contact: {FIRMA['admin']}, telefon {FIRMA['tel']}, e-mail {FIRMA['email']}.", after=10)
    P(doc, B("OPISUL DOCUMENTELOR"), align="c", after=6)
    table(doc, ["Nr.", "Document", "Pagina"], [[str(i + 1), d, None] for i, d in enumerate(opis)],
          widths=[1.2, 13, 2.8])
    P(doc, "Data: ", None)
    semnatura(doc, data=False)
    page_break(doc)

    # 2. formular de ofertă
    antet(doc)
    P(doc, "Către: ", B(F["autoritate"]), after=8)
    P(doc, B("FORMULAR DE OFERTĂ"), align="c", size=13, after=10)
    P(doc, f"1. Examinând anunțul de cumpărare directă nr. {F['nr_anunt']} și specificațiile tehnice, subsemnatul "
           f"{FIRMA['admin']}, reprezentant al ofertantului {FIRMA['denumire']}, ne oferim ca, în conformitate cu "
           "prevederile și cerințele cuprinse în documentația mai sus menționată, să furnizăm produsele solicitate, "
           "pentru suma de ", B(f"{lei(total)} lei"), ", la care se adaugă taxa pe valoarea adăugată în valoare de ",
      B(f"{lei(round(total * 0.21, 2))} lei"), f" (TVA 21%), respectiv {lei(round(total * 1.21, 2))} lei cu TVA.",
      align="j")
    P(doc, f"2. Ne angajăm, în cazul în care oferta noastră este stabilită câștigătoare, să livrăm produsele {F['livrare']}.",
      align="j")
    P(doc, f"3. Ne angajăm să menținem această ofertă valabilă pentru o durată de {F['valabilitate']} zile, respectiv "
           f"până la data de {Y('23.12.2026')}, și ea va rămâne obligatorie pentru noi și poate fi acceptată oricând "
           "înainte de expirarea perioadei de valabilitate.", align="j")
    P(doc, "4. Prețul ofertat este ferm, în lei, și include toate taxele, transportul la beneficiar în București și "
           "descărcarea. Oferta se face pentru întreg lotul (toate cele 6 produse).", align="j")
    P(doc, "5. Precizăm că nu depunem ofertă alternativă. Am înțeles că nu sunteți obligați să acceptați oferta cu cel mai "
           "scăzut preț sau orice altă ofertă pe care o puteți primi.", align="j")
    semnatura(doc)
    page_break(doc)

    # 3. anexa – centralizator
    antet(doc)
    P(doc, B("ANEXĂ LA FORMULARUL DE OFERTĂ – CENTRALIZATOR DE PREȚURI"), align="c", size=13, after=8)
    P(doc, f"„{F['titlu']}” – anunț {F['nr_anunt']}", align="c", after=8)
    tr = [[p[0], p[1], p[3], f"{p[4]:g}", lei(pu), lei(val)] for p, _, _, pu, val in rows]
    tr.append(["", B("TOTAL fără TVA"), "", "", "", B(lei(total))])
    tr.append(["", "TVA 21%", "", "", "", lei(round(total * 0.21, 2))])
    tr.append(["", B("TOTAL cu TVA"), "", "", "", B(lei(round(total * 1.21, 2)))])
    table(doc, ["Nr.", "Produs", "UM", "Cantitate", "Preț unitar lei fără TVA", "Valoare lei fără TVA"], tr,
          widths=[1, 7, 1.5, 2, 2.6, 2.9])
    P(doc, "Prețurile unitare includ transportul și descărcarea la sediul UM 01151 București.", size=10)
    semnatura(doc)
    page_break(doc)

    # 4. propunere tehnică
    antet(doc)
    P(doc, B("PROPUNERE TEHNICĂ"), align="c", size=13, after=0)
    P(doc, "conformitatea produselor ofertate cu specificațiile tehnice", align="c", after=8)
    pt = [[p[0], f"{p[1]}: {p[2]}", f"{p[4]:g} {p[3]}", [p[7], Y(F["marci"][i])], "DA"]
          for i, (p, *_rest) in enumerate(rows)]
    table(doc, ["Nr.", "Cerință (specificații tehnice UM 01151)", "Cantitate", "Produs ofertat – marcă / producător",
                "Conform"], pt, widths=[0.8, 6.5, 2, 6, 1.5], size=9)
    for t in [
        "Fiecare produs este însoțit de fotografie, fișă tehnică și certificat de calitate / declarație de conformitate "
        "(anexate la ofertă), conform condițiilor din anunț.",
        f"Livrare {F['livrare']}; recepția cantitativă și calitativă la sediul beneficiarului.",
        "Garanție: conform garanției producătorului pentru tablă și elementele de fixare; cheresteaua se livrează "
        "sănătoasă, fără putregai, cu umiditatea specifică lemnului de construcții.",
        "Produsele neconforme se înlocuiesc pe cheltuiala noastră în maximum 3 zile lucrătoare.",
    ]:
        bullet(doc, t)
    P(doc, Y("De făcut înainte de trimitere: verificați la furnizor profilul tablei (cută 10 mm, lățime 910/860 mm) și "
             "trecem marca exactă; atașați fotografiile, fișele tehnice și certificatele de calitate pentru toate "
             "cele 6 produse – fără ele oferta nu intră în evaluare."), size=9)
    semnatura(doc)
    page_break(doc)

    # 5. declarații
    L = {"titlu": F["titlu"], "procedura": F["procedura"], "autoritate": F["autoritate"], "decizie": F["decizie"]}
    g.decl_art60(doc, L)
    page_break(doc)
    g.decl_164_165_167(doc, L)
    page_break(doc)
    g.decl_art51(doc, L)
    page_break(doc)
    antet(doc)
    P(doc, B("DECLARAȚIE DE CONSIMȚĂMÂNT"), align="c", size=13, after=0)
    P(doc, "privind prelucrarea datelor cu caracter personal", align="c", after=8)
    P(doc, f"Subsemnatul {FIRMA['admin']}, în calitate de administrator al {FIRMA['denumire']}, îmi exprim acordul ca "
           f"{F['autoritate']} să prelucreze datele cu caracter personal cuprinse în oferta depusă pentru „{F['titlu']}”, "
           "exclusiv în scopul derulării procedurii de achiziție și a eventualului contract, cu respectarea "
           "Regulamentului (UE) 2016/679 și a Legii nr. 190/2018.", align="j")
    semnatura(doc)
    g.footer(doc, f"{FIRMA['denumire']} – oferta pentru {F['nr_anunt']}")
    g.OUT.mkdir(exist_ok=True)
    path = g.OUT / f"Dosar_oferta_BILZI_{F['fisier']}.docx"
    doc.save(path)
    g.compacteaza_docx(path)
    return path, total


# ------------------------------------------------------------------------------------------------ xlsx

def genereaza_xlsx(cheie):
    F = FURNIZARI[cheie]
    rows, tot_mat, transp = calcul(F)
    wb = openpyxl.Workbook()
    bold = Font(bold=True)

    ws = wb.active
    ws.title = "Oferta"
    ws.append([f"{FIRMA['denumire']} – CENTRALIZATOR DE PREȚURI – {F['nr_anunt']} – {F['autoritate']}"])
    ws.append([F["titlu"]])
    ws.append(["Nr.", "Produs", "UM", "Cantitate", "Preț unitar lei fără TVA", "Valoare lei fără TVA"])
    n = len(rows)
    for i, (p, *_r) in enumerate(rows):
        r = 4 + i
        ws.append([p[0], p[1], p[3], p[4], f"=Calcul!I{r}", f"=ROUND(D{r}*E{r},2)"])
    t = 4 + n
    ws.append(["", "TOTAL fără TVA", "", "", "", f"=SUM(F4:F{t - 1})"])
    ws.append(["", "TVA 21%", "", "", "", f"=ROUND(F{t}*0.21,2)"])
    ws.append(["", "TOTAL cu TVA", "", "", "", f"=F{t}+F{t + 1}"])
    ws.append([])
    ws.append(["", "Valoare estimată (lei fără TVA)", "", "", "", F["valoare"]])
    ws.append(["", "Oferta ca % din valoarea estimată", "", "", "", f"=F{t}/F{t + 4}"])
    ws.append(["", "Verificare: oferta ≤ 97% din estimat (0 = OK)", "", "", "", f"=IF(F{t + 5}<=0.97,0,1)"])
    ws.append(["", "Verificare: total Ofertă − total Calcul (≈ 0)", "", "", "", f"=ROUND(F{t}-Calcul!J{t},2)"])
    for c in ("B", "F"):
        ws[f"{c}{t}"].font = bold
    ws.column_dimensions["B"].width = 50
    for c in "EF":
        ws.column_dimensions[c].width = 18
    for r in range(4, t + 5):
        for c in "EF":
            ws[f"{c}{r}"].number_format = g.LEI
    ws[f"F{t + 5}"].number_format = "0.00%"

    wc = wb.create_sheet("Calcul")
    wc.append(["Calculul prețului unitar: cost material (Resurse) + transport repartizat proporțional + profit"])
    wc.append(["Profit", F["profit"], "", "Transport (lei)", f"=VLOOKUP(\"{F['transport']}\",Resurse!$A$2:$E$20,5,FALSE)",
               "Total materiale (lei)", f"=SUM(K4:K{3 + n})"])
    wc.append(["Nr.", "Produs", "Cod resursă", "Cantitate", "Consum resursă / UM", "Cost unitar material",
               "Cost cu transport", "—", "Preț unitar ofertă", "Valoare", "Cost material total", "Cost cu transport total"])
    for i, (p, *_r) in enumerate(rows):
        r = 4 + i
        wc.append([p[0], p[1], p[5], p[4], p[6], f"=E{r}*VLOOKUP(C{r},Resurse!$A$2:$E$20,5,FALSE)",
                   f"=F{r}*(1+$E$2/$G$2)", "", f"=ROUND(G{r}*(1+$B$2),2)", f"=ROUND(D{r}*I{r},2)", f"=D{r}*F{r}", f"=D{r}*G{r}"])
    wc.append(["", "TOTAL", "", "", "", "", f"=SUM(L4:L{t - 1})", "", "", f"=SUM(J4:J{t - 1})"])
    wc.append(["", "Verificare: cost cu transport − (materiale + transport) (≈ 0)", "", "", "", "", "", "", "",
               f"=ROUND(G{t}-G2-E2,2)"])
    wc.append(["", "Profit efectiv (lei)", "", "", "", "", "", "", "", f"=J{t}-G{t}"])
    wc.column_dimensions["B"].width = 40
    for r in range(2, t + 3):
        for c in "EFGIJ":
            if r > 3 or c in "EG":
                wc[f"{c}{r}"].number_format = g.LEI
    wc["B2"].number_format = "0%"

    wr = wb.create_sheet("Resurse")
    wr.append(["Cod", "Denumire", "UM", "Categorie", "Preț lei fără TVA", "Sursa"])
    coduri = [p[5] for p in F["produse"]] + [F["transport"]]
    for c in coduri:
        x = RES[c]
        wr.append([c, x["denumire"], x["um"], x["categorie"], float(x["pret_lei"]), x["sursa"]])
    wr.column_dimensions["B"].width = 55
    wr.column_dimensions["F"].width = 90

    wt = wb.create_sheet("Specificatii")
    wt.append(["Nr.", "Produs", "Cerințe UM 01151", "UM", "Cantitate", "Produs ofertat", "Marcă / producător"])
    for i, p in enumerate(F["produse"]):
        wt.append([p[0], p[1], p[2], p[3], p[4], p[7].rstrip(" –"), F["marci"][i]])
    wt.column_dimensions["C"].width = 70
    wt.column_dimensions["F"].width = 60

    path = g.OUT / f"Propunere_financiara_BILZI_{F['fisier']}.xlsx"
    wb.save(path)
    return path


if __name__ == "__main__":
    for k in sys.argv[1:] or FURNIZARI:
        d, total = genereaza_docx(k)
        x = genereaza_xlsx(k)
        F = FURNIZARI[k]
        print(d, x, f"total {lei(total)} lei = {total / F['valoare']:.1%} din estimat")
