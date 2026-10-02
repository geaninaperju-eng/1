#!/usr/bin/env python3
"""Dosare de ofertă BILZI STEEL PROFILE S.R.L. – licitații acoperișuri (02.10.2026).

Model: dosarul „Lucrări acoperiș hala producție S.D.F. Onești” (ADV1550112), depus pe 01.10.2026.
Ieșire: output/*.docx și output/*.xlsx

Datele care depind de caietul de sarcini (cantități F3, durată, cod CPV, ora limită,
persoane cu funcții de decizie) se tipăresc ca spații de completat, evidențiate cu galben.
"""
import copy
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import PatternFill, Font
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "output"
BLANK = "____________"

FIRMA = {
    "denumire": "BILZI STEEL PROFILE S.R.L.",
    "cui": "39666285",
    "rc": "J10/757/2018",
    "sediu": "Municipiul Buzău, Șoseaua Nordului nr. 9C, județul Buzău",
    "admin": "Șerban Marius",
    "tel": "0761 144 238",
    "email": "marius807@yahoo.com",
}

# Experiență similară dovedibilă (vezi „Experiență similară BILZI STEEL PROFILE – praguri”)
EXPERIENTA = [
    ("Reabilitare și înlocuire acoperiș cămin cultural, sat Lilieci",
     "Comuna Sinești, jud. Ialomița", "nr. 2309 / 23.05.2024", "80.000,00",
     "23.05.2024 – 03.06.2024", "PV de recepție nr. 2455 / 03.06.2024"),
    ("Reparație acoperiș clădire publică Post Poliție Alexeni",
     "UAT Comuna Alexeni, jud. Ialomița", "nr. 2005 / 05.03.2024", "42.454,94",
     "05.03.2024 – 04.04.2024", "PV de recepție nr. 2453 / 04.04.2024"),
]

LICITATII = {
    "dofteana": {
        "fisier": "Parc_Dofteana",
        "titlu": "Înlocuire acoperiș clădire Parc Dofteana – DSBC 2026 II",
        "autoritate": "REGIA NAȚIONALĂ A PĂDURILOR – ROMSILVA prin DIRECȚIA SILVICĂ BACĂU",
        "autoritate_scurt": "Direcția Silvică Bacău",
        "email_depunere": "anunturi.seap@bacau.rosilva.ro",
        "nr_anunt": None,
        "data_anunt": "01.10.2026",
        "termen": "07.10.2026",
        "ora": None,
        "cpv": None,
        "valoare": 252390.00,
        "amplasament": "Clădirea din Parcul Dofteana, comuna Dofteana, județul Bacău (Direcția Silvică Bacău)",
        "procedura": "cumpărare directă",
        "durata": None,
        "catalog": True,
        "decizie": "Pădureanu Leonard – Director, Bitir Ioan – Director Tehnic, Tabacaru Ion – Director Economic",
        "experienta": False,
        "acces": False,
    },
    "c4": {
        "fisier": "Cazarma_3589_Pav_C4",
        "titlu": "Lucrări de reparații acoperiș pav. C4, cazarma 3589 București",
        "autoritate": None,  # unitatea militară care administrează cazarma 3589
        "autoritate_scurt": None,
        "email_depunere": None,
        "nr_anunt": None,
        "data_anunt": "30.09.2026",
        "termen": "05.10.2026",
        "ora": None,
        "cpv": None,
        "valoare": 77677.00,
        "amplasament": "Pavilionul C4 din cazarma 3589, municipiul București",
        "procedura": "cumpărare directă",
        "durata": None,
        "catalog": True,
        "decizie": None,
        "experienta": True,
        "acces": True,
    },
}


# ---------------------------------------------------------------- utilitare docx

def lei(v):
    s = f"{v:,.2f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def new_doc():
    doc = Document()
    sec = doc.sections[0]
    sec.page_height, sec.page_width = Cm(29.7), Cm(21.0)
    sec.left_margin = sec.right_margin = Cm(2.0)
    sec.top_margin = sec.bottom_margin = Cm(1.8)
    st = doc.styles["Normal"]
    st.font.name = "Times New Roman"
    st.font.size = Pt(11)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    st.paragraph_format.space_after = Pt(3)
    return doc


def add_runs(par, parts, size=None):
    """parts: str sau (text, opțiuni). None -> spațiu de completat galben."""
    for part in parts:
        opts = {}
        if isinstance(part, tuple):
            part, opts = part
        if part is None:
            r = par.add_run(opts.get("blank", BLANK))
            r.font.highlight_color = WD_COLOR_INDEX.YELLOW
        else:
            r = par.add_run(str(part))
            if opts.get("y"):
                r.font.highlight_color = WD_COLOR_INDEX.YELLOW
        r.bold = opts.get("b", False)
        r.italic = opts.get("i", False)
        if size or opts.get("size"):
            r.font.size = Pt(opts.get("size", size))
    return par


def P(doc, *parts, align=None, size=None, after=None, before=None):
    par = doc.add_paragraph()
    add_runs(par, parts, size)
    par.alignment = {"c": WD_ALIGN_PARAGRAPH.CENTER, "r": WD_ALIGN_PARAGRAPH.RIGHT,
                     "j": WD_ALIGN_PARAGRAPH.JUSTIFY}.get(align, par.alignment)
    if after is not None:
        par.paragraph_format.space_after = Pt(after)
    if before is not None:
        par.paragraph_format.space_before = Pt(before)
    return par


def B(text):
    return (text, {"b": True})


def Y(text):
    """Text propus, de verificat – evidențiat cu galben."""
    return (text, {"y": True})


def bullet(doc, *parts):
    par = doc.add_paragraph(style="List Bullet")
    add_runs(par, parts)
    par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    return par


def shade(cell, color="D9E2F3"):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color)
    tcPr.append(shd)


def table(doc, header, rows, widths=None, size=10):
    t = doc.add_table(rows=1, cols=len(header))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(header):
        c = t.rows[0].cells[i]
        add_runs(c.paragraphs[0], [B(h)], size)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        shade(c)
    for row in rows:
        cells = t.add_row().cells
        for i, val in enumerate(row):
            parts = val if isinstance(val, list) else [val]
            add_runs(cells[i].paragraphs[0], parts, size)
    if widths:
        for row in t.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Cm(w)
    doc.add_paragraph()
    return t


def page_break(doc):
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def antet(doc):
    f = FIRMA
    P(doc, "OPERATOR ECONOMIC", after=0, size=10)
    P(doc, B(f["denumire"]), after=0)
    P(doc, f"CUI {f['cui']}, {f['rc']}", after=0, size=10)
    P(doc, f["sediu"], after=10, size=10)


def semnatura(doc, data=True):
    if data:
        P(doc, "Data completării: ", None, before=12)
    P(doc, FIRMA["admin"], before=6, after=0)
    P(doc, "(administrator)", after=0)
    P(doc, "Semnătura autorizată și ștampila")


def footer(doc, text):
    for sec in doc.sections:
        par = sec.footer.paragraphs[0]
        par.text = ""
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = par.add_run(text + " – pagina ")
        r.font.size = Pt(8)
        fld = OxmlElement("w:fldSimple")
        fld.set(qn("w:instr"), "PAGE")
        run = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.text = "1"
        run.append(t)
        fld.append(run)
        par._p.append(fld)


# ---------------------------------------------------------------- secțiuni dosar

def scrisoare(doc, L, opis):
    antet(doc)
    P(doc, "Către: ", B(L["autoritate"]) if L["autoritate"] else None, after=0)
    P(doc, "E-mail: ", L["email_depunere"], after=12)
    P(doc, B("SCRISOARE DE ÎNAINTARE"), align="c", size=13, after=10)
    parts = ["Ca urmare a anunțului de cumpărare directă ", L["nr_anunt"],
             f" publicat în SEAP la data de {L['data_anunt']}, privind atribuirea contractului „{L['titlu']}”, cod CPV ",
             L["cpv"], f", noi, {FIRMA['denumire']}, vă transmitem alăturat oferta noastră."]
    P(doc, *parts, align="j")
    if L["catalog"]:
        P(doc, "Prețul ofertat a fost publicat în catalogul electronic SEAP, sub denumirea și codul CPV din anunț, "
               f"până la data de {L['termen']}, ora ", L["ora"], ".", align="j")
    P(doc, f"Persoană de contact: {FIRMA['admin']}, telefon {FIRMA['tel']}, e-mail {FIRMA['email']}.", after=10)
    P(doc, B("OPISUL DOCUMENTELOR"), align="c", after=6)
    table(doc, ["Nr.", "Document", "Pagina"],
          [[str(i + 1), d, None] for i, d in enumerate(opis)], widths=[1.2, 13, 2.8])
    P(doc, "Data: ", None)
    semnatura(doc, data=False)


def propunere_tehnica(doc, L):
    antet(doc)
    P(doc, B("PROPUNERE TEHNICĂ"), align="c", size=13, after=0)
    P(doc, f"pentru contractul „{L['titlu']}”", align="c", after=0)
    P(doc, "Cod CPV ", L["cpv"], align="c", after=10)

    P(doc, B("1. Date generale"))
    table(doc, ["Element", "Detalii"], [
        ["Autoritate contractantă", L["autoritate"]],
        ["Ofertant", f"{FIRMA['denumire']}, CUI {FIRMA['cui']}, {FIRMA['rc']}"],
        ["Amplasament", L["amplasament"]],
        ["Obiect", [f"Execuția lucrărilor „{L['titlu']}”, conform caietului de sarcini și listei de cantități F3"]],
        ["Durata de execuție ofertată", [L["durata"] if L["durata"] else None,
                                         " zile lucrătoare de la ordinul de începere, în limita duratei din caietul de sarcini"]],
        ["Începerea lucrărilor", "În maximum 72 de ore de la emiterea ordinului de începere"],
        ["Garanția lucrărilor", [Y("24"), " de luni calendaristice de la recepția la terminarea lucrărilor"]],
        ["Distanța sediu – amplasament", [Y(L["distanta"])]],
    ], widths=[5, 12])

    P(doc, B("2. Obiectul lucrărilor"))
    P(doc, "Ne angajăm să executăm integral lucrările prevăzute în caietul de sarcini și în lista de cantități F3. "
           "Lucrările cuprind, în principal:", align="j")
    for t in [
        "desfacerea învelitorii existente, a elementelor de tinichigerie și a sistemului pluvial degradat, cu sortarea "
        "materialelor recuperabile și predarea lor beneficiarului, pe bază de proces-verbal;",
        "verificarea structurii de rezistență a acoperișului (șarpantă, astereală, rigle) și înlocuirea elementelor din "
        "lemn degradate, cu lemn de rășinoase de calitate conform ST 014-1996;",
        "ignifugarea și tratarea antiseptică a elementelor din lemn noi și a celor păstrate;",
        "montarea foliei anticondens și a învelitorii noi din tablă tip țiglă, fixată cu șuruburi autoforante cu garnitură;",
        "lucrări de tinichigerie: coame cu profile de etanșare, borduri de fronton și de streașină, dolii și "
        "racordări la coșuri și calcane, opritori de zăpadă;",
        "montarea sistemului de colectare și evacuare a apelor pluviale (jgheaburi, burlane și accesorii);",
        "transportul materialelor și evacuarea molozului și a deșeurilor la depozite autorizate.",
    ]:
        bullet(doc, t)
    P(doc, Y("Lista de mai sus se aliniază la articolele din F3 după primirea caietului de sarcini."), size=9)

    P(doc, B("3. Organizarea de șantier"), before=6)
    for t in [
        "delimitarea și semnalizarea zonei de lucru și a zonei de depozitare a materialelor, stabilite împreună cu "
        "reprezentantul beneficiarului, astfel încât activitatea din clădire să fie afectată cât mai puțin;",
        "depozitarea materialelor pe suporți, ferite de umezeală, tabla fiind păstrată în ambalajul original până la montaj;",
        "schele, sisteme de acces la înălțime și protecții colective (balustrade, plase, linii de ancorare);",
        "acoperirea provizorie cu prelate a zonelor descoperite, la sfârșitul fiecărei zile și în caz de precipitații;",
        "colectarea molozului în containere sau zone delimitate și evacuarea lui periodică.",
    ]:
        bullet(doc, t)
    if L["acces"]:
        P(doc, "Accesul personalului și al vehiculelor în cazarmă se face numai pe baza listei nominale aprobate de "
               "unitate și cu respectarea regulilor interne privind accesul și protecția informațiilor.", align="j")

    P(doc, B("4. Tehnologia de execuție"), before=6)
    P(doc, "Lucrările se execută pe tronsoane, astfel încât suprafața descoperită simultan să poată fi închisă în aceeași zi.",
      align="j")
    for t in [
        "Etapa 1 – Predarea amplasamentului, trasarea și constatarea stării elementelor existente, consemnate în "
        "procesul-verbal de predare-primire a amplasamentului.",
        "Etapa 2 – Desfacerea învelitoarei existente, a tinichigeriei și a jgheaburilor, pe tronsoane.",
        "Etapa 3 – Remedierea structurii din lemn: înlocuirea căpriorilor, a asterelei și a riglelor degradate, "
        "verificarea pantelor și a planeității.",
        "Etapa 4 – Ignifugarea elementelor din lemn, conform fișei tehnice a produsului.",
        "Etapa 5 – Montarea foliei anticondens și a tablei tip țiglă, de la streașină spre coamă.",
        "Etapa 6 – Tinichigeria (coame, borduri, dolii, racorduri) și opritorii de zăpadă.",
        "Etapa 7 – Sistemul pluvial: cârlige, jgheaburi cu pantă spre racorduri, burlane cu coliere, etanșarea îmbinărilor.",
        "Etapa 8 – Curățenia generală, evacuarea molozului, verificarea etanșeității și notificarea scrisă a "
        "beneficiarului pentru recepție.",
    ]:
        bullet(doc, t)

    P(doc, B("5. Resurse umane"), before=6)
    table(doc, ["Funcție / meserie", "Număr", "Observații"], [
        ["Șef de echipă / conducător tehnic al lucrării", "1", "Răspunde de execuție și calitate"],
        ["Dulgheri", Y("3"), "Structură lemn, astereală, rigle"],
        ["Montatori învelitori metalice / tinichigii", Y("3"), "Învelitoare, tinichigerie, sistem pluvial"],
        ["Muncitori necalificați", Y("2"), "Desfaceri, manipulare materiale, curățenie"],
        ["Conducător auto", "1", "Transport materiale și moloz"],
    ], widths=[7, 2, 8])
    P(doc, "Tot personalul este instruit în domeniul securității și sănătății în muncă și pentru lucrul la înălțime.",
      align="j")

    P(doc, B("6. Utilaje, echipamente și unelte"), before=6)
    for t in ["autocamion pentru transportul materialelor și al molozului;",
              "schele metalice și scări, cu sisteme de ancorare;",
              "scule electrice profesionale: fierăstraie circulare, mașini de înșurubat, foarfeci pentru tablă, mașini de găurit;",
              "echipament pentru aplicarea soluției ignifuge;",
              "echipamente individuale de protecție: căști, hamuri, centuri de siguranță, mănuși, încălțăminte de protecție."]:
        bullet(doc, t)

    P(doc, B("7. Materiale"), before=6)
    P(doc, "Materialele respectă standardele și normele în vigoare în România și sunt însoțite de certificat de calitate "
           "sau declarație de performanță, certificat de garanție și, după caz, agrement tehnic. Învelitoarea, "
           "tinichigeria și sistemul pluvial sunt din gama BILKA: ",
      Y("tablă tip țiglă CLASIC mat 0,50 mm, culoare RAL ____"),
      ", coame, borduri și opritori de zăpadă mat 0,50 mm, jgheab 125 mm și burlan 90 mm colorate. "
      "Specificațiile care indică o anumită marcă se consideră însoțite de mențiunea „sau echivalent”.", align="j")

    P(doc, B("8. Asigurarea calității și recepția"), before=6)
    for t in ["verificarea materialelor la livrare, pe baza documentelor de însoțire;",
              "verificarea pe faze de execuție și procese-verbale de lucrări ascunse acolo unde este cazul;",
              "remedierea pe cheltuiala noastră a oricărei lucrări care nu respectă caietul de sarcini;",
              "situații de lucrări cu antemăsurătorile aferente, la finalizarea fiecărei etape."]:
        bullet(doc, t)
    P(doc, "Lucrările se execută cu respectarea Legii nr. 10/1995 privind calitatea în construcții, a HG nr. 273/1994 "
           "privind recepția lucrărilor, a normativelor NP 069-2014, C 37-1988, ST 014-1996, C 56-2002 și a celorlalte "
           "reglementări din caietul de sarcini.", align="j")

    P(doc, B("9. Securitate și sănătate în muncă, PSI și protecția mediului"), before=6)
    for t in ["respectarea Legii nr. 319/2006 a securității și sănătății în muncă, cu accent pe lucrul la înălțime;",
              "respectarea normelor PSI, inclusiv la folosirea sculelor electrice lângă materiale combustibile;",
              "limitarea prafului și a zgomotului, colectarea selectivă a deșeurilor și evacuarea lor la depozite autorizate."]:
        bullet(doc, t)

    P(doc, B("10. Grafic de execuție"), before=6)
    P(doc, "Graficul de execuție fizic și valoric este anexat la propunerea financiară.")
    P(doc, "Declarăm că propunerea tehnică respectă integral cerințele caietului de sarcini și că am ținut cont de "
           "obligațiile relevante din domeniile mediului, social și al relațiilor de muncă.", align="j")
    semnatura(doc)


def formular_oferta(doc, L):
    P(doc, B("FORMULAR DE OFERTĂ"), align="c", size=13, after=10)
    P(doc, "Către ", L["autoritate"])
    P(doc, f"Titlul contractului: „{L['titlu']}”", after=8)
    P(doc, f"1. Examinând documentația de atribuire, subsemnații, reprezentanți ai ofertantului {FIRMA['denumire']}, "
           f"declarăm că, în conformitate cu prevederile și cerințele cuprinse în documentația mai sus menționată, vom "
           f"executa „{L['titlu']}” pentru suma de ", None, " lei fără T.V.A., la care se adaugă T.V.A. în valoare de ",
      None, " lei, în conformitate cu anexa la formularul de ofertă.", align="j")
    P(doc, "2. Ne angajăm ca, în cazul în care oferta noastră este stabilită câștigătoare, să începem lucrările cât mai "
           "curând posibil după primirea ordinului de începere și să terminăm lucrările în ",
      L["durata"], " zile lucrătoare.", align="j")
    P(doc, "3. Ne angajăm să menținem această ofertă valabilă pentru o durată de ", Y("90 (nouăzeci)"),
      " de zile, respectiv până la data de ", None, ", și ea va rămâne obligatorie pentru noi și poate fi acceptată "
      "oricând înainte de expirarea perioadei de valabilitate.", align="j")
    P(doc, "4. Am înțeles și consimțim ca, în cazul în care oferta noastră este stabilită ca fiind câștigătoare, să "
           "constituim garanția de bună execuție în conformitate cu prevederile din documentația de atribuire.", align="j")
    P(doc, "5. Precizăm că:")
    P(doc, "[ ] depunem ofertă alternativă, ale cărei detalii sunt prezentate într-un formular de ofertă separat, marcat "
           "în mod clar „alternativă”;")
    P(doc, "[X] nu depunem ofertă alternativă.")
    P(doc, "6. Până la încheierea și semnarea contractului de achiziție publică, această ofertă, împreună cu comunicarea "
           "transmisă de dumneavoastră, prin care oferta noastră este acceptată ca fiind câștigătoare, vor constitui un "
           "contract angajant între noi.", align="j")
    P(doc, "7. Înțelegem că nu sunteți obligați să acceptați oferta cu cel mai scăzut preț sau orice altă ofertă primită.",
      align="j")
    P(doc, "Data: ", None, before=8)
    P(doc, f"{FIRMA['admin']}, în calitate de administrator, legal autorizat să semnez oferta pentru și în numele "
           f"{FIRMA['denumire']}.", align="j")
    P(doc, "OPERATOR ECONOMIC", after=0)
    P(doc, B(FIRMA["denumire"]), after=0)
    P(doc, "Semnătura autorizată și ștampila")


def anexa_oferta(doc, L):
    P(doc, B("ANEXA LA FORMULARUL DE OFERTĂ"), align="c", size=13, after=10)
    P(doc, f"Titlul contractului: „{L['titlu']}”", after=8)
    table(doc, ["Denumire", "Valoare lei (fără T.V.A.)"], [
        [f"{L['titlu']} – total general deviz ofertă", None],
        ["T.V.A. 21%", None],
        ["Total cu T.V.A.", None],
    ], widths=[12, 5])
    P(doc, f"Valoarea estimată de autoritatea contractantă: {lei(L['valoare'])} lei fără T.V.A. Oferta nu o poate depăși.",
      size=9)
    P(doc, "Notă: Propunerea financiară cuprinde, pe lângă Formularul de ofertă și Anexa la acesta, Lista cu cantitățile "
           "de lucrări (F3) cu prețuri unitare, listele de consumuri C6, C7, C8, C9 și Graficul de execuție fizic și valoric.",
      align="j")
    semnatura(doc)


def decl_contract(doc, L):
    antet(doc)
    P(doc, B("DECLARAȚIE"), align="c", size=13, after=0)
    P(doc, "privind însușirea modelului de contract", align="c", after=8)
    P(doc, f"Titlul contractului: „{L['titlu']}”", after=8)
    P(doc, f"Subsemnatul {FIRMA['admin']}, reprezentant legal al {FIRMA['denumire']}, declar pe propria răspundere că, "
           f"în calitate de ofertant la procedura de {L['procedura']} organizată de ", L["autoritate"],
      f" pentru atribuirea contractului „{L['titlu']}”, ne însușim modelul de contract pus la dispoziție de autoritatea "
      "contractantă în cadrul documentației de atribuire.", align="j")
    P(doc, "În cazul în care oferta noastră va fi declarată câștigătoare, ne angajăm să semnăm contractul în forma "
           "prezentată în documentația de atribuire.", align="j")
    semnatura(doc)


def decl_art51(doc, L):
    antet(doc)
    P(doc, B("DECLARAȚIE PRIVIND SĂNĂTATEA ȘI SECURITATEA ÎN MUNCĂ"), align="c", size=13, after=0)
    P(doc, "(privind respectarea art. 51 din Legea nr. 98/2016 privind achizițiile publice)", align="c", after=8)
    P(doc, f"Titlul contractului: „{L['titlu']}”", after=8)
    P(doc, f"Subsemnatul {FIRMA['admin']}, reprezentant împuternicit al {FIRMA['denumire']}, declar pe propria "
           "răspundere că mă angajez să realizez activitățile din contract în conformitate cu reglementările obligatorii "
           "în domeniile mediului, social și al relațiilor de muncă stabilite prin legislația adoptată la nivelul Uniunii "
           "Europene, legislația națională, prin acorduri colective sau prin tratatele, convențiile și acordurile "
           "internaționale în aceste domenii. De asemenea, declar pe propria răspundere că la elaborarea ofertei am ținut "
           "cont de obligațiile în domeniul mediului, social și al relațiilor de muncă și am inclus în ofertă costul "
           "pentru îndeplinirea acestor obligații.", align="j")
    P(doc, "Totodată, declar că am luat la cunoștință de prevederile art. 326 „Falsul în declarații” din Codul Penal.",
      align="j")
    semnatura(doc)


def decl_art60(doc, L):
    antet(doc)
    P(doc, B("DECLARAȚIE PRIVIND EVITAREA CONFLICTULUI DE INTERESE"), align="c", size=13, after=0)
    P(doc, "potrivit art. 59–60 din Legea nr. 98/2016 privind achizițiile publice", align="c", after=8)
    P(doc, f"Titlul contractului: „{L['titlu']}”", after=8)
    P(doc, f"1. Subsemnatul {FIRMA['admin']}, reprezentant împuternicit al {FIRMA['denumire']}, cu sediul în "
           f"{FIRMA['sediu']}, în calitate de ofertant la procedura de {L['procedura']}, declar pe propria răspundere, "
           "sub sancțiunea excluderii din procedură și sub sancțiunile aplicabile faptei de fals în acte publice, că nu "
           "mă aflu în situațiile prevăzute la art. 60 din Legea nr. 98/2016.", align="j")
    P(doc, "Persoanele ce dețin funcții de decizie în cadrul autorității contractante, nominalizate potrivit art. 21 "
           "alin. (5) și (6) din H.G. nr. 395/2016, sunt: ", L["decizie"], ".", align="j")
    P(doc, "2. Voi informa imediat autoritatea contractantă dacă vor interveni modificări în prezenta declarație pe "
           "parcursul procedurii sau, în cazul în care vom fi desemnați câștigători, pe parcursul derulării contractului, "
           "având în vedere și prevederile art. 61 din Legea nr. 98/2016.", align="j")
    P(doc, "Informațiile furnizate sunt complete și corecte în fiecare detaliu. Înțeleg că autoritatea contractantă are "
           "dreptul de a solicita orice documente doveditoare și că, în cazul în care această declarație nu este conformă "
           "cu realitatea, sunt pasibil de încălcarea prevederilor legislației penale privind falsul în declarații.",
      align="j")
    semnatura(doc)


def decl_164_165_167(doc, L):
    antet(doc)
    P(doc, B("DECLARAȚIE PRIVIND NEÎNCADRAREA ÎN SITUAȚIILE"), align="c", size=13, after=0)
    P(doc, B("PREVĂZUTE LA ART. 164, 165 ȘI 167 DIN LEGEA NR. 98/2016"), align="c", size=13, after=8)
    P(doc, f"Titlul contractului: „{L['titlu']}”", after=8)
    P(doc, f"Subsemnatul {FIRMA['admin']}, reprezentant legal al {FIRMA['denumire']}, în calitate de ofertant, declar pe "
           "propria răspundere, sub sancțiunea excluderii din procedură și a sancțiunilor aplicate faptei de fals în "
           "declarații, că:", align="j")
    for t in [
        "nici operatorul economic, nici administratorul acestuia nu au fost condamnați prin hotărâre definitivă pentru "
        "vreuna dintre infracțiunile prevăzute la art. 164 alin. (1) din Legea nr. 98/2016;",
        "operatorul economic și-a îndeplinit obligațiile de plată a impozitelor, taxelor și contribuțiilor la bugetele "
        "componente ale bugetului general consolidat, conform art. 165 din Legea nr. 98/2016;",
        "operatorul economic nu se află în niciuna dintre situațiile prevăzute la art. 167 alin. (1) din Legea nr. 98/2016 "
        "(nu este în insolvență, faliment sau lichidare, nu a încălcat obligațiile din domeniul mediului, social și al "
        "relațiilor de muncă, nu a avut deficiențe grave în executarea unor contracte anterioare și nu a prezentat "
        "informații false).",
    ]:
        bullet(doc, t)
    P(doc, "Înțeleg că autoritatea contractantă are dreptul de a solicita documente doveditoare (certificate de atestare "
           "fiscală, cazier judiciar și fiscal) în scopul verificării acestei declarații.", align="j")
    semnatura(doc)


def experienta(doc, L):
    antet(doc)
    P(doc, B("LISTA PRINCIPALELOR LUCRĂRI SIMILARE EXECUTATE"), align="c", size=13, after=0)
    P(doc, "în ultimii 5 ani", align="c", after=8)
    P(doc, f"Titlul contractului: „{L['titlu']}”", after=8)
    table(doc, ["Nr.", "Obiectul contractului", "Beneficiar", "Contract", "Valoare fără TVA (lei)", "Perioada",
                "Document constatator"],
          [[str(i + 1), e[0], e[1], e[2], e[3], e[4], e[5]] for i, e in enumerate(EXPERIENTA)],
          widths=[0.8, 4, 3, 2.4, 2, 2.4, 2.6], size=9)
    P(doc, f"{FIRMA['denumire']} a executat lucrările de mai sus în calitate de contractant unic. Anexăm copii ale "
           "documentelor constatatoare (contract, proces-verbal de recepție la terminarea lucrărilor).", align="j")
    P(doc, B("Total: 122.454,94 lei fără TVA, în 2 contracte cu autorități publice."))
    P(doc, Y("Atenție: PV-ul Alexeni nr. 2453/04.04.2024 îl trece pe Șerban Marius ca reprezentant al „SC LAVITEX PROD SRL” "
             "(eroare materială). Cereți Primăriei Alexeni o rectificare sau o recomandare pe BILZI, altfel se poate contesta."),
      size=9)
    semnatura(doc)


def lista_personal(doc, L):
    antet(doc)
    P(doc, B("LISTA NOMINALĂ A PERSONALULUI ȘI A AUTOVEHICULELOR"), align="c", size=13, after=0)
    P(doc, "pentru accesul în cazarmă pe durata execuției lucrărilor", align="c", after=8)
    P(doc, f"Titlul contractului: „{L['titlu']}”", after=8)
    table(doc, ["Nr.", "Nume și prenume", "CNP / seria și nr. CI", "Funcția"],
          [[str(i), None, None, None] for i in range(1, 9)], widths=[1, 6, 5, 5])
    table(doc, ["Nr.", "Autovehicul (marcă, tip)", "Nr. de înmatriculare", "Conducător auto"],
          [[str(i), None, None, None] for i in range(1, 3)], widths=[1, 6, 5, 5])
    P(doc, Y("Se completează doar dacă unitatea o cere prin anunț sau la semnarea contractului."), size=9)
    semnatura(doc)


def genereaza_docx(cheie):
    L = LICITATII[cheie]
    doc = new_doc()
    opis = ["Scrisoare de înaintare și opis",
            "Certificat constatator emis de O.N.R.C. (cu CAEN 4391 autorizat)",
            "Propunerea tehnică",
            "Formularul de ofertă",
            "Anexa la formularul de ofertă",
            "Propunerea financiară: F3 cu prețuri unitare, C6, C7, C8, C9, grafic de execuție fizic și valoric",
            "Declarație privind însușirea modelului de contract",
            "Declarație privind respectarea art. 51 din Legea nr. 98/2016 (mediu, social, relații de muncă)",
            "Declarație privind evitarea conflictului de interese (art. 59–60 din Legea nr. 98/2016)"]
    sectiuni = [propunere_tehnica, formular_oferta, anexa_oferta, decl_contract, decl_art51, decl_art60]
    if cheie == "c4":
        opis += ["Declarație privind neîncadrarea în art. 164, 165 și 167 din Legea nr. 98/2016",
                 "Lista lucrărilor similare executate, cu documente constatatoare",
                 "Lista nominală a personalului și a autovehiculelor pentru acces în cazarmă"]
        sectiuni += [decl_164_165_167, experienta, lista_personal]
    scrisoare(doc, L, opis)
    for s in sectiuni:
        page_break(doc)
        s(doc, L)
    footer(doc, f"{FIRMA['denumire']} – Ofertă „{L['titlu']}”")
    OUT.mkdir(exist_ok=True)
    path = OUT / f"Dosar_oferta_BILZI_{L['fisier']}.docx"
    doc.save(path)
    return path


# ---------------------------------------------------------------- propunere financiară (xlsx)

GALBEN = PatternFill("solid", fgColor="FFFF00")


def genereaza_xlsx(cheie, model):
    """Pornește de la foaia de calcul Onești (formule F3→Analiza→C6–C9→Grafic) și o pregătește
    pentru noua licitație: titluri, valoare estimată, cantități F3 golite (de completat din caietul de sarcini)."""
    L = LICITATII[cheie]
    wb = openpyxl.load_workbook(model)
    titlu_vechi = "Lucrări acoperiș hala producție S.D.F. Onești"
    benef = L["autoritate_scurt"] or "____________"
    for ws in wb:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and not c.value.startswith("="):
                    v = c.value
                    v = v.replace(f"Obiectiv: {titlu_vechi} – Beneficiar: R.N.P. Romsilva, Direcția Silvică Bacău",
                                  f"Obiectiv: {L['titlu']} – Beneficiar: {benef}")
                    v = v.replace(titlu_vechi, L["titlu"])
                    v = v.replace("(Romsilva – DS Bacău, ADV1550112)", f"({benef})")
                    v = v.replace("05.10.2026, ora 09:00", f"{L['termen']}, ora ____")
                    v = v.replace("anunturi.seap@bacau.rosilva.ro", L["email_depunere"] or "____________")
                    v = v.replace("199.506,08", lei(L["valoare"]))
                    c.value = v
    f3 = wb["F3"]
    f3["N48"] = L["valoare"]
    for r in range(7, 38):
        if f3.cell(r, 5).value is not None:
            f3.cell(r, 5).value = None
            f3.cell(r, 5).fill = GALBEN
    f3["C51"] = ("Cantitățile (coloana E) se copiază din F3-ul anexat la caietul de sarcini. Articolele, consumurile din "
                 "Analiza și prețurile se ajustează după F3 – articolele actuale sunt cele de la Onești.")
    f3["C51"].font = Font(bold=True, color="C00000")
    f3["C52"] = "Atenție: verificați dacă F3 conține desfacerea învelitorii existente și evacuarea molozului."
    f3["E43"] = 0.25
    path = OUT / f"Propunere_financiara_BILZI_{L['fisier']}.xlsx"
    wb.save(path)
    return path


if __name__ == "__main__":
    model = Path(sys.argv[1]) if len(sys.argv) > 1 else None
    LICITATII["dofteana"]["distanta"] = "cca. 190 km (Buzău – Focșani – Onești – Dofteana)"
    LICITATII["c4"]["distanta"] = "cca. 115 km (Buzău – București)"
    for k in LICITATII:
        print(genereaza_docx(k))
        if model:
            print(genereaza_xlsx(k, model))
