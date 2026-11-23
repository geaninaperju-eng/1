#!/usr/bin/env python3
"""Dosare de ofertă BILZI STEEL PROFILE S.R.L. – licitații acoperișuri (02.10.2026).

Model: dosarul „Lucrări acoperiș hala producție S.D.F. Onești” (ADV1550112), depus pe 01.10.2026.
Ieșire: output/*.docx și output/*.xlsx

Datele din anunț și din caietul de sarcini (nr. anunț, CPV, oră, durată, garanție, articolele și
cantitățile F3) sunt completate; prețurile și ce rămâne de verificat sunt evidențiate cu galben.
"""
import re
import sys
import zipfile
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


def B(text):
    return (text, {"b": True})


def Y(text):
    """Text propus, de verificat – evidențiat cu galben."""
    return (text, {"y": True})


LICITATII = {
    "dofteana": {
        "fisier": "Parc_Dofteana",
        "titlu": "Înlocuire acoperiș clădire Parc Dofteana – DSBC 2026 II",
        "autoritate": "REGIA NAȚIONALĂ A PĂDURILOR – ROMSILVA prin DIRECȚIA SILVICĂ BACĂU",
        "autoritate_scurt": "Direcția Silvică Bacău",
        "email_depunere": "anunturi.seap@bacau.rosilva.ro",
        "nr_anunt": "ADV1549806",
        "data_anunt": "01.10.2026",
        "termen": "07.10.2026",
        "ora": "10:00",
        "cpv": "45261910-6 – Reparare de acoperișuri",
        "valoare": 252389.69,
        "amplasament": "Clădirea Parc Dofteana, Ocolul Silvic Tg. Ocna, localitatea Dofteana, județul Bacău",
        "procedura": "cumpărare directă",
        "durata": "60 (șaizeci) de zile calendaristice de la ordinul de începere (2 luni, conform caietului de "
                  "sarcini și art. 5.1 din contract), dar nu mai târziu de 31.12.2026",
        "garantie": "24",
        "catalog": True,
        "decizie": "Pădureanu Leonard – Director, Bitir Ioan – Director Tehnic, Tabacaru Ion – Director Economic",
        "experienta": False,
        "acces": False,
        "vizita": False,
        "lucrari": [
            "demontarea învelitorii existente din țiglă ceramică, cu recuperarea materialului, a sistemului pluvial, a "
            "instalației de paratrăsnet, a șorțurilor, capacelor și tubulaturii ceramice a coșurilor, a lambriului de la "
            "intradosul streașinii și a elementelor din lemn degradate (astereală, șipci, contrașipci, pazii), precum și "
            "a firidei electrice în vederea repoziționării;",
            "reparații locale la căpriori și grinzi prin eclisare și completare de secțiune, reparații și consolidări "
            "locale la lucarne, terasă și elementele decorative din lemn;",
            "curățarea vopselei vechi de pe elementele din lemn prin sablare laser, chituire, tratament de profunzime, "
            "impregnare și vopsire cu lazură opacă microporoasă în 2 straturi;",
            "înlocuirea generală a asterelei din scândură de rășinoase de 25 mm, montarea șipcilor 50×50 mm și a "
            "contrașipcilor 30×50 mm, tratate fungicid/insecticid;",
            "montarea foliei anticondens de minimum 120 g/mp și a învelitorii noi din țiglă ceramică Creaton "
            "Melodie/Balance, roșu cupru angobat (sau echivalent), cu accesoriile de fixare, conform NP 069-2014;",
            "coamă ventilată din țiglă ceramică, dolii, șorțuri de streașină, de fronton și de racord la perete (lucarne) "
            "din tablă prevopsită de 0,50 mm și opritori de zăpadă;",
            "lambriu și pazii din lemn de larice termotratat clasa A la intradosul streașinii, prinse cu șuruburi inox;",
            "refacerea instalației de paratrăsnet (I 7-2011, SR EN 62305), cu buletin PRAM, și repoziționarea firidei "
            "electrice, de către electrician autorizat ANRE;",
            "sistem pluvial din tablă prevopsită: jgheab 125 mm, burlan Ø90 mm, colțuri, racorduri, prelungitoare, coturi, "
            "coliere și descărcare liberă la sol;",
            "coșuri de fum (4 buc., cca. 31,90 mp): reparații cu cărămidă plină și mortar CT29, refacerea rosturilor, "
            "hidrofobizare, tubulatură inox AISI 304 Ø180 mm, coronamente și șorțuri de coș din tablă prevopsită;",
            "transportul și manipularea materialelor pe șantier și evacuarea deșeurilor din demontări, colectate selectiv "
            "(OUG 92/2021, HG 856/2002).",
        ],
        "etape": [
            "Etapa 1 – Predarea amplasamentului, constatarea stării elementelor existente, mostre și fișe tehnice pentru "
            "țiglă, accesorii și tinichigerie, supuse aprobării beneficiarului înainte de montaj.",
            "Etapa 2 – Demontări pe tronsoane: învelitoare cu recuperarea țiglei, sistem pluvial, paratrăsnet, șorțuri și "
            "tubulatură coșuri, lambriu, elemente din lemn degradate, firida electrică.",
            "Etapa 3 – Verificarea structurii din lemn după decopertare; reparații prin eclisare, consolidări la lucarne și "
            "terasă (orice situație neprevăzută se comunică beneficiarului înainte de intervenție).",
            "Etapa 4 – Sablare laser, chituire, tratament și vopsire a elementelor din lemn păstrate.",
            "Etapa 5 – Astereală nouă, folie anticondens, șipci și contrașipci, țiglă ceramică, coamă ventilată.",
            "Etapa 6 – Tinichigerie (dolii, șorțuri de streașină, fronton și racord), opritori de zăpadă.",
            "Etapa 7 – Coșuri de fum: reparații zidărie, rosturi, hidrofobizare, tubulatură inox, coronamente și șorțuri.",
            "Etapa 8 – Lambriu și pazii din larice termotratat; sistem pluvial.",
            "Etapa 9 – Paratrăsnet și firida electrică; măsurători PRAM; curățenie, evacuarea deșeurilor, recepție.",
        ],
        "materiale": [
            "țiglă ceramică ", Y("Creaton Melodie/Balance, roșu cupru angobat (sau echivalent)"),
            ", cu piese speciale, coamă, capete de coamă și bandă de coamă ventilată din aceeași gamă; folie anticondens "
            "minimum 120 g/mp; tablă prevopsită de minimum 0,50 mm, finisaj mat, pentru dolii, șorțuri și coronamente; "
            "sistem pluvial din tablă prevopsită (jgheab 125 mm, burlan Ø90 mm); cherestea de rășinoase uscată tehnic și "
            "tratată; lambriu și pazii din larice termotratat clasa A; tubulatură inox AISI 304 Ø180 mm; conductor OL-Zn Ø8 mm "
            "pentru paratrăsnet.",
        ],
        "f3": [
            ("02 Demontări", [
                ("DEM001", "Demontare învelitoare țiglă ceramică existentă, cu recuperarea materialului", "mp", 600),
                ("DEM002", "Demontare sistem pluvial existent (jgheaburi, burlane, cârlige, brățări)", "ml", 175.82),
                ("DEM003", "Demontare instalație paratrăsnet existentă inclusiv console și elemente de fixare", "LS", 1),
                ("DEM004", "Demontare șorțuri, capace și elemente metalice aferente coșurilor de fum", "buc", 4),
                ("DEM005", "Demontare tubulatură ceramică existentă coș fum", "ml", 40),
                ("DEM006", "Demontare lambriu intrados streașină existent", "mp", 155),
                ("DEM007", "Demontare elemente lemn degradate (astereală, șipci, contrașipci, pazii)", "mp", 600),
                ("DEM008", "Demontare firidă electrică și elemente suport existente pentru repoziționare", "buc", 1),
                ("DEM09", "Transportul manual al materialelor, în spații libere și neaccidentate, prin purtat direct pe "
                          "primii 10 m distanță orizontală, cu încărcătura de cel mult 50 kg, la o distanță de cel mult 60 m",
                 "t", 19),
            ]),
            ("03 Reparații suport lemn", [
                ("REP001", "Reparații locale căpriori și grinzi lemn existente prin eclisare și completare secțiune în pod",
                 "mc", 1),
                ("REP002", "Reparații locale lucarne, terasă existentă, elemente decor lemn (include și consolidări "
                           "structură de lemn)", "buc", 4),
                ("REP003", "Curățare strat vechi de vopsea de pe elemente lemn prin sablare laser", "ora", 200),
                ("REP004", "Chituire, tratament de profunzime, impregnare și vopsire lemn exterior vechi după curățare laser",
                 "mp", 50),
            ]),
            ("04 Înlocuire învelitoare", [
                ("ACS001", "Înlocuire generală astereală existentă din scândură rășinoase de 25 mm", "mp", 600),
                ("ACS002", "Montaj șipci 50×50 mm și contrașipci rășinoase 30×50 mm", "mp", 600),
                ("ACS004", "Montaj folie anticondens 120 g/mp", "mp", 600),
                ("ACS005", "Montaj țiglă ceramică inclusiv accesorii de fixare", "mp", 600),
                ("ACS006", "Montaj coamă ventilată din țiglă ceramică", "ml", 92.18),
                ("ACS007", "Montaj dolie din tablă prevopsită", "ml", 26.63),
                ("ACS008", "Montaj șorț streașină din tablă prevopsită", "ml", 110),
                ("ACS009", "Montaj șorț fronton din tablă prevopsită", "ml", 9.80),
                ("ACS010", "Montaj opritori de zăpadă pentru țigla ceramică", "ml", 70),
                ("ACS012", "Montaj șorț de racord la perete (lucarne)", "ml", 25.40),
            ]),
            ("05 Montaj intrados streașină, pazii din lemn", [
                ("ACS001", "Montaj lambriu lemn – larice termotratat, clasa A", "mp", 155),
                ("ACS003", "Montaj pazie lemn – larice termotratat, clasa A", "ml", 105),
            ]),
            ("06 Lucrări la instalații electrice și de protecție la trăsnet", [
                ("ACS012", "Refacere instalație paratrăsnet acoperiș", "LS", 1),
                ("ACS013", "Repoziționare firidă electrică și adaptare suport", "buc", 1),
            ]),
            ("07 Tinichigerie și pluvial", [
                ("TIN001", "Montaj jgheab 125 mm RAL 8019", "ml", 116.01),
                ("TIN002", "Montaj burlan 90 mm / 3 m", "ml", 60),
                ("TIN003", "Montaj colț exterior jgheab 90 grade 125", "buc", 5),
                ("TIN004", "Montaj racord jgheab-burlan 125/90", "buc", 12),
                ("TIN005", "Montaj prelungitor intermediar burlan 90 (lungime produs 1 m / 1,2 m)", "buc", 12),
                ("TIN006", "Montaj coturi burlan 60 grade", "buc", 24),
                ("TIN007", "Montaj cot evacuare burlan", "buc", 12),
                ("TIN008", "Montaj coliere burlan 90", "buc", 24),
                ("TIN009", "Descărcare ape pluviale liber la sol", "buc", 12),
            ]),
            ("08 Coșuri de fum", [
                ("COS000", "Reparații și curățare exterioară coșuri de fum existente", "mp", 31.90),
                ("COS001", "Conservare finisaj original coș fum, refacere rosturi și hidrofobizare", "mp", 31.90),
                ("COS002", "Montaj tubulatură coș fum inox", "buc", 4),
                ("COS003", "Confecționare coronament coș din tablă 0,50 mm RAL 7016, ieșire 40 cm față de coș", "buc", 4),
                ("COS004", "Montaj șorțuri coș – set complet față, laterale, spate", "buc", 4),
            ]),
        ],
        "note_f3": "Sursa: F3 din documentația de atribuire (Deviz nr. 8076/1/22.07.2026, devizele 02–08). F3-ul "
                   "original are și resursele pe fiecare articol (material, manoperă, utilaj, transport); C6–C9 se "
                   "completează din aceleași liste. Culorile din F3 (jgheab RAL 8019, șorț streașină RAL 8017, "
                   "coronament RAL 7016) diferă de caietul de sarcini (roșu cupru) – de clarificat.",
    },
    "c4": {
        "fisier": "Cazarma_3589_Pav_C4",
        "titlu": "Lucrări de reparații acoperiș pav. C4, cazarma 3589 București",
        "autoritate": "INSTITUTUL NAȚIONAL DE CERCETARE-DEZVOLTARE MEDICO-MILITARĂ „CANTACUZINO”",
        "autoritate_scurt": "INCDMM „Cantacuzino”",
        "email_depunere": "office.cantacuzino@mapn.ro",
        "nr_anunt": "ADV1550471",
        "data_anunt": "30.09.2026",
        "termen": "05.10.2026",
        "ora": "14:00",
        "cpv": "45453000-7 – Lucrări de reparații generale și de renovare",
        "valoare": 77677.10,
        "amplasament": "Pavilionul C4 din cazarma 3589, INCDMM „Cantacuzino”, Splaiul Independenței nr. 103, sector 5, "
                       "București",
        "procedura": "cumpărare directă",
        "durata": "45 (patruzeci și cinci) de zile calendaristice de la data emiterii ordinului de începere",
        "garantie": "36",
        "catalog": True,
        "decizie": Y("de completat – anunțul nu le nominalizează; documentația este aprobată de Col. medic "
                     "Conf. Univ. Dr. Cătălin-Gabriel Smarandache (comandant) și avizată de Col. Mihai-Andrei "
                     "Părăușanu (director administrativ) – cereți lista la bap@cantacuzino.ro"),
        "experienta": True,
        "acces": True,
        "vizita": True,
        "lucrari": [
            "desfacerea învelitorii din olane / țigle solzi sau profilate pe șipci, inclusiv desfacerea șipcilor doliilor "
            "(480 mp), cu sortarea materialelor recuperabile;",
            "refacerea învelitorii din țiglă solzi sau olane pe șipci de lemn noi, tratate, cu coame, inclusiv doliile, "
            "paziile și șorțurile (480 mp): se refolosesc țiglele recuperate în stare bună, iar cele sparte sau degradate se "
            "înlocuiesc cu țiglă nouă compatibilă ca model și culoare (ofertat: cca. 15% din suprafață, adică cca. 2.600 "
            "buc.; procentul se confirmă la vizită), cu verificarea și înlocuirea foliei anticondens unde este cazul;",
            "reparații la învelitoarea din țigle profilate, cu țigle și coame în mortar de ciment, la acoperiș fără "
            "astereală (65 mp);",
            "jgheaburi din tablă zincată 0,5 mm, semirotunde D=12,5 cm, executate pe șantier, cu colțuri, capace și ștuț "
            "de racord (70 m);",
            "panouri parazăpezi din metal cu ancore prinse în căpriori (70 buc.);",
            "etanșarea / fixarea protecțiilor din tablă la îmbinări cu inele din bandă de aluminiu 1×20 mm (60 m);",
            "plasă metalică de protecție pe cadru de oțel, cu ușă din plasă (64 mp);",
            "manipularea materialelor cu macaraua (35 t) și schelă metalică de susținere 3–6 m la parter (64 mp);",
            "gestionarea deșeurilor rezultate (depozitare, sortare și evacuare).",
        ],
        "etape": [
            "Etapa 1 – Vizita la amplasament (obligatorie înainte de ofertare), predarea amplasamentului, aprobarea listei "
            "nominale de acces în cazarmă.",
            "Etapa 2 – Montarea schelei, a sistemelor de siguranță și a jgheaburilor de evacuare a molozului.",
            "Etapa 3 – Desfacerea tinichigeriei și a învelitorii, de la coamă spre streașină, pe fâșii; sortarea țiglelor "
            "recuperabile; coborârea molozului în siguranță (fără aruncare de pe acoperiș).",
            "Etapa 4 – Verificarea și repararea foliei anticondens, șipci noi, montarea țiglelor de la streașină spre coamă, "
            "cu fixare mecanică perimetrală; dolii din tablă zincată de minimum 0,5 mm.",
            "Etapa 5 – Reparațiile la zona fără astereală: înlocuirea țiglelor profilate rupte, coame în mortar M100.",
            "Etapa 6 – Jgheaburi semirotunde D=12,5 cm cu pantă de 2–5 mm/m, cârlige la 60–80 cm, îmbinări cositorite.",
            "Etapa 7 – Parazăpezi pe rândul 2–3 de țigle, ancorate în căpriori; inele de etanșare din bandă Al.",
            "Etapa 8 – Plasa metalică de protecție pe cadru de oțel, cu ușă; curățenie, evacuarea deșeurilor, recepție.",
        ],
        "materiale": [
            "țiglă solzi / olane ", Y("compatibile cu învelitoarea existentă (model și culoare de stabilit la vizită)"),
            ", coame, folie anticondens, șipci de rășinoase tratate; tablă zincată la cald de 0,5 mm pentru jgheaburi și "
            "dolii, aliaj Staniu-Plumb LP30 pentru lipituri; parazăpezi din oțel galvanizat la cald; bandă de aluminiu "
            "1×20 mm; mortar de ciment-var M100; plasă metalică și profile din oțel pentru cadru.",
        ],
        "f3": [
            ("Pavilion C4 – antemăsurătoare (Anexa 1)", [
                ("RPCT26B1", "Desfacerea învelitorilor din olane, țigle solzi sau profilate pe șipci, incl. desfacerea "
                             "șipcilor doliilor", "mp", 480),
                ("RPCI01C", "Învelitori din țiglă solzi sau olane, coame așezate pe șipci de lemn, inclusiv doliile, "
                            "paziile, șorțurile", "mp", 480),
                ("RPCI05XB", "Reparații la învelitori de țigle profilate cu țigle și coame din mortar de ciment la "
                             "acoperiș fără astereală", "mp", 65),
                ("RCSI18A", "Jgheaburi tablă zincată 0,5 mm, semirotunde, D=12,5 cm, exec. pe șantier, incl. "
                            "colțuri/capace/ștuț racord", "m", 70),
                ("RPDE13A", "Montarea panourilor de parazăpezi din metal, cu ancore", "buc", 70),
                ("RPIZD31B", "Etanșarea sau fixarea protecțiilor din tablă la îmbinări sau a izolației cu inele din "
                             "bandă Al 1×20 mm", "m", 60),
                ("W1C11B1", "Plasă metalică de protecție pe cadru de oțel, ușă din plasă – montare", "mp", 64),
                ("TRB22E6B", "Manipulat materiale și elemente prefabricate cu macara turn, greutatea sarcinii la "
                             "fiecare transport = 0,5–1 t", "tonă", 35),
                ("CB41B1", "Susțineri din schelă metalică, sarcina 1000 daN/mp, cu înălțimea de 3–6 m, la parter",
                 "mp", 64),
            ]),
        ],
        "note_f3": "Sursa: Antemăsurătoarea pav. C4 (Anexa 1, întocmit ing. Iulia Păunescu) din documentația anunțului "
                   "ADV1550471. Specificațiile tehnice (Anexa 2) cer: termen de execuție 45 de zile calendaristice, "
                   "garanție 36 de luni, vizitarea obligatorie a amplasamentului.",
    },
}


# ---------------------------------------------------------------- licitații noi 12.11.2026
CFR_GL_DECIZIE = ("Marian Profir – Director; Cristian Calinov – Contabil Șef; Adrian Bocșan – Șef Divizie Exploatare "
                  "Tehnic; Marian Ilie – Consilier juridic; Nicoleta Iorga – Șef Serviciu O.T.D. și IT; Lucian Jacota – "
                  "Șef Serviciu Aprovizionare, Achiziții și Administrativ; Grigore Dan Ungureanu – Expert Achiziții "
                  "Publice; Stelian Băncilă – Economist Serviciul Aprovizionare, Achiziții și Administrativ")
CFR_CT_DECIZIE = ("Niculae Auraș Gavrilă – Director; Florentina Sima – Contabil Șef; Ionuț Popa – Șef Divizie Exploatare "
                  "Comercială și Marketing; Hristu Aurelian Tanef – Șef Serviciu AAA; Marinela Olteanu – Șef Serviciu OTD "
                  "și IT; Cristina Avisalon – Șef Serviciu Financiar; Violeta Mihaela Dragomir – Economist Serviciul AAA; "
                  "Paul Țânțar – Inginer Serviciul AAA; Valentin Roman – Inginer Serviciul OTD și IT; Mirela Georgescu – "
                  "Consilier Juridic; Florentina Țurcanu – Consilier Juridic")

LICITATII.update({
    "romeo": {
        "fisier": "ADV1556879_Cabana_Romeo",
        "titlu": "Înlocuire învelitoare acoperiș Cabana de vânătoare Romeo – DSBC2026",
        "autoritate": "REGIA NAȚIONALĂ A PĂDURILOR – ROMSILVA prin DIRECȚIA SILVICĂ BACĂU",
        "autoritate_scurt": "Direcția Silvică Bacău (O.S. Tg. Ocna)",
        "email_depunere": "anunturi.seap@bacau.rosilva.ro",
        "nr_anunt": "ADV1556879",
        "data_anunt": "10.11.2026",
        "termen": "16.11.2026",
        "ora": "08:00",
        "cpv": "45261910-6 – Reparare de acoperișuri",
        "valoare": 51539.86,
        "amplasament": "Cabana de vânătoare Romeo, fond forestier, Ocolul Silvic Tg. Ocna, județul Bacău (coordonate "
                       "Stereo 70 din caietul de sarcini; acces foarte greu – ultimii 7,5 km pe drum forestier)",
        "procedura": "cumpărare directă",
        "durata": "30 (treizeci) de zile calendaristice de la ordinul de începere, dar nu mai târziu de 31.12.2026",
        "garantie": "24",
        "catalog": True,
        "decizie": "Pădureanu Leonard – Director, Bitir Ioan – Director Tehnic, Tabacaru Ion – Director Economic",
        "experienta": False, "acces": False, "vizita": False,
        "distanta": "cca. 190 km (Buzău – Focșani – Onești – Tg. Ocna) + 7,5 km drum forestier",
        "saptamani": 5,
        "durata_scurt": "30 zile calendaristice (5 săptămâni)",
        "lucrari": [
            "desfacerea învelitorii existente din șindrilă bituminoasă (256,50 mp), cu sortarea și evacuarea deșeurilor;",
            "astereală din OSB 10 mm în zonele degradate (50 mp);",
            "rigle din lemn de rășinoase tratate pe căpriori (256,50 mp) și folie anticondens (307,80 mp);",
            "învelitoare din țiglă metalică BILKA Balcanic mat 0,50 mm, maro închis RAL 8019 (256,50 mp), prinsă cu "
            "holzșuruburi cu garnitură EPDM;",
            "coame mari BILKA RAL 8019 cu bandă de coamă și capace (13,50 ml), dolii (13,50 ml), borduri de streașină "
            "(27 ml) și de fronton (38 ml), opritori de zăpadă (36 ml);",
            "sistem pluvial BILKA RAL 8019: jgheab semicircular D125 (13,50 m), burlane D90 (12 m), cârlige, racorduri, "
            "capace, coturi, coliere, prelungitoare;",
            "transportul materialelor până la cabană și evacuarea deșeurilor.",
        ],
        "etape": [
            "Etapa 1 – Predarea amplasamentului, transportul materialelor pe drumul forestier, organizarea de șantier.",
            "Etapa 2 – Desfacerea șindrilei bituminoase pe tronsoane, verificarea astereli și a căpriorilor.",
            "Etapa 3 – Înlocuirea astereli degradate cu OSB 10 mm, folie anticondens, rigle tratate.",
            "Etapa 4 – Montajul țiglei metalice de la streașină spre coamă, cu holzșuruburi cu garnitură EPDM.",
            "Etapa 5 – Coame, dolii, borduri, opritori de zăpadă; etanșări cu mastic.",
            "Etapa 6 – Sistem pluvial; curățenie, evacuarea deșeurilor, recepția la terminarea lucrărilor.",
        ],
        "materiale": [
            "țiglă metalică ", Y("BILKA Balcanic mat 0,50 mm RAL 8019 (ca în F3)"),
            ", coame, capace, dolii, borduri și opritori din aceeași gamă; folie anticondens; OSB 10 mm; rigle de "
            "rășinoase tratate; sistem pluvial BILKA D125/D90 RAL 8019; holzșuruburi cu garnitură EPDM; mastic.",
        ],
        "f3": [
            ("Deviz DMH937 – Înlocuire învelitoare acoperiș Cabana Romeo", [
                ("RPCT26D1", "Desfacerea învelitorilor din șindrilă bituminoasă – Tegola", "mp", 256.5),
                ("CE18A1", "asim – Astereală executată din OSB gr. = 10 mm", "mp", 50),
                ("3270480", "OSB 10 mm", "mp", 51.5),
                ("CE20A1", "Rigle din lemn de rășinoase așezate pe căpriori", "mp", 256.5),
                ("CE05A1", "Învelitori din tablă BILKA Balcanic mat 0,5 mm maro închis RAL 8019 – montaj", "mp", 256.5),
                ("3270481", "Țiglă metalică BILKA Balcanic mat (RAL 8019), 0,5 mm", "mp", 307.8),
                ("7299982", "Folie anticondens", "mp", 307.8),
                ("7299995", "Holzșurub", "buc", 2565),
                ("7299992", "Tablă plană", "buc", 2),
                ("7299996", "Opritor zăpadă", "ml", 36),
                ("CE08A%9", "asim – Montat coame pentru învelitori din tablă BILKA RAL 8019", "ml", 13.5),
                ("3270482", "Coamă mare BILKA maro închis mat (RAL 8019) 0,5 mm", "buc", 7),
                ("3270483", "Capac coamă BILKA maro mat (RAL 8019) 0,5 mm", "buc", 2),
                ("7299994", "Bandă pentru coamă", "buc", 3),
                ("CE08A%6", "asim – Montat bordură streașină", "ml", 27),
                ("7299997", "Bordură streașină", "ml", 27),
                ("CE08A%7", "asim – Montat bordură fronton BILKA RAL 8019", "ml", 38),
                ("3270468", "Bordură fronton BILKA maro închis mat (RAL 8019) 0,5 mm", "buc", 19),
                ("CE08A%8", "asim – Montat dolii pentru învelitori din tablă BILKA RAL 8019", "ml", 13.5),
                ("3270484", "Dolie BILKA maro închis mat (RAL 8019) 0,5 mm EN 10143", "buc", 7),
                ("7299993", "Mastic premium", "buc", 3),
                ("CE13A1", "asim – Montat jgheaburi din tablă BILKA RAL 8019", "m", 13.5),
                ("3270485", "Jgheab de scurgere metalic semicircular D 125 mm BILKA RAL 8019", "m", 14),
                ("7299989", "Cârlig jgheab", "buc", 19),
                ("7299987", "Elemente îmbinare jgheab", "buc", 4),
                ("7299986", "Racord jgheab/burlan", "buc", 3),
                ("3270487", "Capac jgheab BILKA 125 mm RAL 8019", "buc", 2),
                ("CE14A1", "asim – Montat burlane din tablă BILKA RAL 8019", "m", 12),
                ("3270486", "Burlan de scurgere metalic circular BILKA RAL 8019 D 90 mm", "m", 12),
                ("7299988", "Colier burlan", "buc", 9),
                ("7299984", "Prelungitor intermediar 1 m", "buc", 3),
                ("7299985", "Cot de 60 grade", "buc", 9),
                ("TRA02A40", "Transportul rutier al materialelor, semifabricatelor cu autocamionul pe dist. = 40 km",
                 "tonă", 10),
            ]),
        ],
        "note_f3": "Sursa: Formularul F3 – Deviz ofertă DMH937 (O.S. Tg. Ocna) din documentația anunțului ADV1556879, "
                   "33 de poziții, cu aceleași coduri și cantități. Caietul de sarcini: termen 30 de zile, garanție 24 "
                   "de luni, vizitarea amplasamentului recomandată (acces foarte greu).",
    },
    "loto": {
        "fisier": "ADV1556660_Loto_Bacau_terasa",
        "titlu": "Lucrări de reparații hidroizolație la terasa agenției 04-013 din cadrul SZ Galați – PLJ Bacău",
        "autoritate": "COMPANIA NAȚIONALĂ LOTERIA ROMÂNĂ S.A.",
        "autoritate_scurt": "C.N. Loteria Română S.A.",
        "email_depunere": "achizitii.publice@loto.ro",
        "nr_anunt": "ADV1556660",
        "data_anunt": "11.11.2026",
        "termen": "18.11.2026",
        "ora": "12:00",
        "cpv": "45261310-0 – Lucrări de hidroizolare",
        "valoare": 44235.18,
        "amplasament": "Agenția Loto 04-013, str. Logofăt Tăutu nr. 1C, Bacău",
        "procedura": "cumpărare directă",
        "durata": Y("15 (cincisprezece) zile calendaristice de la ordinul de începere"),
        "garantie": "36",
        "catalog": False,
        "decizie": Y("de completat – anunțul nu le nominalizează; cereți lista la achizitii.publice@loto.ro"),
        "experienta": False, "acces": False, "vizita": False,
        "distanta": "cca. 180 km (Buzău – Focșani – Bacău)",
        "saptamani": 3,
        "durata_scurt": "15 zile calendaristice (3 săptămâni)",
        "lucrari": [
            "desfacerea straturilor friabile ale hidroizolației existente și a învelitorii aticurilor, pregătirea "
            "suprafeței (95 mp);",
            "reparații locale cu membrană bituminoasă V3 după desfacere (95 mp) și șapă de pantă pentru corecția "
            "pantelor spre gura de scurgere (95 mp);",
            "amorsă bituminoasă (95 mp), membrană P4 termosudată – strat 1 (95 mp) și membrană PA4,5 cu ardezie – "
            "strat 2 (95 mp), cu racordări la atice;",
            "gură de scurgere Ø90–110 mm (1 buc.) și căciuli de ventilație (2 buc.);",
            "glafuri din tablă la atice, cu profil picurător (24 ml);",
            "ridicarea materialelor pe terasă, evacuarea molozului și a deșeurilor (0,2 t).",
        ],
        "etape": [
            "Etapa 1 – Predarea amplasamentului; protejarea agenției (activitatea continuă).",
            "Etapa 2 – Desfacerea straturilor friabile și a glafurilor, curățarea suportului.",
            "Etapa 3 – Șapă de pantă, reparații locale cu V3, uscarea suportului.",
            "Etapa 4 – Amorsă, membrană P4, membrană PA4,5 cu ardezie, racordări la atice și guri.",
            "Etapa 5 – Gura de scurgere, căciulile de ventilație, glafurile de atic; proba de etanșeitate, recepție.",
        ],
        "materiale": [
            "membrane bituminoase termosudabile ", Y("V3, P4 și PA4,5 cu ardezie (Sika Festa/Ecobit sau echivalent)"),
            ", amorsă bituminoasă cu solvent; șapă de ciment pentru pante; gură de scurgere Ø110 cu manșon bituminos; "
            "căciuli de ventilație; tablă prevopsită 0,50 mm pentru glafuri cu picurător.",
        ],
        "f3": [
            ("Deviz estimativ – terasa agenției 04-013 Bacău", [
                ("1", "Desfacere straturi friabile – pregătire suprafață", "mp", 95),
                ("2", "Montare membrană bituminoasă V3 kg/mp – strat 1 reparații locale și completări după desfacere",
                 "mp", 95),
                ("3", "Turnare șapă de pantă – corecție pante", "mp", 95),
                ("4", "Aplicare amorsă bituminoasă", "mp", 95),
                ("5", "Montare membrană bituminoasă P4 kg/mp – strat 1", "mp", 95),
                ("6", "Montare membrană bituminoasă PA4,5 kg/mp – strat 2", "mp", 95),
                ("7", "Montare guri de scurgere diametru 90–110 mm", "buc", 1),
                ("8", "Montare căciulă ventilație 90–110 mm", "buc", 2),
                ("9", "Montare glafuri din tablă la atice (profil picurător)", "ml", 24),
                ("10", "Material mărunt", "buc", 1),
                ("11", "Ridicare materiale și accesorii pe acoperiș", "buc", 1),
                ("12", "Evacuare gunoi, moloz, deșeuri", "tone", 0.2),
            ]),
        ],
        "note_f3": "Sursa: „Deviz estimativ lucrare AG 04-013 – PLJ Bacău” din anunțul ADV1556660 (12 poziții, fără "
                   "prețuri). Vizionarea amplasamentului: zile lucrătoare 12–15, dl. Cosmin Năstac, 0750.622.503. "
                   "Achizitorul nu acordă avans; prețul este ferm.",
    },
    "cfr_galati": {
        "fisier": "SCN1181822_CFR_Galati_hidroizolatie",
        "titlu": "Reparație hidroizolație fostă clădire centrală termică Revizia Vagoane Galați",
        "autoritate": "SNTFC „CFR CĂLĂTORI” S.A. – Sucursala Regională de Transport Feroviar de Călători Galați",
        "autoritate_scurt": "S.R.T.F.C. Galați",
        "email_depunere": "depunere online în SEAP (procedura simplificată SCN1181822)",
        "nr_anunt": "SCN1181822",
        "data_anunt": "05.11.2026",
        "termen": "23.11.2026",
        "ora": "15:00",
        "cpv": "45261310-0 – Lucrări de hidroizolare",
        "valoare": 104904.00,
        "amplasament": "Revizia Vagoane Galați, str. Alexandru Davila nr. 1, Galați – fosta clădire a centralei "
                       "termice (terasă 459 mp, cu 114 panouri fotovoltaice de 455 W)",
        "procedura": "procedură simplificată (Legea nr. 99/2016)",
        "sectorial": True,
        "durata": "60 (șaizeci) de zile de la ordinul de începere (durata contractului: 5 luni)",
        "garantie": Y("36"),
        "catalog": False,
        "decizie": CFR_GL_DECIZIE,
        "experienta": False, "acces": False, "vizita": False,
        "distanta": "cca. 130 km (Buzău – Brăila – Galați)",
        "saptamani": 9,
        "durata_scurt": "60 zile (9 săptămâni)",
        "lucrari": [
            "demontarea celor 114 panouri fotovoltaice de pe acoperiș, cu depozitarea lor în siguranță;",
            "decopertarea membranei afectate (459 mp), reparații cu mortar rapid (4 mp) și la golurile din tavan (2 mp);",
            "aplicarea amorsei bituminoase cu solvent (459 mp) și a hidroizolației din membrană termosudabilă cu ardezie "
            "(459 mp), cu racordări la atice și la gurile de scurgere;",
            "remontarea celor 114 panouri fotovoltaice, inclusiv setarea și conexiunea pentru funcționare (electrician "
            "autorizat ANRE);",
            "utilizarea macaralei (4 ore) și transportul deșeurilor la depozit autorizat (2,5 t).",
        ],
        "etape": [
            "Etapa 1 – Predarea amplasamentului; deconectarea sistemului fotovoltaic de către electrician autorizat.",
            "Etapa 2 – Demontarea și coborârea panourilor (macara), depozitare pe paleți, inventariere.",
            "Etapa 3 – Decopertarea membranei afectate, curățare, reparații cu mortar rapid și la tavan.",
            "Etapa 4 – Amorsă, membrană cu ardezie termosudată la peste +5 °C, fără ploaie; racordări.",
            "Etapa 5 – Remontarea structurii și a panourilor, reconectare, punere în funcțiune, verificări.",
            "Etapa 6 – Evacuarea deșeurilor, curățenie, recepție.",
        ],
        "materiale": [
            "membrană bituminoasă termosudabilă cu armătură din poliester și ardezie ",
            Y("min. 4 mm, flexibilitate la rece ≤ -10 °C, stabilitate la cald ≥ 120 °C (Sika Ecobit PA4,5 sau echivalent)"),
            "; amorsă bituminoasă cu solvent; mortar rapid de reparații; accesorii de fixare și conectare a panourilor.",
        ],
        "f3": [
            ("Lista cantităților de lucrări – hidroizolație fostă CT Revizia Vagoane", [
                ("1", "Demontare panouri fotovoltaice de pe acoperiș", "buc", 114),
                ("2", "Decopertare membrană afectată", "mp", 459),
                ("3", "Efectuare reparații cu mortar rapid", "mp", 4),
                ("4", "Reparații goluri tavan", "mp", 2),
                ("5", "Aplicare amorsă cu solvent", "mp", 459),
                ("6", "Hidroizolație cu ardezie", "mp", 459),
                ("7", "Montare panouri fotovoltaice pe acoperiș, inclusiv setare și conexiune pentru funcționare",
                 "buc", 114),
                ("8", "Utilizare macara", "ore", 4),
                ("9", "Transportul și predarea la depozitul deșeurilor", "tone", 2.5),
            ]),
        ],
        "note_f3": "Sursa: Lista cantităților de lucrări din caietul de sarcini GL6/2/865/12.05.2026 (S.R.T.F.C. Galați), "
                   "anunț SCN1181822. Garanție de participare 1.049 lei (3 luni), garanție de bună execuție 5%. "
                   "Fără cerințe de experiență, cifră de afaceri sau ISO.",
    },
    "cfr_mangalia": {
        "fisier": "SCN1181778_CFR_Mangalia_Vila_C",
        "titlu": "Reparație acoperiș clădire Dormitor personal CF – Stația Mangalia (Vila C)",
        "autoritate": "SNTFC „CFR CĂLĂTORI” S.A. – Sucursala Regională de Transport Feroviar de Călători Constanța",
        "autoritate_scurt": "S.R.T.F.C. Constanța",
        "email_depunere": "depunere online în SEAP (procedura simplificată SCN1181778)",
        "nr_anunt": "SCN1181778",
        "data_anunt": "03.11.2026",
        "termen": "20.11.2026",
        "ora": "15:00",
        "cpv": "45261900-3 – Lucrări de reparare și de întreținere a acoperișurilor",
        "valoare": 172000.00,
        "amplasament": "Clădirea Dormitor personal CF – Stația Mangalia (Vila C), nr. inventar C08/Y801/803117, "
                       "Șos. Constanței nr. 86, Mangalia, județul Constanța (P+1E, acoperiș cca. 236 mp)",
        "procedura": "procedură simplificată (Legea nr. 99/2016)",
        "sectorial": True,
        "durata": "60 (șaizeci) de zile de la predarea amplasamentului (durata contractului: 5 luni)",
        "garantie": "24",
        "catalog": False,
        "decizie": CFR_CT_DECIZIE,
        "experienta": False, "acces": False, "vizita": True,
        "distanta": "cca. 230 km (Buzău – Slobozia – Constanța – Mangalia)",
        "saptamani": 9,
        "durata_scurt": "60 zile (9 săptămâni)",
        "lucrari": [
            "demontarea instalației de panouri solare, a învelitorii din tablă profilată, a jgheaburilor, burlanelor și "
            "a astereli;",
            "astereală nouă din scândură de rășinoase tratată, termoizolație din saltele de vată minerală, folie "
            "anticondens, șipci și contrașipci (spațiu de ventilare 3–5 cm);",
            "învelitoare din țiglă metalică din oțel zincat (min. 275 g/mp Zn) și prevopsit, grosime min. 0,50 mm, cu "
            "garanție de producător de min. 25 de ani, prinsă cu șuruburi autoforante cu garnitură EPDM;",
            "coame, dolii, borduri, jgheaburi și burlane din același sistem integrat;",
            "remontarea panourilor solare existente / pregătirea pentru instalația fotovoltaică;",
            "evacuarea deșeurilor la depozit autorizat.",
        ],
        "etape": [
            "Etapa 0 – Vizitarea obligatorie a amplasamentului (solicitare la marinela.olteanu@cfrcalatori.ro) și "
            "antemăsurătoarea proprie, pe care se bazează oferta.",
            "Etapa 1 – Predarea amplasamentului, convenția SSM/PSI, organizarea de șantier.",
            "Etapa 2 – Demontarea panourilor solare, a pluvialului, a tablei și a astereli, pe tronsoane.",
            "Etapa 3 – Astereală, vată minerală, folie anticondens, șipci și contrașipci.",
            "Etapa 4 – Țiglă metalică, coame, dolii, borduri; sistem pluvial.",
            "Etapa 5 – Remontarea panourilor solare, curățenie, evacuarea deșeurilor, recepție.",
        ],
        "materiale": [
            "țiglă metalică ", Y("BILKA mat 0,50 mm, Z275, garanție 25+ ani (sau echivalent)"),
            ", cu coame, dolii, borduri și sistem pluvial BILKA din același sistem; scândură de rășinoase tratată; vată "
            "minerală 10 cm; folie anticondens; șipci și contrașipci; șuruburi autoforante cu garnitură EPDM.",
        ],
        "f3": [
            ("Antemăsurătoare estimată (de confirmat la vizita obligatorie)", [
                ("D01", "Demontare instalație panouri solare existentă, cu depozitare pentru remontare", "set", 1),
                ("D02", "Desfacere învelitoare din tablă profilată, inclusiv tinichigerie", "mp", 236),
                ("D03", "Demontare jgheaburi și burlane", "ml", 70),
                ("D04", "Desfacere astereală degradată", "mp", 236),
                ("A01", "Astereală din scândură de rășinoase 25 mm, tratată", "mp", 236),
                ("A02", "Termoizolație cu saltele de vată minerală 10 cm", "mp", 236),
                ("A03", "Folie anticondens", "mp", 236),
                ("A04", "Șipci și contrașipci de rășinoase tratate", "mp", 236),
                ("A05", "Învelitoare din țiglă metalică 0,50 mm Z275, cu accesorii de fixare", "mp", 236),
                ("A06", "Coame din tablă, cu bandă de ventilare și capace", "ml", 30),
                ("A07", "Dolii din tablă", "ml", 12),
                ("A08", "Borduri de streașină și fronton", "ml", 64),
                ("P01", "Jgheaburi D125 cu cârlige și accesorii", "ml", 40),
                ("P02", "Burlane D90 cu coliere și coturi", "ml", 30),
                ("S01", "Remontare instalație panouri solare existentă", "set", 1),
                ("E01", "Evacuarea deșeurilor la depozit autorizat (container)", "buc", 2),
            ]),
        ],
        "note_f3": "Caietul de sarcini CT6/3/1807/30.10.2026 NU conține listă de cantități: oferta se face pe baza "
                   "antemăsurătorii proprii, după vizitarea OBLIGATORIE a amplasamentului (cca. 236 mp de acoperiș). "
                   "Cantitățile de mai sus sunt estimate și trebuie confirmate la vizită. Garanție de participare 1.720 "
                   "lei (90 de zile). Fără cerințe de experiență, cifră de afaceri sau ISO.",
    },
})

# ---------------------------------------------------------------- licitații noi 13.11.2026
CFR_BV_DECIZIE = ("Răzvan Cîrjan – Director; Radu Caleș – Contabil Șef; Valentin Bulmaci – Șef Serviciu O.T.D. și IT; "
                  "Paula Caleș – Șef Serviciu Achiziții și Administrativ")
SALUBRIS_DECIZIE = ("ec. Cătălin Neculău – Director General; ec. Virginica Pipirig – Contabil Șef; ec. Alina Pantea – "
                    "Control Financiar Preventiv; cons. jr. Marinel Anușca – Serviciul Juridic; ec. Oana Bularda – "
                    "Serviciul Achiziții Aprovizionare-Pază; Carmen Damian – Serviciul Tehnic")
LICITATII.update({
    "cfr_sibiu": {
        "fisier": "SCN1182052_CFR_Sibiu_baraca",
        "titlu": "Reparații acoperiș baracă metalică la Depoul Sibiu – S.R.T.F.C. Brașov",
        "autoritate": "SNTFC „CFR CĂLĂTORI” S.A. – Sucursala Regională de Transport Feroviar de Călători Brașov",
        "autoritate_scurt": "S.R.T.F.C. Brașov",
        "email_depunere": "depunere online în SEAP (procedura simplificată SCN1182052)",
        "nr_anunt": "SCN1182052",
        "data_anunt": "12.11.2026",
        "termen": "02.12.2026",
        "ora": "15:00",
        "cpv": "45261910-6 – Reparare de acoperișuri",
        "valoare": 60000.00,
        "amplasament": "Depoul Sibiu – hala procesului tehnologic (baraca metalică), str. Dorobanților nr. 22, Sibiu "
                       "(o apă a acoperișului în două ape, cca. 350 mp, plăci de azbociment din 1991)",
        "procedura": "procedură simplificată (Legea nr. 99/2016)",
        "sectorial": True,
        "durata": "30 (treizeci) de zile calendaristice de la predarea amplasamentului",
        "garantie": Y("36"),
        "catalog": False,
        "decizie": CFR_BV_DECIZIE,
        "experienta": False, "acces": False, "vizita": True,
        "distanta": "cca. 290 km (Buzău – Brașov – Făgăraș – Sibiu)",
        "saptamani": 5,
        "durata_scurt": "30 zile calendaristice (5 săptămâni)",
        "lucrari": [
            "organizarea și delimitarea zonei de lucru, plan de lucru pentru îndepărtarea azbestului notificat la "
            "Inspectoratul Teritorial de Muncă Sibiu (HG nr. 1875/2005);",
            "demontarea învelitorii din plăci de azbociment (350 mp) fără spargere, cu umectarea/încapsularea fibrelor, "
            "manipularea, colectarea și ambalarea etanșă în folie PE și saci big-bag marcați;",
            "transportul și predarea deșeurilor cu azbest (cod 17 06 05*) la operator autorizat, cu documente de "
            "trasabilitate (formular de încărcare–descărcare deșeuri periculoase);",
            "verificarea și pregătirea structurii metalice (pane, contravântuiri), remedierea locală a elementelor "
            "deteriorate și protecție anticorozivă;",
            "învelitoare nouă din tablă cutată prevopsită grosime 0,50 mm (350 mp), fixată cu șuruburi autoforante cu "
            "garnitură EPDM, cu suprapuneri etanșate;",
            "elemente de tinichigerie (coamă/racord cu apa existentă, borduri, închideri), profile de etanșare și "
            "elemente pentru evacuarea apelor pluviale;",
            "curățarea amplasamentului și evacuarea tuturor deșeurilor rezultate.",
        ],
        "etape": [
            "Etapa 0 – Vizionarea obligatorie a obiectivului (str. Dorobanților nr. 22, Sibiu) înainte de depunerea "
            "ofertei; antemăsurătoarea proprie.",
            "Etapa 1 – Predarea amplasamentului; plan de lucru azbest notificat la ITM; delimitare, semnalizare, linii "
            "de ancorare.",
            "Etapa 2 – Umectare, demontarea plăcilor de azbociment pe tronsoane, coborâre fără aruncare, ambalare etanșă.",
            "Etapa 3 – Verificarea structurii metalice, remedieri locale, protecție anticorozivă.",
            "Etapa 4 – Montarea tablei cutate pe tronsonul descoperit în aceeași zi; etanșări și tinichigerie.",
            "Etapa 5 – Sistem pluvial, verificare vizuală și tehnică, predarea deșeurilor la operator autorizat, "
            "curățenie, recepție.",
        ],
        "materiale": [
            "tablă cutată prevopsită ", Y("BILKA T35 lucios 0,50 mm, culoare la alegerea beneficiarului (sau echivalent)"),
            ", cu coame și accesorii de tinichigerie din același material; șuruburi autoforante pentru oțel cu "
            "garnitură EPDM; profile de etanșare din spumă și bandă butilică; jgheaburi și burlane BILKA; folie PE "
            "0,2 mm și saci big-bag pentru azbest; soluție de încapsulare a fibrelor; vopsea anticorozivă.",
        ],
        "f3": [
            ("Lista de cantități – Anexa la CS ETT54/3/1/318/2026", [
                ("1", "Organizarea și pregătirea zonei de lucru", "lucrare", 1),
                ("2", "Demontarea învelitorii existente din plăci de azbociment", "mp", 350),
                ("3", "Manipulare, colectare și ambalare plăci de azbociment", "mp", 350),
                ("4", "Transport și predare deșeuri din azbociment la operator autorizat", "lucrare", 1),
                ("5", "Verificarea și pregătirea structurii metalice existente", "mp", 350),
                ("6", "Remedierea locală a elementelor metalice deteriorate, dacă este cazul", "lucrare", 1),
                ("7", "Învelitoare tablă metalică profilată, grosime 0,5 mm", "mp", 350),
                ("8", "Elemente de fixare pentru tablă metalică, cu garnituri de etanșare", "buc", 2100),
                ("9", "Elemente de etanșare și accesorii pentru îmbinări", "ml", 70),
                ("10", "Elemente de tinichigerie, margini, racorduri, închideri", "ml", 90),
                ("11", "Elemente pentru evacuarea apelor pluviale", "ml", 60),
                ("12", "Curățarea și evacuarea deșeurilor rezultate în urma lucrărilor", "lucrare", 1),
            ]),
        ],
        "note_f3": "Sursa: Lista de cantități din Anexa la caietul de sarcini ETT54/3/1/318/05.10.2026 (S.R.T.F.C. "
                   "Brașov), anunț SCN1182052. Pozițiile 8–11 sunt „conform necesarului/măsurătorilor” – cantitățile "
                   "sunt estimate de noi și se confirmă la VIZIONAREA OBLIGATORIE a obiectivului. Deșeuri cu azbest "
                   "estimate la 4,5 t (cca. 13 kg/mp). Garanție de participare 600 lei (6 luni), garanție de bună execuție "
                   "5%. Fără cerințe de experiență, cifră de afaceri sau ISO.",
    },
    "salubris": {
        "fisier": "ADV1557141_Salubris_Iasi_hidroizolatie",
        "titlu": "Lucrări de reparații hidroizolație – birourile Salubris S.A., str. Aurel Vlaicu nr. 76, Iași",
        "autoritate": "SALUBRIS S.A. Iași",
        "autoritate_scurt": "Salubris S.A. Iași",
        "email_depunere": "office@salubris.ro (simultan cu postarea ofertei în SEAP)",
        "nr_anunt": "ADV1557141",
        "data_anunt": "12.11.2026",
        "termen": "16.11.2026",
        "ora": "09:00",
        "cpv": "45261310-0 – Lucrări de hidroizolare",
        "valoare": 13000.00,
        "amplasament": "Birourile Salubris S.A., str. Aurel Vlaicu nr. 76, Iași (terasă cca. 110 mp)",
        "procedura": "cumpărare directă",
        "durata": "7 (șapte) zile lucrătoare de la semnarea contractului și emiterea ordinului de începere "
                  "(maximum admis: 10 zile lucrătoare)",
        "garantie": "60",
        "catalog": True,
        "decizie": SALUBRIS_DECIZIE,
        "experienta": False, "acces": False, "vizita": False,
        "distanta": "cca. 300 km (Buzău – Focșani – Bacău – Iași)",
        "saptamani": 2,
        "durata_scurt": "7 zile lucrătoare (2 săptămâni)",
        "lucrari": [
            "pregătirea suprafeței în funcție de situația existentă – curățare, îndepărtarea porțiunilor desprinse, "
            "tăierea bulelor (110 mp);",
            "șapă de corecție a pantei spre gurile de scurgere (15 mp);",
            "amorsarea suprafeței cu amorsă bituminoasă (110 mp);",
            "strat izolant din membrană bituminoasă PA 4,2 mm cu ardezie, flexibilitate la rece ≤ -15 °C, termosudată "
            "(110 mp), cu racordări la atice;",
            "șlițuirea peretelui pentru rost (8 ml) și rost din tablă zincată fixat pe perete și etanșat (8 ml);",
            "izolarea gurilor de scurgere (2 buc.), evacuarea deșeurilor.",
        ],
        "etape": [
            "Etapa 1 – Ordinul de începere; aducerea echipei și a materialelor; protejarea zonei de lucru.",
            "Etapa 2 – Pregătirea suprafeței, șapa de corecție a pantei, uscarea.",
            "Etapa 3 – Amorsă, membrană PA 4,2 mm termosudată, racordări la atice și la gurile de scurgere.",
            "Etapa 4 – Șlițuire, rost din tablă zincată, etanșare; proba de etanșeitate, curățenie, recepție.",
        ],
        "materiale": [
            "membrană bituminoasă SBS ", Y("PA 4,2 mm cu ardezie, flexibilitate la rece ≤ -15 °C (Sika/Bauder/Icopal "
            "sau echivalent)"), ", amorsă bituminoasă cu solvent; șapă de ciment pentru pante; tablă zincată 0,50 mm "
            "pentru rost; mastic poliuretanic; butelii propan pentru termosudare.",
        ],
        "f3": [
            ("Lista cantităților de lucrări – solicitarea Salubris nr. 33347/12.11.2026", [
                ("1", "Pregătire suprafață în funcție de situația existentă", "mp", 110),
                ("2", "Turnare șapă corecție pantă", "mp", 15),
                ("3", "Amorsare suprafață", "mp", 110),
                ("4", "Strat izolant PA 4,2 mm (flexibilitate la rece -15 °C)", "mp", 110),
                ("5", "Șlițuire perete pentru rost", "ml", 8),
                ("6", "Rost tablă zincată pe perete", "ml", 8),
                ("7", "Izolații guri scurgere", "buc", 2),
            ]),
        ],
        "note_f3": "Sursa: solicitarea de ofertă Salubris S.A. nr. 33347/12.11.2026 (anunț ADV1557141). Termen de execuție "
                   "maxim 10 zile lucrătoare, garanție 60 de luni, garanție de bună execuție 7%, plata la 30 de zile. "
                   "Deplasarea și cazarea echipei (Buzău – Iași) sunt incluse în poziția 1. Informații tehnice: ing. "
                   "Carmen Damian, 0745 586 852, carmen.damian@salubris.ro.",
    },
})


# ---------------------------------------------------------------- licitații 16.11.2026
LICITATII.update({
    "cfr_sibiu_admin": {
        "fisier": "CN1098200_CFR_Sibiu_acoperis_admin",
        "valabilitate": "180 (o sută optzeci)",
        "titlu": "Lucrări de reparații acoperiș clădire administrativă la Depoul Sibiu – S.R.T.F.C. Brașov",
        "autoritate": "SNTFC „CFR CĂLĂTORI” S.A. – Sucursala Regională de Transport Feroviar de Călători Brașov",
        "autoritate_scurt": "S.R.T.F.C. Brașov",
        "email_depunere": "depunere online în SEAP (licitația deschisă CN1098200)",
        "nr_anunt": "CN1098200",
        "data_anunt": "09.11.2026",
        "termen": "14.12.2026",
        "ora": "15:00",
        "cpv": "45261910-6 – Reparare de acoperișuri",
        "valoare": 300000.00,
        "amplasament": "Depoul Sibiu – clădirea administrativă (construită în 1910), str. Dorobanților nr. 22, Sibiu "
                       "(învelitoare din țiglă ceramică pe șarpantă din lemn, cca. 1.000–1.300 mp, 10 coșuri de fum)",
        "procedura": "licitație deschisă (Legea nr. 99/2016)",
        "sectorial": True,
        "durata": "60 (șaizeci) de zile calendaristice de la predarea amplasamentului",
        "garantie": Y("36"),
        "catalog": False,
        "decizie": CFR_BV_DECIZIE,
        "experienta": False, "acces": False, "vizita": True,
        "distanta": "cca. 290 km (Buzău – Brașov – Făgăraș – Sibiu)",
        "saptamani": 9,
        "durata_scurt": "60 zile calendaristice (9 săptămâni)",
        "lucrari": [
            "organizarea și delimitarea zonei de lucru, protecția spațiilor de birouri care rămân în funcțiune, linii "
            "de ancorare și plase de protecție, schelă pe zonele de streașină;",
            "desfacerea învelitorii din țiglă ceramică (1.000 mp), sortarea țiglei recuperabile, coborârea controlată "
            "cu macaraua/tobogan, fără aruncare de la înălțime;",
            "demontarea șipcilor și a elementelor deteriorate ale suportului învelitorii;",
            "demolarea controlată, etapizată, a celor 10 coșuri de fum scoase din uz și închiderea golurilor în șarpantă;",
            "verificarea vizuală și tehnică a șarpantei, înlocuirea elementelor degradate (căpriori, pane, cosoroabe) "
            "cu lemn de rășinoase tratat ignifug și fungicid, consolidări locale cu eclise și scoabe;",
            "verificarea streașinii pe toată lungimea: înlocuirea elementelor degradate, curățarea și vopsirea în două "
            "straturi a elementelor din lemn păstrate;",
            "folie anticondens min. 120 g/mp, contrașipci și șipci din rășinoase tratate;",
            "învelitoare nouă din țiglă metalică 0,50 mm, RAL 8017 maro închis, cu protecție anticorozivă Z275 și "
            "strat de vopsea rezistent la UV, fixată cu șuruburi autoforante cu garnitură EPDM;",
            "tinichigerie aferentă: coame, dolii, șorțuri la racorduri, borduri de streașină și fronton, pazii;",
            "încărcarea, transportul și predarea deșeurilor (țiglă, moloz, lemn) la depozit autorizat, cu documente "
            "de trasabilitate; curățarea finală și eliberarea amplasamentului.",
        ],
        "etape": [
            "Etapa 0 – Vizionarea obligatorie a obiectivului (str. Dorobanților nr. 22, Sibiu) înainte de depunerea "
            "ofertei; măsurători proprii și constatarea stării șarpantei și a streașinii.",
            "Etapa 1 – Predarea amplasamentului; delimitare, semnalizare, linii de ancorare, schelă la streașină, "
            "protecția spațiilor interioare.",
            "Etapa 2 – Desfacerea învelitoarei pe tronsoane (maxim ce se poate reacoperi în 1–2 zile), sortarea și "
            "coborârea țiglei, demolarea coșurilor de fum.",
            "Etapa 3 – Verificarea șarpantei pe tronsonul descoperit, înlocuirea elementelor degradate, consolidări, "
            "tratamente; constatările se consemnează cu beneficiarul.",
            "Etapa 4 – Folie anticondens, contrașipci, șipci și țiglă metalică pe același tronson (protecție provizorie "
            "cu prelate la sfârșitul zilei).",
            "Etapa 5 – Coame, dolii, șorțuri, borduri, pazii; refacerea și vopsirea streașinii.",
            "Etapa 6 – Evacuarea deșeurilor, curățenia finală, verificarea etanșeității, recepția la terminarea lucrărilor.",
        ],
        "materiale": [
            "țiglă metalică ", Y("BILKA Clasic/Balcanic mat 0,50 mm, RAL 8017 maro închis, Z275 (sau echivalent)"),
            ", cu coame, capace de coamă, dolii, borduri și șorțuri din același material; folie anticondens "
            "min. 120 g/mp; șipci 50×50 mm și contrașipci 30×50 mm din rășinoase tratate; cherestea de rășinoase "
            "tratată pentru înlocuirea elementelor de șarpantă; scoabe, eclise, cuie, șuruburi autoforante cu EPDM; "
            "bandă de coamă ventilată; lazură/vopsea pentru lemn exterior, tratament fungicid-insecticid.",
        ],
        "f3": [
            ("Lista de cantități – Anexa la CS ETT54/3/1/319/2026", [
                ("1", "Organizarea și pregătirea zonei de lucru", "lucrare", 1),
                ("2", "Demontarea învelitorii existente din țiglă ceramică", "mp", 1000),
                ("3", "Sortarea, manipularea și coborârea țiglei rezultate din demolare", "mp", 1000),
                ("4", "Demontarea elementelor deteriorate ale suportului învelitorii", "mp", 1000),
                ("5", "Demolarea coșurilor de fum existente", "buc", 10),
                ("6", "Verificarea și remedierea locală a șarpantei existente", "lucrare", 1),
                ("7", "Înlocuirea elementelor de șarpantă degradate (estimat 30% din suprafață)", "mp", 300),
                ("8", "Demontarea streașinii existente deteriorate (estimat)", "ml", 70),
                ("9", "Montare streașină nouă din lemn (estimat)", "ml", 70),
                ("10", "Vopsirea elementelor din lemn ale streașinii (estimat)", "mp", 120),
                ("11", "Montare folie anticondens", "mp", 1000),
                ("12", "Montare șipci și contrașipci din lemn", "mp", 1000),
                ("13", "Montare învelitoare din tablă tip țiglă metalică", "mp", 1000),
                ("14", "Elemente de tinichigerie aferente acoperișului (borduri streașină și fronton) (estimat)", "ml", 160),
                ("15", "Coame pentru învelitoare (estimat)", "ml", 70),
                ("16", "Șorțuri, dolii, pazii și elemente de racord (estimat)", "ml", 90),
                ("17", "Încărcarea și transportul materialelor rezultate din demolări (estimat)", "t", 60),
                ("18", "Curățarea finală și eliberarea amplasamentului", "lucrare", 1),
            ]),
        ],
        "note_f3": "Sursa: Lista de cantități din Anexa la caietul de sarcini ETT54/3/1/319/05.10.2026 (S.R.T.F.C. "
                   "Brașov), anunț CN1098200. Pozițiile 7–10 și 14–17 sunt „conform constatărilor/măsurătorilor” – "
                   "cantitățile sunt estimate de noi (30% din șarpantă de înlocuit; perimetru streașină cca. 140 ml; "
                   "deșeuri cca. 45 t țiglă + 12 t moloz din coșuri + 3 t lemn) și se confirmă la VIZIONAREA "
                   "OBLIGATORIE. Garanție de participare 3.000 lei (6 luni), garanție de bună execuție 5%. "
                   "Fără cerințe de experiență, cifră de afaceri sau ISO.",
    },
})


# ---------------------------------------------------------------- licitații 18.11.2026
BJT_DECIZIE = ("dr. Dima Ligia – Manager; Mihai Tudor – referent de specialitate (persoană de contact); Elena Ghionu – "
               "achiziții")
DGASPC_DECIZIE = (
    "Director General: Ristea Mihaela; Director Executiv Protecția Copilului: Marcoșanu Silviu Mihai; Director Executiv "
    "Asistență Socială Persoane Adulte: Urziceanu Petre; Serviciul Achiziții Publice Urmărire Contracte: Sotir Anișoara, "
    "Constantin Antonela Dana, Grosu Chirana Georgiana, Neculae Laura-Cristina, Aldea Răzvan Gabriel, Masaca Cosmin, "
    "Șerban Laura Alexandra; Compartiment Taxe, Avize, Autorizații și Urmărire Lucrări: Lupu Marius Dumitru, Iacoboaie "
    "Mirabela; Serviciu Juridic Contencios: Ghițulescu Anca, Guță Titiana, Larie Petre, Checiu Cristina, Ciocîrlan "
    "Monica-Mihaela, Creangă Cătălina Georgeta Iuliana; Serviciul Contabilitate, Buget: Stamule Zoica, Neculaev Alina, "
    "Geamalingă Tanța, Cosma Felicia Iuliana, Marius Eremia, Butoi Meral, Acatrinei Florentina; CIRUL-DASPA: Macovei "
    "Marius, Andrei Liviu, Mititelu Radu, Meroșneac Cătălin, Monea Florin, Nurla Aysel, Iordache Ionela")
LICITATII.update({
    "bjt_tulcea": {
        "fisier": "ADV1558035_Biblioteca_Tulcea_acoperis",
        "valabilitate": "30 (treizeci)",
        "titlu": "Lucrări de reparații curente acoperiș Biblioteca Județeană „Panait Cerna” Tulcea",
        "autoritate": "BIBLIOTECA JUDEȚEANĂ „PANAIT CERNA” TULCEA",
        "autoritate_scurt": "Biblioteca Județeană „Panait Cerna” Tulcea",
        "email_depunere": "bjt@tulcealibrary.ro (subiect: „Ofertă – Lucrări reparații curente acoperiș – Biblioteca "
                          "Județeană Panait Cerna Tulcea”)",
        "nr_anunt": "ADV1558035",
        "data_anunt": "17.11.2026",
        "termen": "20.11.2026",
        "ora": "12:00",
        "cpv": "45453000-7 – Lucrări de reparații generale și de renovare",
        "valoare": 32979.69,
        "amplasament": "Sediul Bibliotecii Județene „Panait Cerna”, str. Isaccei nr. 16A, Tulcea – acoperiș tip "
                       "terasă/pantă mică (≤ 7%), 217 mp",
        "procedura": "cumpărare directă",
        "durata": "10 (zece) zile calendaristice de la comunicarea ordinului de începere (maximum admis: 30 de zile)",
        "garantie": "24",
        "catalog": False,
        "decizie": BJT_DECIZIE,
        "experienta": False, "acces": False, "vizita": False,
        "distanta": "cca. 230 km (Buzău – Brăila – Măcin – Tulcea)",
        "saptamani": 2,
        "durata_scurt": "10 zile calendaristice (2 săptămâni)",
        "lucrari": [
            "curățarea și aspirarea suprafețelor în vederea pregătirii pentru hidroizolație (217 mp);",
            "amorsarea suprafețelor cu amorsă bituminoasă, conform instrucțiunilor producătorului (217 mp);",
            "strat de membrană bituminoasă cauciucată (SBS) cu inserție din împâslitură de fibră de sticlă, simplu "
            "așezată/termosudată, pe suprafețe orizontale sau înclinate până la 7%, inclusiv suportul la perforațiile "
            "stratului de difuziune (217 mp);",
            "stratul hidroizolant cauciucat de finisaj – membrană SBS cu ardezie, rezistentă UV, cu racordări la atice, "
            "străpungeri și guri de scurgere (217 mp);",
            "protejarea elementelor adiacente, colectarea și evacuarea deșeurilor, predarea fișelor tehnice și a "
            "declarațiilor de performanță ale materialelor.",
        ],
        "etape": [
            "Etapa 1 – Ordinul de începere; aducerea echipei și a materialelor; protejarea zonelor adiacente.",
            "Etapa 2 – Curățarea și aspirarea suprafeței, verificarea suportului; amorsarea și uscarea amorsei.",
            "Etapa 3 – Stratul 1 de membrană SBS cu fibră de sticlă, cu suprapuneri de min. 10 cm.",
            "Etapa 4 – Stratul 2 de membrană SBS cu ardezie, decalat față de stratul 1, racordări la atice și guri.",
            "Etapa 5 – Verificarea etanșeității, curățenia, evacuarea deșeurilor, recepția.",
        ],
        "materiale": [
            "membrane bituminoase elastomerice SBS ", Y("4 kg/mp cu fibră de sticlă (strat 1) și cu ardezie gri (strat 2) – "
            "Sika/Bauder/Icopal sau echivalent"), ", amorsă bituminoasă cu solvent; butelii propan pentru termosudare; "
            "mastic poliuretanic pentru racorduri.",
        ],
        "f3": [
            ("Formular 3 – Anexa nr. 1 la anunțul de publicitate nr. 2579/17.11.2026", [
                ("1", "Curățarea și aspirarea suprafețelor", "mp", 217),
                ("2", "Amorsarea suprafețelor în vederea executării hidroizolației", "mp", 217),
                ("3", "Executarea stratului de membrană cauciucată la acoperișuri și terase", "mp", 217),
                ("4", "Executarea stratului hidroizolant cauciucat la terase și acoperișuri", "mp", 217),
                ("5", "Alte cheltuieli (deplasare și cazare echipă, material mărunt, evacuare deșeuri)", "buc", 1),
            ]),
        ],
        "note_f3": "Sursa: Anexa nr. 1 – Formular 3 din anunțul de publicitate nr. 2579/17.11.2026 (ADV1558035). Fără "
                   "cerințe de experiență, cifră de afaceri sau ISO. Vizitarea amplasamentului este recomandată (nu "
                   "obligatorie) – contact Mihai Tudor, tel. 0240 513 833. Se atașează fișele tehnice și declarațiile "
                   "de performanță ale membranelor. Câștigătorul publică oferta în catalogul electronic SEAP.",
    },
    "dgaspc_mangalia": {
        "fisier": "ADV1558019_DGASPC_Mangalia_terasa",
        "titlu": "Execuția lucrărilor necesare remedierii infiltrațiilor și refacerii sistemului de termo-hidroizolație "
                 "al acoperișului tip terasă al imobilului în care funcționează Centrul de Zi și Recuperare Mangalia",
        "autoritate": "DIRECȚIA GENERALĂ DE ASISTENȚĂ SOCIALĂ ȘI PROTECȚIA COPILULUI CONSTANȚA",
        "autoritate_scurt": "D.G.A.S.P.C. Constanța",
        "email_depunere": "office@dgaspc-ct.ro / achizitii@dgaspc-ct.ro (sau depunere la sediu: str. Decebal nr. 22, "
                          "et. III, Constanța)",
        "nr_anunt": "ADV1558019",
        "data_anunt": "17.11.2026",
        "termen": "25.11.2026",
        "ora": "14:00",
        "cpv": "45261900-3 – Lucrări de reparare și de întreținere de acoperișuri",
        "valoare": 511297.80,
        "amplasament": "Centrul de Zi și Recuperare Mangalia (D.G.A.S.P.C. Constanța) – acoperiș tip terasă, cca. 650 mp",
        "procedura": "cumpărare directă",
        "durata": "40 (patruzeci) de zile calendaristice de la primirea ordinului de începere (maximum admis: 45 de zile)",
        "garantie": "24",
        "catalog": False,
        "decizie": DGASPC_DECIZIE,
        "experienta": False, "acces": False, "vizita": False,
        "distanta": "cca. 270 km (Buzău – Slobozia – Constanța – Mangalia)",
        "saptamani": 6,
        "durata_scurt": "40 zile calendaristice (6 săptămâni)",
        "lucrari": [
            "desfacerea hidroizolațiilor și termoizolațiilor existente, inclusiv straturile de protecție și suport "
            "(650 mp), curățarea cu peria de sârmă și rașchetă și spălarea cu jet de apă sub presiune;",
            "colectarea, coborârea și transportul molozului (cca. 26 t) la depozit autorizat, cu documente de trasabilitate;",
            "amorsă bituminoasă în două straturi și barieră de vapori din membrană bituminoasă cu inserție de aluminiu, "
            "sudată cu flacăra (650 mp);",
            "termoizolație din plăci de polistiren extrudat XPS 10 cm, 500 kPa, lipite cu spumă poliuretanică, folie de "
            "separație PE (650 mp);",
            "plasă sudată Ø6/100×100 mm și șapă de egalizare și pantă din beton B400 cu fibre de polipropilenă (49 mc);",
            "hidroizolație în două straturi: membrană SBS 4 kg/mp fără ardezie + membrană SBS cu ardezie gri, "
            "rezistentă UV (650 mp); racordări la atice și scafe (150 ml), etanșarea rosturilor de dilatare (150 ml);",
            "11 guri de scurgere Ø125 cu flanșe hidroizolante și parafrunzare, 10 dispozitive de ventilare a stratului "
            "de difuzie, profile de dilatare Hidroplasto (150 m) și înlocuirea ușii de acces pe terasă.",
        ],
        "etape": [
            "Etapa 1 – Ordinul de începere (începerea efectivă în max. 3 zile lucrătoare); predarea amplasamentului, "
            "organizarea de șantier, protecții, protejarea spațiilor centrului care rămâne în funcțiune.",
            "Etapa 2 – Desfacerea pe tronsoane a straturilor vechi, curățare, spălare, evacuarea molozului; protecție "
            "provizorie cu prelate/membrană la sfârșitul zilei.",
            "Etapa 3 – Amorsă, barieră de vapori, XPS 10 cm, folie PE, plasă sudată.",
            "Etapa 4 – Șapa de pantă din beton B400 cu fibre (turnare mecanizată), maturare și uscare.",
            "Etapa 5 – Guri de scurgere, flanșe, dispozitive de ventilare; hidroizolația în 2 straturi, racordări la "
            "atice și rosturi, profile de dilatare.",
            "Etapa 6 – Ușa de acces pe terasă, proba de etanșeitate (inundare), curățenia finală, recepția.",
        ],
        "materiale": [
            "membrane bituminoase SBS ", Y("4 kg/mp fără ardezie și cu ardezie gri rezistentă UV, barieră de vapori cu "
            "inserție de Al (Sika/Bauder/Icopal sau echivalent)"), "; polistiren extrudat XPS 100 mm, 500 kPa ",
            Y("(Austrotherm XPS TOP 50 sau echivalent)"), "; amorsă bituminoasă; bitum H 80/90; folie PE; plasă sudată "
            "Ø6/100×100; beton B400 cu fibre de polipropilenă; guri de scurgere Ø125 cu flanșă; profile de dilatare "
            "Hidroplasto/Nitriflex 830 sau echivalent; tablă zincată 0,5 mm.",
        ],
        "f3": [
            ("F3 – Lista cu cantități de lucrări (Hidroizolație Mangalia – estimare hidroizolație)", [
                ("1", "RCSE42A% – Desfacerea hidroizolațiilor, termoizolațiilor lipite cu bitum, cu straturile de "
                      "protecție și suport, la terase", "mp", 650),
                ("2", "RpIzE16A% – Curățarea confecțiilor și construcțiilor cu peria de sârmă și rașchetă", "mp", 650),
                ("3", "TRB01A11 – Colectarea, manipularea și coborârea molozului din decopertare (cca. 40 kg/mp)", "t", 26),
                ("4", "TRB01A11 – Transportul molozului cu autobasculanta la groapa autorizată/centru de reciclare, "
                      "inclusiv taxe", "t", 26),
                ("5", "H1B02A1 – Pregătirea suprafeței orizontale prin spălare cu jet de apă sub presiune", "mp", 650),
                ("6", "IZF01A – Amorsarea suprafețelor cu soluție bituminoasă (bitum tăiat), în două straturi", "mp", 650),
                ("6.4", "6101456 – Amorsă bituminoasă Matizol", "kg", 97.5),
                ("7", "IZF04B(1) – Montarea barierei de vapori din membrană bituminoasă cu inserție de aluminiu, sudată "
                      "cu flacăra (inclusiv suprapuneri 10%)", "mp", 650),
                ("7.4", "2600048 – Bitum pentru materiale și lucrări de hidroizolații tip H 80/90", "kg", 975),
                ("7.5", "2601640 – Barieră de vapori din membrană bituminoasă cu inserție de aluminiu", "mp", 715),
                ("8", "IZG06B(1) – Montarea plăcilor de polistiren extrudat XPS 10 cm, 500 kPa, lipite cu spumă "
                      "poliuretanică", "mp", 650),
                ("8.1", "20024709 – Polistiren extrudat XPS 100 mm (Austrotherm XPS 50 GK sau echivalent)", "mp", 715),
                ("9", "IZF09XA(1) – Așezarea foliei de separație din polietilenă peste polistiren", "mp", 650),
                ("9.2", "11232 – Folie PE", "mp", 682.5),
                ("10", "TE06C1(1) – Plasă de armătură sudată STNB d=6 mm, ochiuri 100×100 mm", "mp", 650),
                ("11", "CA01XG02(1) – Prepararea și turnarea șapei mecanizate de egalizare și pantă din beton B400 cu "
                       "fulgi de polipropilenă", "mc", 49),
                ("11.5", "2100086 – Ciment Portland", "kg", 20090),
                ("12", "IZF02A(1) – Hidroizolație strat 1: membrană bituminoasă SBS termosudabilă 4 kg/mp, fără ardezie",
                 "mp", 650),
                ("13", "IZF02B(1) – Hidroizolație strat 2 (finisaj): membrană bituminoasă SBS termosudabilă cu "
                       "protecție de ardezie gri, rezistentă UV", "mp", 650),
                ("13.5", "2600050 – Bitum pentru materiale și lucrări de hidroizolații tip H 82/92", "kg", 325),
                ("14", "IZF05A1(1) – Tratamentul și etanșarea verticală a aticului și scafelor cu membrană bituminoasă "
                       "cu ardezie (ridicare 30–40 cm)", "ml", 150),
                ("14.4", "2600907 – Carton bitumat cu strat de acoperire din ardezie", "mp", 172.5),
                ("15", "RpIzC20E% – Etanșarea rosturilor de dilatare/tasare ale aticurilor, cu protecție din tablă "
                       "(benzi de cca. 300 mm)", "ml", 150),
                ("15.2", "2600660 – Pânză bitumată fără strat de acoperire PI 40", "mp", 153),
                ("15.3", "2600751 – Carton bitumat cu strat de acoperire din nisip CA 300", "mp", 45),
                ("16", "IZF02XG(1) – Flanșă cu hidroizolație pentru guri de scurgere", "buc", 21),
                ("16.3", "Bitum pentru materiale și lucrări de hidroizolații tip H 80/90", "kg", 42),
                ("16.4", "7336692 – Flanșă cu hidroizolație", "buc", 21),
                ("17", "CE25A01+ – Gură de scurgere pentru terase și balcoane, neagră, L 500 mm, D 125 mm", "buc", 11),
                ("17.1", "20055418 – Racord jgheab/burlan Ø125", "buc", 11),
                ("18", "RCSE26B%(1) – Dispozitive de comunicare a stratului de difuzie cu atmosfera, din tablă zincată "
                       "0,5 mm, cu căciulă de protecție", "buc", 10),
                ("19", "RCSE27A%(1) – Parafrunzar Ø120 mm din sârmă zincată pentru gurile de scurgere", "buc", 11),
                ("20", "CR01A02+(1) – Profile de dilatare Hidroplasto pentru pardoseală, tip metal și Nitriflex 830",
                 "m", 150),
                ("20.1", "20025301 – Profil de prindere Hidroplasto metal pentru hidroizolație", "m", 150),
                ("21", "TRA01A50 – Transportul rutier al materialelor și semifabricatelor cu autobasculanta, 50 km",
                 "t", 150),
                ("22", "M3D13A1(1) – Ușă de vizitare/acces pe terasă, înlocuită, cu etanșare", "buc", 1),
            ]),
        ],
        "note_f3": "Sursa: F3 – Lista cu cantități de lucrări din documentația ADV1558019 (Deviz 360, „Hidroizolație "
                   "Mangalia”), cu pozițiile de materiale x.y păstrate separat, ca în formularul autorității. Fără "
                   "cerințe de experiență, cifră de afaceri sau ISO. În propunerea tehnică se nominalizează șeful de "
                   "lucrări cu diplomă de inginer în construcții și personalul calificat pentru hidroizolații. Garanție "
                   "minimă 24 de luni. Depunere pe e-mail office@dgaspc-ct.ro / achizitii@dgaspc-ct.ro sau la sediu "
                   "(str. Decebal nr. 22, et. III). Relații: Sotir Anișoara / Șerban Laura Alexandra, 0241 480 876, "
                   "0732 598 294. Clarificări cu min. 3 zile lucrătoare înainte de termen (până pe 20.11.2026).",
    },
})


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
        ["Durata de execuție ofertată", [L["durata"]]],
        ["Începerea lucrărilor", "În maximum 72 de ore de la emiterea ordinului de începere"],
        ["Garanția lucrărilor", [L["garantie"], " de luni calendaristice de la recepția la terminarea lucrărilor"]],
        ["Distanța sediu – amplasament", [Y(L["distanta"])]],
    ], widths=[5, 12])

    P(doc, B("2. Obiectul lucrărilor"))
    P(doc, "Ne angajăm să executăm integral lucrările prevăzute în caietul de sarcini și în lista de cantități F3. "
           "Lucrările cuprind, în principal:", align="j")
    for t in L["lucrari"]:
        bullet(doc, t)
    
    P(doc, B("3. Organizarea de șantier"), before=6)
    for t in [
        "delimitarea și semnalizarea zonei de lucru și a zonei de depozitare a materialelor, stabilite împreună cu "
        "reprezentantul beneficiarului, astfel încât activitatea din clădire să fie afectată cât mai puțin;",
        "depozitarea materialelor pe paleți și suporți, ferite de umezeală și de lovituri, în ambalajul original până la montaj;",
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
    for t in L["etape"]:
        bullet(doc, t)

    P(doc, B("5. Resurse umane"), before=6)
    table(doc, ["Funcție / meserie", "Număr", "Observații"], [
        ["Șef de echipă / conducător tehnic al lucrării", "1", "Răspunde de execuție și calitate"],
        ["Dulgheri", Y("3"), "Structură lemn, astereală, rigle"],
        ["Montatori acoperiș / tinichigii", Y("3"), "Învelitoare, tinichigerie, sistem pluvial"],
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
           "sau declarație de performanță, certificat de garanție și, după caz, agrement tehnic. Principalele materiale: ",
      *L["materiale"],
      " Specificațiile care indică o anumită marcă se consideră însoțite de mențiunea „sau echivalent”.", align="j")

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
           f"executa „{L['titlu']}” pentru suma de ", suma(L, "total"), " lei fără T.V.A., la care se adaugă T.V.A. în "
      "valoare de ", suma(L, "tva"), " lei, în conformitate cu anexa la formularul de ofertă.", align="j")
    P(doc, "2. Ne angajăm ca, în cazul în care oferta noastră este stabilită câștigătoare, să începem lucrările cât mai "
           "curând posibil după primirea ordinului de începere și să terminăm lucrările în ",
      L["durata"], ".", align="j")
    P(doc, "3. Ne angajăm să menținem această ofertă valabilă pentru o durată de ", Y(L.get("valabilitate", "90 (nouăzeci)")),
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
           f"{FIRMA['denumire']}", align="j")
    P(doc, "OPERATOR ECONOMIC", after=0)
    P(doc, B(FIRMA["denumire"]), after=0)
    P(doc, "Semnătura autorizată și ștampila")


def suma(L, cheie):
    c = L.get("calc")
    if not c:
        return None
    v = c["total"] + c["tva"] if cheie == "total_tva" else c[cheie]
    return lei(v)


def capitole(L):
    return [cap for cap, _ in L["f3"]] if len(L["f3"]) > 1 else []


def suma_cap(L, cap):
    """Valoarea capitolului cu CAM, indirecte și profit repartizate proporțional (ca în foaia Grafic);
    diferența de rotunjire de 1–2 bani se trece la ultimul capitol, ca suma să dea exact totalul."""
    c = L.get("calc")
    if not c:
        return None
    caps = capitole(L)
    val = {}
    for cp in caps:
        directe = sum(sum(v.values()) for x, *_, v in c["linii"] if x == cp)
        val[cp] = round(directe / c["directe"] * c["total"], 2)
    val[caps[-1]] = round(c["total"] - sum(val[cp] for cp in caps[:-1]), 2)
    return lei(val[cap])


def anexa_oferta(doc, L):
    P(doc, B("ANEXA LA FORMULARUL DE OFERTĂ"), align="c", size=13, after=10)
    P(doc, f"Titlul contractului: „{L['titlu']}”", after=8)
    table(doc, ["Denumire", "Valoare lei (fără T.V.A.)"], [
        *[[cap, suma_cap(L, cap)] for cap in capitole(L)],
        [B("Total general deviz ofertă"), [B(suma(L, "total"))] if L.get("calc") else None],
        ["T.V.A. 21%", suma(L, "tva")],
        [B("Total cu T.V.A."), [B(suma(L, "total_tva"))] if L.get("calc") else None],
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
    elif cheie in ("romeo", "loto", "cfr_galati", "cfr_mangalia", "cfr_sibiu", "salubris", "bjt_tulcea",
                   "dgaspc_mangalia"):
        opis += ["Declarație privind neîncadrarea în art. 164, 165 și 167 din Legea nr. 98/2016",
                 "Lista lucrărilor similare executate, cu documente constatatoare (informativ)"]
        sectiuni += [decl_164_165_167, experienta]
    if L.get("sectorial"):
        opis.insert(1, "DUAE completat în SEAP (răspuns la DUAE-cererea entității contractante)")
        opis.insert(2, "Dovada constituirii garanției de participare")
    if L.get("vizita"):
        opis.insert(3, "Confirmarea vizitării amplasamentului (obligatorie – fără ea oferta nu este luată în considerare)")
    scrisoare(doc, L, opis)
    for s in sectiuni:
        page_break(doc)
        s(doc, L)
    if L.get("sectorial"):
        lege_sectoriala(doc, L)
    footer(doc, f"{FIRMA['denumire']} – Ofertă „{L['titlu']}”")
    OUT.mkdir(exist_ok=True)
    path = OUT / f"Dosar_oferta_BILZI_{L['fisier']}.docx"
    doc.save(path)
    compacteaza_docx(path)
    return path


SECTORIAL = [
    ("respectarea art. 51 din Legea nr. 98/2016 privind achizițiile publice",
     "respectarea obligațiilor de mediu, sociale și de muncă – Legea nr. 99/2016 privind achizițiile sectoriale"),
    ("respectarea art. 51 din Legea nr. 98/2016", "obligațiile de mediu, sociale și de muncă (Legea nr. 99/2016)"),
    ("art. 59–60 din Legea nr. 98/2016 privind achizițiile publice",
     "art. 72–73 din Legea nr. 99/2016 privind achizițiile sectoriale"),
    ("art. 59–60 din Legea nr. 98/2016", "art. 72–73 din Legea nr. 99/2016"),
    ("prevăzute la art. 60 din Legea nr. 98/2016", "prevăzute la art. 73 din Legea nr. 99/2016"),
    ("art. 21 alin. (5) și (6) din H.G. nr. 395/2016", "H.G. nr. 394/2016"),
    ("având în vedere și prevederile art. 61 din Legea nr. 98/2016", "conform Legii nr. 99/2016"),
    ("ART. 164, 165 ȘI 167 DIN LEGEA NR. 98/2016", "ART. 177, 178 ȘI 180 DIN LEGEA NR. 99/2016"),
    ("art. 164, 165 și 167 din Legea nr. 98/2016", "art. 177, 178 și 180 din Legea nr. 99/2016"),
    ("art. 164 alin. (1) din Legea nr. 98/2016", "art. 177 alin. (1) din Legea nr. 99/2016"),
    ("art. 165 din Legea nr. 98/2016", "art. 178 din Legea nr. 99/2016"),
    ("art. 167 alin. (1) din Legea nr. 98/2016", "art. 180 alin. (1) din Legea nr. 99/2016"),
    ("anunțului de cumpărare directă", "anunțului de participare simplificat"),
    ("Legea nr. 98/2016", "Legea nr. 99/2016"),
]


def lege_sectoriala(doc, L=None):
    """Contract sectorial (CFR): trimiterile la Legea nr. 98/2016 devin trimiteri la Legea nr. 99/2016."""
    pars = list(doc.paragraphs)
    for t in doc.tables:
        for row in t.rows:
            for c in row.cells:
                pars += c.paragraphs
    for par in pars:
        for run in par.runs:
            x = run.text
            for a, b in SECTORIAL:
                x = x.replace(a, b)
            if L and "licitație deschisă" in L.get("procedura", ""):
                x = x.replace("anunțului de participare simplificat", "anunțului de participare")
            if x != run.text:
                run.text = x


def compacteaza_docx(path):
    """Scoate din șablonul python-docx părțile nefolosite (stylesWithEffects, miniatura), ca fișierul să fie mic."""
    scoase = ("word/stylesWithEffects.xml", "docProps/thumbnail.jpeg")
    with zipfile.ZipFile(path) as z:
        parti = {n: z.read(n) for n in z.namelist() if n not in scoase}
    for n in ("[Content_Types].xml", "_rels/.rels", "word/_rels/document.xml.rels"):
        x = parti[n].decode("utf-8")
        x = re.sub(r'<Override[^>]*(stylesWithEffects|thumbnail)[^>]*/>', "", x)
        x = re.sub(r'<Relationship[^>]*(stylesWithEffects|thumbnail)[^>]*/>', "", x)
        parti[n] = x.encode("utf-8")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for n, data in parti.items():
            z.writestr(n, data)


# ---------------------------------------------------------------- propunere financiară (xlsx)

GALBEN = PatternFill("solid", fgColor="FFFF00")
GRI = PatternFill("solid", fgColor="D9E2F3")
SUBTOTAL = PatternFill("solid", fgColor="EDEDED")
LEI = '#,##0.00'


def genereaza_xlsx(cheie):
    """F3 cu articolele și cantitățile din documentația de atribuire; prețurile unitare (galben) se completează,
    totalurile, TVA-ul și graficul de execuție se calculează din formule."""
    L = LICITATII[cheie]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "F3"
    bold = Font(bold=True)
    ws["A1"] = FIRMA["denumire"]
    ws["A1"].font = bold
    ws["A2"] = f"Obiectiv: {L['titlu']} – Beneficiar: {L['autoritate_scurt']}"
    ws["A3"] = "FORMULAR F3 – Lista cu cantitățile de lucrări pe categorii de lucrări (cu prețuri)"
    ws["A3"].font = Font(bold=True, size=13)
    hdr = ["Nr.", "Cod articol", "Capitol de lucrări", "U.M.", "Cantitate", "Preț unitar fără TVA (lei)",
           "Total fără TVA (lei)"]
    for c, h in enumerate(hdr, 1):
        cell = ws.cell(5, c, h)
        cell.font = bold
        cell.fill = GRI
    r = 6
    subtotaluri, capitole = [], []
    for cap, articole in L["f3"]:
        ws.cell(r, 1, cap).font = bold
        r += 1
        first = r
        for n, (cod, den, um, cant) in enumerate(articole, 1):
            ws.cell(r, 1, n)
            ws.cell(r, 2, cod)
            ws.cell(r, 3, den)
            ws.cell(r, 4, um)
            ws.cell(r, 5, cant).number_format = '#,##0.00'
            ws.cell(r, 6).fill = GALBEN
            ws.cell(r, 6).number_format = LEI
            ws.cell(r, 7, f"=ROUND(E{r}*F{r},2)").number_format = LEI
            r += 1
        ws.cell(r, 3, f"Total {cap}").font = bold
        ws.cell(r, 7, f"=SUM(G{first}:G{r - 1})").number_format = LEI
        for c in range(1, 8):
            ws.cell(r, c).fill = SUBTOTAL
        subtotaluri.append(f"G{r}")
        capitole.append((cap, f"F3!G{r}"))
        r += 2
    tot = r
    ws.cell(r, 3, "TOTAL GENERAL fără TVA").font = bold
    ws.cell(r, 7, "=" + "+".join(subtotaluri)).number_format = LEI
    ws.cell(r + 1, 3, "TVA 21%")
    ws.cell(r + 1, 7, f"=ROUND(G{r}*0.21,2)").number_format = LEI
    ws.cell(r + 2, 3, "TOTAL GENERAL cu TVA").font = bold
    ws.cell(r + 2, 7, f"=G{r}+G{r + 1}").number_format = LEI
    ws.cell(r + 4, 3, "Valoarea estimată de autoritatea contractantă (fără TVA) – oferta nu o poate depăși")
    ws.cell(r + 4, 7, L["valoare"]).number_format = LEI
    ws.cell(r + 5, 3, "Diferență față de valoarea estimată (trebuie să fie ≥ 0)")
    ws.cell(r + 5, 7, f"=G{r + 4}-G{tot}").number_format = LEI
    ws.cell(r + 7, 1, "Notă: " + L["note_f3"]).font = Font(italic=True, size=9)
    ws.cell(r + 8, 1, "Prețurile unitare includ materialele, manopera, utilajele, transportul, cheltuielile indirecte și "
                      "profitul. Se completează doar celulele galbene.").font = Font(italic=True, size=9)
    ws.cell(r + 10, 1, f"Data: ____________        {FIRMA['admin']} – administrator, semnătura și ștampila")
    for col, w in zip("ABCDEFG", [6, 11, 70, 7, 11, 16, 18]):
        ws.column_dimensions[col].width = w
    for row in ws.iter_rows(min_row=6, max_row=r, max_col=3):
        row[2].alignment = openpyxl.styles.Alignment(wrap_text=True, vertical="top")

    # Anexa la formularul de ofertă
    an = wb.create_sheet("Anexa oferta")
    an["A1"] = "ANEXA LA FORMULARUL DE OFERTĂ"
    an["A1"].font = Font(bold=True, size=13)
    an["A2"] = f"Titlul contractului: „{L['titlu']}”"
    an["A4"], an["B4"] = "Denumire", "Valoare lei (fără TVA)"
    an["A4"].font = an["B4"].font = bold
    for k, (cap, ref) in enumerate(capitole, 5):
        an.cell(k, 1, cap)
        an.cell(k, 2, f"={ref}").number_format = LEI
    k = 5 + len(capitole)
    an.cell(k, 1, "TOTAL fără TVA").font = bold
    an.cell(k, 2, f"=F3!G{tot}").number_format = LEI
    an.cell(k + 1, 1, "TVA 21%")
    an.cell(k + 1, 2, f"=F3!G{tot + 1}").number_format = LEI
    an.cell(k + 2, 1, "TOTAL cu TVA").font = bold
    an.cell(k + 2, 2, f"=F3!G{tot + 2}").number_format = LEI
    an.column_dimensions["A"].width = 70
    an.column_dimensions["B"].width = 22

    # Grafic de execuție fizic și valoric (procente pe săptămâni, de completat)
    gr = wb.create_sheet("Grafic")
    sapt = L["saptamani"]
    gr["A1"] = "GRAFIC DE EXECUȚIE FIZIC ȘI VALORIC"
    gr["A1"].font = Font(bold=True, size=13)
    gr["A2"] = f"Durata: {L['durata']}. În celulele galbene se trece procentul din capitol executat în fiecare săptămână."
    gr.cell(4, 1, "Capitol").font = bold
    gr.cell(4, 2, "Valoare (lei)").font = bold
    for w in range(sapt):
        gr.cell(4, 3 + w, f"S{w + 1}").font = bold
    gr.cell(4, 3 + sapt, "Total %").font = bold
    for k, (cap, ref) in enumerate(capitole, 5):
        gr.cell(k, 1, cap)
        gr.cell(k, 2, f"={ref}").number_format = LEI
        for w in range(sapt):
            gr.cell(k, 3 + w).fill = GALBEN
            gr.cell(k, 3 + w).number_format = '0%'
        c0 = openpyxl.utils.get_column_letter(3)
        c1 = openpyxl.utils.get_column_letter(2 + sapt)
        gr.cell(k, 3 + sapt, f"=SUM({c0}{k}:{c1}{k})").number_format = '0%'
    k = 5 + len(capitole)
    gr.cell(k, 1, "Valoare executată pe săptămână (lei)").font = bold
    for w in range(sapt):
        col = openpyxl.utils.get_column_letter(3 + w)
        gr.cell(k, 3 + w, "=" + "+".join(f"$B${i}*N({col}{i})" for i in range(5, k))).number_format = LEI
    gr.column_dimensions["A"].width = 55
    gr.column_dimensions["B"].width = 16

    OUT.mkdir(exist_ok=True)
    path = OUT / f"Propunere_financiara_BILZI_{L['fisier']}.xlsx"
    wb.save(path)
    return path


if __name__ == "__main__":
    LICITATII["dofteana"]["distanta"] = "cca. 190 km (Buzău – Focșani – Onești – Dofteana)"
    LICITATII["c4"]["distanta"] = "cca. 115 km (Buzău – București)"
    LICITATII["dofteana"]["saptamani"] = 9
    LICITATII["c4"]["saptamani"] = 7
    import oferta_financiara as ofin
    import analize_bilzi
    LICITATII["c4"]["analize"] = analize_bilzi.C4
    LICITATII["c4"]["indirecte"] = 0.03
    LICITATII["dofteana"]["analize"] = analize_bilzi.scaleaza_manopera(
        analize_bilzi.DOFTEANA, LICITATII["dofteana"]["f3"], analize_bilzi.TINTE_MANOPERA_DOFTEANA)
    LICITATII["c4"]["durata_scurt"] = "45 zile calendaristice (7 săptămâni)"
    LICITATII["dofteana"]["durata_scurt"] = "60 zile calendaristice (9 săptămâni)"
    # repartizarea procentuală pe săptămâni, pe capitole (în ordinea din F3)
    LICITATII["c4"]["grafic"] = [[0.15, 0.20, 0.20, 0.20, 0.15, 0.10, 0.0]]
    LICITATII["dofteana"]["grafic"] = [
        [0.40, 0.40, 0.20, 0, 0, 0, 0, 0, 0],          # demontări
        [0, 0.25, 0.35, 0.30, 0.10, 0, 0, 0, 0],       # reparații lemn
        [0, 0, 0.15, 0.20, 0.20, 0.20, 0.15, 0.10, 0],  # învelitoare
        [0, 0, 0, 0, 0, 0.30, 0.40, 0.30, 0],          # lambriu, pazii
        [0.20, 0, 0, 0, 0, 0, 0, 0.40, 0.40],          # paratrăsnet, firidă
        [0, 0, 0, 0, 0, 0, 0.40, 0.40, 0.20],          # tinichigerie, pluvial
        [0, 0, 0.20, 0.30, 0.30, 0.20, 0, 0, 0],       # coșuri
    ]
    # ADV1558636 – republicarea Dofteana (DSBC 2026 III, 20.11.2026): aceeași documentație (Deviz 8076/1/22.07.2026,
    # același F3, aceeași valoare estimată); caiet nr. 11858/16.11.2026, contract fără limita 31.12.2026.
    import copy
    D3 = copy.deepcopy(LICITATII["dofteana"])
    D3.update({
        "fisier": "ADV1558636_Parc_Dofteana_III",
        "titlu": "Înlocuire acoperiș clădire Parc Dofteana – DSBC 2026 III",
        "nr_anunt": "ADV1558636",
        "data_anunt": "20.11.2026",
        "termen": "26.11.2026",
        "ora": "08:00",
        "durata": "60 (șaizeci) de zile calendaristice – 2 (două) luni de la ordinul de începere (caietul de sarcini "
                  "nr. 11858/16.11.2026 și art. 5.1 din proiectul de contract); începerea lucrărilor în maximum 72 de "
                  "ore de la ordinul de începere (art. 8.1)",
        "depunere": ("e-mail anunturi.seap@bacau.rosilva.ro până la 26.11.2026 ora 08:00 + valoarea ofertei publicată "
                     "în catalogul SICAP (CPV 45261910-6, denumirea din anunț) până la aceeași oră"),
    })
    D3["note_f3"] = D3["note_f3"] + (" Republicare ADV1558636 (DSBC 2026 III): F3 identic cu cel din ADV1549806 "
                                     "(verificat pe pagina devizului 02 – 600 mp, 175,82 ml); garanție de bună "
                                     "execuție 5% (art. 6.1 din contract).")
    LICITATII["dofteana3"] = D3
    LICITATII["romeo"]["analize"] = analize_bilzi.ROMEO
    LICITATII["loto"]["analize"] = analize_bilzi.LOTO
    LICITATII["cfr_galati"]["analize"] = analize_bilzi.CFR_GALATI
    LICITATII["cfr_mangalia"]["analize"] = analize_bilzi.CFR_MANGALIA
    LICITATII["cfr_sibiu"]["analize"] = analize_bilzi.CFR_SIBIU
    LICITATII["salubris"]["analize"] = analize_bilzi.SALUBRIS
    LICITATII["cfr_sibiu_admin"]["analize"] = analize_bilzi.CFR_SIBIU_ADMIN
    LICITATII["cfr_sibiu_admin"]["depunere"] = ("online în SEAP (DUAE + ofertă + garanție de participare 3.000 lei); "
                                                "vizionare obligatorie")
    LICITATII["cfr_sibiu_admin"]["grafic"] = [[0.08, 0.14, 0.15, 0.15, 0.14, 0.12, 0.10, 0.07, 0.05]]
    LICITATII["cfr_sibiu"]["depunere"] = "online în SEAP (DUAE + ofertă + garanție de participare 600 lei); vizionare obligatorie"
    LICITATII["salubris"]["depunere"] = "ofertă postată în SEAP și, la aceeași oră, e-mail la office@salubris.ro"
    LICITATII["cfr_sibiu"]["grafic"] = [[0.20, 0.30, 0.25, 0.15, 0.10]]
    LICITATII["salubris"]["grafic"] = [[0.65, 0.35]]
    LICITATII["romeo"]["depunere"] = "e-mail anunturi.seap@bacau.rosilva.ro + prețul publicat în catalogul SICAP"
    LICITATII["loto"]["depunere"] = "e-mail achizitii.publice@loto.ro, cu referire la ADV1556660"
    LICITATII["cfr_galati"]["depunere"] = "online în SEAP (DUAE + ofertă + garanție de participare 1.049 lei)"
    LICITATII["cfr_mangalia"]["depunere"] = "online în SEAP (DUAE + ofertă + garanție de participare 1.720 lei); vizită obligatorie"
    LICITATII["romeo"]["grafic"] = [[0.15, 0.30, 0.30, 0.20, 0.05]]
    LICITATII["loto"]["grafic"] = [[0.30, 0.45, 0.25]]
    LICITATII["cfr_galati"]["grafic"] = [[0.10, 0.15, 0.15, 0.15, 0.15, 0.10, 0.10, 0.05, 0.05]]
    LICITATII["cfr_mangalia"]["grafic"] = [[0.10, 0.15, 0.15, 0.15, 0.15, 0.10, 0.10, 0.05, 0.05]]
    LICITATII["bjt_tulcea"]["analize"] = analize_bilzi.BJT_TULCEA
    LICITATII["bjt_tulcea"]["depunere"] = "e-mail bjt@tulcealibrary.ro (PDF semnat), vizită recomandată"
    LICITATII["bjt_tulcea"]["grafic"] = [[0.6, 0.4]]
    LICITATII["dgaspc_mangalia"]["analize"] = analize_bilzi.DGASPC_MANGALIA
    LICITATII["dgaspc_mangalia"]["depunere"] = "e-mail office@dgaspc-ct.ro / achizitii@dgaspc-ct.ro sau la sediul DGASPC Constanța"
    LICITATII["dgaspc_mangalia"]["grafic"] = [[0.12, 0.18, 0.20, 0.20, 0.18, 0.12]]
    resurse = ofin.citeste_resurse()
    OUT.mkdir(exist_ok=True)
    chei = sys.argv[1:] or list(LICITATII)
    for k in chei:
        L = LICITATII[k]
        adaos, calc = ofin.alege_adaos(L, resurse, indirecte=L.get("indirecte", 0.05))
        L["calc"] = calc
        p = ofin.genereaza(L, FIRMA, resurse, OUT / f"Propunere_financiara_BILZI_{L['fisier']}.xlsx", adaos,
                           L.get("indirecte", 0.05), calc)
        print(genereaza_docx(k))
        print(p, f"adaos {adaos:.1%}", f"total {calc['total']:,.2f} / estimat {L['valoare']:,.2f}",
              f"= {calc['total'] / L['valoare']:.1%}", "directe", round(calc['directe']))
