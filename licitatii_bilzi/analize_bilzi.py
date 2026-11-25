"""Consumuri de resurse pe unitatea fiecărui articol din F3 (analize de preț), pe licitații.

Cheia: (capitol F3, cod articol) – exact ca în LICITATII[...]["f3"]. Valoarea: listă (cod resursă, consum / UM articol).
Consumurile sunt estimate după norme uzuale (Indicatoarele de norme de deviz C, RPC, TRA) și practica BILZI;
se pot ajusta în foaia „Analiza” din Excel. Codurile de resurse sunt în preturi/resurse.csv.
"""

# ---------------------------------------------------------------- C4 Cantacuzino (pav. C4, cazarma 3589)
CAP_C4 = "Pavilion C4 – antemăsurătoare (Anexa 1)"
C4 = {
    # desfacere învelitoare țiglă solzi/olane pe șipci, cu sortarea țiglelor recuperabile
    (CAP_C4, "RPCT26B1"): [("O02", 0.15), ("O03", 0.22), ("T02", 0.0030)],
    # învelitoare din țiglă solzi refăcută: țigle recuperate + 25% țigle noi, șipci noi, folie unde e cazul
    (CAP_C4, "RPCI01C"): [("M30", 5.4), ("M31", 0.06), ("M03", 0.0065), ("M08", 0.35), ("M05", 0.12),
                          ("M06", 0.25), ("M32", 0.10), ("O02", 0.45), ("O01", 0.10), ("O03", 0.20),
                          ("U01", 0.03), ("T01", 1.20)],
    # reparații țigle profilate, coame în mortar de ciment, fără astereală
    (CAP_C4, "RPCI05XB"): [("M30", 6.0), ("M33", 6.0), ("O02", 0.60), ("O03", 0.30), ("T01", 0.50)],
    # jgheab semirotund D=12,5 din tablă zincată 0,5 confecționat pe șantier
    (CAP_C4, "RCSI18A"): [("M34", 0.38), ("M35", 1.60), ("M36", 0.06), ("M37", 0.15), ("O02", 1.10), ("O03", 0.20),
                          ("U07", 0.30)],
    # panou parazăpadă metalic cu ancore în căpriori
    (CAP_C4, "RPDE13A"): [("M38", 1.0), ("M39", 2.0), ("O02", 0.45), ("U02", 0.10)],
    # inele din bandă de aluminiu 1×20 mm
    (CAP_C4, "RPIZD31B"): [("M40", 1.10), ("M41", 2.0), ("O02", 0.15)],
    # plasă metalică de protecție pe cadru din oțel, cu ușă
    (CAP_C4, "W1C11B1"): [("M42", 1.10), ("M43", 6.0), ("M44", 0.25), ("M45", 0.20), ("O06", 1.00), ("O03", 0.25),
                          ("U05", 0.30), ("U02", 0.10)],
    # manipulare materiale cu macaraua, 0,5–1 t / transport
    (CAP_C4, "TRB22E6B"): [("U04", 0.25), ("O03", 0.50)],
    # schelă metalică de susținere 3–6 m (închiriere ~1,5 luni, montaj și demontaj)
    (CAP_C4, "CB41B1"): [("U06", 1.5), ("O08", 0.40)],
}

# ---------------------------------------------------------------- Parc Dofteana (Romsilva DS Bacău)
D1 = "02 Demontări"
D2 = "03 Reparații suport lemn"
D3 = "04 Înlocuire învelitoare"
D4 = "05 Montaj intrados streașină, pazii din lemn"
D5 = "06 Lucrări la instalații electrice și de protecție la trăsnet"
D6 = "07 Tinichigerie și pluvial"
D7 = "08 Coșuri de fum"
DOFTEANA = {
    (D1, "DEM001"): [("O02", 0.25), ("O03", 0.35)],
    (D1, "DEM002"): [("O02", 0.15), ("O03", 0.10)],
    (D1, "DEM003"): [("O05", 6.0), ("O03", 6.0)],
    (D1, "DEM004"): [("O02", 1.5), ("O03", 1.0)],
    (D1, "DEM005"): [("O04", 0.5), ("O03", 0.6)],
    (D1, "DEM006"): [("O01", 0.30), ("O03", 0.20)],
    (D1, "DEM007"): [("O01", 0.25), ("O03", 0.20), ("T02", 0.0045)],
    (D1, "DEM008"): [("O05", 3.0)],
    (D1, "DEM09"): [("O03", 1.6)],
    (D2, "REP001"): [("M01", 1.10), ("M04", 25.0), ("M50", 3.0), ("O01", 22.0), ("O03", 6.0), ("U01", 2.0)],
    (D2, "REP002"): [("M01", 0.08), ("M04", 3.0), ("M50", 1.0), ("M51", 1.0), ("O01", 12.0), ("O03", 3.0),
                     ("U01", 1.0)],
    (D2, "REP003"): [("U03", 0.18), ("O07", 0.236), ("M56", 0.01)],
    (D2, "REP004"): [("M52", 0.30), ("M53", 0.20), ("M54", 0.25), ("O07", 0.90)],
    (D3, "ACS001"): [("M02", 0.0275), ("M05", 0.10), ("M06", 0.30), ("O01", 0.40), ("O03", 0.10),
                     ("U01", 0.03), ("T01", 1.0)],
    (D3, "ACS002"): [("M03", 0.0095), ("M05", 0.10), ("M06", 0.10), ("O01", 0.30), ("U01", 0.02), ("T01", 0.4)],
    (D3, "ACS004"): [("M08", 1.12), ("M09", 0.10), ("O02", 0.08)],
    (D3, "ACS005"): [("M10", 1.10), ("M11", 4.0), ("O02", 0.65), ("O03", 0.25), ("T01", 3.0)],
    (D3, "ACS006"): [("M12", 1.0), ("M13", 1.0), ("M14", 0.15), ("O02", 0.45)],
    (D3, "ACS007"): [("M20", 0.62), ("M09", 0.05), ("O02", 0.55)],
    (D3, "ACS008"): [("M20", 0.40), ("M09", 0.05), ("O02", 0.30)],
    (D3, "ACS009"): [("M20", 0.40), ("M09", 0.05), ("O02", 0.35)],
    (D3, "ACS010"): [("M15", 1.2), ("O02", 0.30)],
    (D3, "ACS012"): [("M20", 0.45), ("M09", 0.05), ("M17", 0.10), ("O02", 0.50)],
    (D4, "ACS001"): [("M18", 1.20), ("M19", 22.0), ("M52", 0.15), ("O01", 0.90), ("O03", 0.15), ("U02", 0.10)],
    (D4, "ACS003"): [("M55", 1.05), ("M19", 4.0), ("O01", 0.35)],
    (D5, "ACS012"): [("M60", 25.0), ("M61", 45.0), ("M62", 1.0), ("M63", 1.0), ("O05", 12.0), ("O03", 6.0)],
    (D5, "ACS013"): [("M64", 1.0), ("O05", 8.0)],
    (D6, "TIN001"): [("M21", 1.0), ("M22", 1.5), ("M17", 0.03), ("O02", 0.35)],
    (D6, "TIN002"): [("M23", 1.0), ("O02", 0.25)],
    (D6, "TIN003"): [("M24", 1.0), ("O02", 0.30)],
    (D6, "TIN004"): [("M25", 1.0), ("O02", 0.25)],
    (D6, "TIN005"): [("M26", 1.0), ("O02", 0.15)],
    (D6, "TIN006"): [("M27", 1.0), ("O02", 0.15)],
    (D6, "TIN007"): [("M28", 1.0), ("O02", 0.15)],
    (D6, "TIN008"): [("M29", 1.0), ("O02", 0.10)],
    (D6, "TIN009"): [("M65", 1.0), ("O03", 0.50)],
    (D7, "COS000"): [("M70", 5.0), ("M71", 5.0), ("O04", 2.2), ("O03", 1.2)],
    (D7, "COS001"): [("M71", 2.5), ("M72", 0.30), ("O04", 1.0), ("O03", 0.3)],
    (D7, "COS002"): [("M73", 1.0), ("O02", 4.0), ("O03", 2.0)],
    (D7, "COS003"): [("M75", 1.30), ("M09", 0.20), ("O02", 2.5)],
    (D7, "COS004"): [("M20", 2.40), ("M09", 0.20), ("M17", 0.30), ("O02", 3.0)],
}


# Ore de manoperă din C7 al proiectantului (Deviz 8076/1/22.07.2026) × 1,15 rezervă; electricianul majorat la 16 h
# (paratrăsnet + firidă). Analizele Dofteana se scalează pe meserii la aceste totaluri.
TINTE_MANOPERA_DOFTEANA = {"O01+O02": (406.90 + 90.52) * 1.15, "O03": (38.00 + 206.08 + 180.74) * 1.15,
                           "O04": (19.14 + 16.00) * 1.15, "O05": 16.0}


def scaleaza_manopera(analize, f3, tinte):
    """Înmulțește consumurile de manoperă pe meserii astfel încât totalul de ore să fie egal cu ținta."""
    cant = {(cap, cod): q for cap, arts in f3 for cod, _, _, q in arts}
    for grup, tinta in tinte.items():
        coduri = grup.split("+")
        total = sum(cant[k] * c for k, lst in analize.items() for rc, c in lst if rc in coduri)
        f = tinta / total
        for k, lst in analize.items():
            analize[k] = [(rc, round(c * f, 5) if rc in coduri else c) for rc, c in lst]
    return analize


# ---------------------------------------------------------------- licitații 12.11.2026
# Cabana Romeo (ADV1556879) – F3 DMH937: articolele „asim – montat” sunt doar manoperă, materialele au rânduri proprii.
CAP_ROMEO = "Deviz DMH937 – Înlocuire învelitoare acoperiș Cabana Romeo"
ROMEO = {(CAP_ROMEO, k): v for k, v in {
    "RPCT26D1": [("O02", 0.08), ("O03", 0.22), ("T02", 0.004)],
    "CE18A1": [("M05", 0.05), ("O01", 0.30), ("O03", 0.05), ("U01", 0.03)],
    "3270480": [("N13", 1.0)],
    "CE20A1": [("M03", 0.006), ("M05", 0.05), ("M06", 0.10), ("O01", 0.25), ("O03", 0.05), ("U01", 0.02)],
    "CE05A1": [("O02", 0.40), ("O03", 0.10), ("U02", 0.05)],
    "3270481": [("N01", 1.0)],
    "7299982": [("M08", 1.0), ("O02", 0.03)],
    "7299995": [("M09", 1.0)],
    "7299992": [("N08", 1.0), ("O02", 1.0)],
    "7299996": [("N07", 0.5), ("M09", 3.0), ("O02", 0.15)],
    "CE08A%9": [("O02", 0.35)],
    "3270482": [("N02", 1.0)],
    "3270483": [("N03", 1.0)],
    "7299994": [("N09", 1.0)],
    "CE08A%6": [("O02", 0.20)],
    "7299997": [("N04", 0.5)],
    "CE08A%7": [("O02", 0.25)],
    "3270468": [("N05", 1.0)],
    "CE08A%8": [("O02", 0.40)],
    "3270484": [("N06", 1.0)],
    "7299993": [("N10", 1.0)],
    "CE13A1": [("O02", 0.40), ("O03", 0.10)],
    "3270485": [("M21", 1.0)],
    "7299989": [("M22", 1.0)],
    "7299987": [("N11", 1.0)],
    "7299986": [("M25", 1.0)],
    "3270487": [("N12", 1.0)],
    "CE14A1": [("O02", 0.30), ("O03", 0.10)],
    "3270486": [("M23", 1.0)],
    "7299988": [("M29", 1.0)],
    "7299984": [("M26", 1.0)],
    "7299985": [("M27", 1.0)],
    "TRA02A40": [("T03", 1.0)],
}.items()}

# Terasa Loto Bacău (ADV1556660) – 95 mp, sistem bistrat P4 + PA4,5 ardezie peste V3 de reparații
CAP_LOTO = "Deviz estimativ – terasa agenției 04-013 Bacău"
LOTO = {(CAP_LOTO, k): v for k, v in {
    "1": [("O03", 0.35), ("O02", 0.10)],
    "2": [("N14", 1.15), ("N18", 0.20), ("O02", 0.15), ("O03", 0.05)],
    "3": [("N19", 60.0), ("O04", 0.50), ("O03", 0.50)],
    "4": [("N17", 0.35), ("O02", 0.04)],
    "5": [("N15", 1.15), ("N18", 0.25), ("O02", 0.22), ("O03", 0.08)],
    "6": [("N16", 1.15), ("N18", 0.25), ("O02", 0.25), ("O03", 0.08)],
    "7": [("N20", 1.0), ("O02", 2.0)],
    "8": [("N21", 1.0), ("O02", 1.0)],
    "9": [("M20", 0.45), ("M09", 4.0), ("N10", 0.10), ("O02", 0.50)],
    "10": [("N26", 1.0)],
    "11": [("O03", 16.0)],
    "12": [("T02", 0.5), ("O03", 6.0)],
}.items()}

# CFR Galați – Revizia Vagoane (SCN1181822) – 459 mp membrană cu ardezie, 114 panouri PV demontate/remontate
CAP_CFR_GL = "Lista cantităților de lucrări – hidroizolație fostă CT Revizia Vagoane"
CFR_GALATI = {(CAP_CFR_GL, k): v for k, v in {
    "1": [("O05", 0.15), ("O02", 0.35), ("O03", 0.35)],
    "2": [("O03", 0.30), ("O02", 0.05)],
    "3": [("N22", 25.0), ("O04", 1.5), ("O03", 0.5)],
    "4": [("N22", 30.0), ("O04", 2.5), ("O03", 1.0)],
    "5": [("N17", 0.35), ("O02", 0.04), ("O03", 0.02)],
    "6": [("N16", 1.15), ("N18", 0.25), ("O02", 0.25), ("O03", 0.08)],
    "7": [("N23", 1.0), ("O05", 0.35), ("O02", 0.50), ("O03", 0.40)],
    "8": [("U04", 1.0)],
    "9": [("T02", 0.4), ("O03", 2.0)],
}.items()}

# CFR Mangalia – Vila C (SCN1181778) – antemăsurătoare estimată pe 236 mp (de confirmat la vizita obligatorie)
CAP_CFR_MG = "Antemăsurătoare estimată (de confirmat la vizita obligatorie)"
CFR_MANGALIA = {(CAP_CFR_MG, k): v for k, v in {
    "D01": [("O05", 16.0), ("O02", 16.0), ("O03", 16.0)],
    "D02": [("O02", 0.12), ("O03", 0.18)],
    "D03": [("O02", 0.10), ("O03", 0.05)],
    "D04": [("O01", 0.10), ("O03", 0.10)],
    "A01": [("M02", 0.0275), ("M05", 0.10), ("M06", 0.30), ("O01", 0.40), ("O03", 0.10), ("U01", 0.03)],
    "A02": [("N25", 1.05), ("O01", 0.12), ("O03", 0.05)],
    "A03": [("M08", 1.12), ("M09", 0.10), ("O02", 0.08)],
    "A04": [("M03", 0.0095), ("M05", 0.10), ("M06", 0.10), ("O01", 0.30), ("U01", 0.02)],
    "A05": [("N24", 1.10), ("M09", 7.0), ("O02", 0.45), ("O03", 0.10), ("U02", 0.05), ("T03", 0.006)],
    "A06": [("N02", 0.50), ("N03", 0.07), ("N09", 0.20), ("M09", 4.0), ("O02", 0.35)],
    "A07": [("N06", 0.55), ("M09", 3.0), ("O02", 0.40)],
    "A08": [("N05", 0.55), ("M09", 4.0), ("O02", 0.25)],
    "P01": [("M21", 1.05), ("M22", 1.6), ("N11", 0.25), ("M24", 0.05), ("M25", 0.10), ("O02", 0.40)],
    "P02": [("M23", 1.0), ("M29", 0.6), ("M27", 0.15), ("O02", 0.30)],
    "S01": [("N23", 20.0), ("O05", 16.0), ("O02", 16.0), ("O03", 8.0)],
    "E01": [("T02", 1.0), ("O03", 8.0)],
}.items()}

# ---------------------------------------------------------------- licitații 13.11.2026
# CFR Sibiu – baracă metalică Depoul Sibiu (SCN1182052): 350 mp azbociment → tablă cutată 0,50 mm; echipă de 4 la 290 km
CAP_CFR_SB = "Lista de cantități – Anexa la CS ETT54/3/1/318/2026"
CFR_SIBIU = {(CAP_CFR_SB, k): v for k, v in {
    # 3 drumuri dus-întors (vizionare, mobilizare, demobilizare) × 580 km; 4 oameni × 10 nopți; plan azbest ITM
    "1": [("R11", 1740.0), ("R10", 40.0), ("R12", 1.0), ("O02", 4.0), ("O03", 8.0)],
    "2": [("O02", 0.15), ("O03", 0.25), ("R07", 0.035), ("R08", 0.15)],
    "3": [("R06", 1.0), ("O03", 0.10)],
    "4": [("R09", 4.5), ("O03", 6.0)],  # plăci ondulate 6 mm ≈ 12–13 kg/mp → 4,5 t
    "5": [("O06", 0.03), ("O03", 0.03), ("M44", 0.02)],
    "6": [("M43", 60.0), ("M44", 5.0), ("M45", 2.0), ("O06", 16.0), ("O03", 8.0), ("U05", 8.0)],
    "7": [("R01", 1.12), ("O02", 0.25), ("O03", 0.10), ("U02", 0.08), ("U04", 0.012)],
    "8": [("R02", 1.0)],
    "9": [("R05", 1.05), ("O02", 0.05)],
    "10": [("R04", 0.40), ("R03", 0.15), ("M09", 4.0), ("N10", 0.10), ("O02", 0.35)],
    "11": [("M21", 0.6), ("M22", 1.0), ("M23", 0.45), ("M29", 0.3), ("M25", 0.05), ("M27", 0.1), ("O02", 0.45)],
    "12": [("O03", 16.0), ("T03", 0.5)],
}.items()}

# Salubris Iași (ADV1557141) – terasă 110 mp, membrană PA 4,2 mm -15 °C; deplasare și cazare în poziția 1
CAP_SAL = "Lista cantităților de lucrări – solicitarea Salubris nr. 33347/12.11.2026"
SALUBRIS = {(CAP_SAL, k): v for k, v in {
    # 1 drum dus-întors cca. 605 km (5,5 km/mp × 110), 3 oameni × 3 nopți (≈9), evacuare 0,33 t deșeuri
    "1": [("O03", 0.25), ("O02", 0.05), ("R11", 5.5), ("R10", 0.082), ("T03", 0.003)],
    "2": [("N19", 60.0), ("O04", 0.50), ("O03", 0.50)],
    "3": [("N17", 0.30), ("O02", 0.04)],
    "4": [("R13", 1.12), ("N18", 0.25), ("O02", 0.25), ("O03", 0.08)],
    "5": [("O04", 0.50), ("R14", 0.50)],
    "6": [("M34", 0.20), ("N10", 0.15), ("M09", 3.0), ("O02", 0.40)],
    "7": [("N16", 1.0), ("N17", 0.2), ("O02", 1.5)],
}.items()}

# ---------------------------------------------------------------- licitații 16.11.2026
# CFR Sibiu – acoperiș clădire administrativă Depoul Sibiu (CN1098200): 1.000 mp țiglă ceramică → țiglă metalică
# 0,50 mm RAL 8017; echipă de 5 la 290 km, 60 de zile (cazare luni–vineri, acasă în weekend)
CAP_CFR_SBA = "Lista de cantități – Anexa la CS ETT54/3/1/319/2026"
CFR_SIBIU_ADMIN = {(CAP_CFR_SBA, k): v for k, v in {
    # 10 drumuri dus-întors × 580 km; 5 oameni × 40 nopți; schelă la streașină, macara, linii de ancorare
    "1": [("R11", 5800.0), ("R10", 200.0), ("R17", 1.0), ("U06", 400.0), ("O08", 24.0), ("U04", 16.0),
          ("O02", 8.0), ("O03", 16.0)],
    "2": [("O02", 0.12), ("O03", 0.18)],
    "3": [("O03", 0.15), ("U04", 0.005)],
    "4": [("O01", 0.10), ("O03", 0.10)],
    # coș cca. 1,5 mc zidărie deasupra/în pod; închiderea golului în învelitoare
    "5": [("O04", 6.0), ("O03", 8.0), ("R16", 0.03), ("M05", 0.5)],
    "6": [("M04", 80.0), ("M50", 10.0), ("R16", 0.5), ("O01", 40.0), ("O03", 16.0)],
    "7": [("R16", 0.025), ("M04", 0.4), ("M06", 0.15), ("O01", 0.80), ("O03", 0.40), ("U01", 0.05)],
    "8": [("O01", 0.30), ("O03", 0.20)],
    "9": [("R16", 0.012), ("M02", 0.02), ("M05", 0.2), ("M06", 0.2), ("O01", 1.2), ("O03", 0.4), ("U01", 0.1)],
    "10": [("M52", 0.30), ("M53", 0.15), ("O07", 0.80), ("R14", 0.10)],
    "11": [("M08", 1.12), ("M09", 0.10), ("O02", 0.08)],
    "12": [("M03", 0.0095), ("M05", 0.10), ("M06", 0.10), ("O01", 0.30), ("U01", 0.02)],
    "13": [("N24", 1.10), ("M09", 7.0), ("O02", 0.40), ("O03", 0.10), ("U02", 0.05), ("T01", 1.6)],
    "14": [("N04", 0.30), ("N05", 0.25), ("M09", 4.0), ("O02", 0.30)],
    "15": [("N02", 0.50), ("N03", 0.07), ("N09", 0.20), ("M09", 4.0), ("O02", 0.35)],
    "16": [("N06", 0.55), ("N08", 0.15), ("M09", 4.0), ("N10", 0.05), ("O02", 0.45)],
    "17": [("R15", 1.0), ("O03", 1.0)],
    "18": [("O03", 24.0)],
}.items()}

# ---------------------------------------------------------------- licitații 18.11.2026
# Biblioteca Județeană „Panait Cerna” Tulcea (ADV1558035): 217 mp hidroizolație bituminoasă în 2 straturi
CAP_BJT = "Formular 3 – Anexa nr. 1 la anunțul de publicitate nr. 2579/17.11.2026"
BJT_TULCEA = {(CAP_BJT, k): v for k, v in {
    "1": [("O03", 0.15), ("O02", 0.03), ("H25", 0.05)],
    "2": [("H24", 0.30), ("O02", 0.05)],
    # strat 1: membrană SBS 4 kg/mp cu inserție din împâslitură de fibră de sticlă, termosudată, suprapuneri 10–15%
    "3": [("H12", 1.15), ("N18", 0.25), ("O02", 0.25), ("O03", 0.08)],
    # strat 2: membrană SBS cu ardezie (strat hidroizolant cauciucat de finisaj)
    "4": [("H13", 1.12), ("N18", 0.25), ("O02", 0.27), ("O03", 0.08)],
    # 2 drumuri dus-întors Buzău–Tulcea (2 × 460 km), 3 oameni × 4 nopți, material mărunt, evacuare deșeuri
    "5": [("R11", 920.0), ("R10", 12.0), ("N26", 1.0), ("N10", 4.0), ("T03", 1.0), ("O03", 6.0)],
}.items()}

# DGASPC Constanța – Centrul de Zi și Recuperare Mangalia (ADV1558019): termo-hidroizolație terasă 650 mp
# echipă de 6 la cca. 270 km, 40 de zile (cazare luni–vineri), pozițiile x.y = materialele separate din F3
CAP_DGASPC = "F3 – Lista cu cantități de lucrări (Hidroizolație Mangalia – estimare hidroizolație)"
DGASPC_MANGALIA = {(CAP_DGASPC, k): v for k, v in {
    # desfacere hidroizolație + termoizolație veche; include deplasarea și cazarea echipei (6 oameni × 30 nopți,
    # 6 drumuri × 540 km) și protecțiile de șantier
    "1": [("O03", 0.40), ("O02", 0.06), ("R14", 0.03), ("R10", 0.277), ("R11", 5.0), ("R17", 0.0015)],
    "2": [("O03", 0.18), ("O02", 0.02)],
    "3": [("O03", 4.0), ("U04", 0.08)],
    "4": [("R15", 1.0)],
    "5": [("O03", 0.06), ("H11", 0.06)],
    "6": [("O02", 0.06), ("O03", 0.02)],
    "6.4": [("H24", 1.0)],
    "7": [("N18", 0.25), ("O02", 0.15), ("O03", 0.05)],
    "7.4": [("H02", 1.0)],
    "7.5": [("H01", 1.0)],
    "8": [("H04", 0.11), ("O02", 0.12), ("O03", 0.04)],
    "8.1": [("H03", 1.0)],
    "9": [("O03", 0.02)],
    "9.2": [("H05", 1.0)],
    "10": [("H06", 1.0), ("O04", 0.06), ("O03", 0.06)],
    # șapă de pantă din beton B400 cu fulgi PP: agregate, fibre, turnare mecanizată (cimentul e la 11.5)
    "11": [("H07", 1.05), ("H08", 0.9), ("O04", 2.2), ("O03", 2.2), ("H10", 0.35)],
    "11.5": [("H09", 1.0)],
    "12": [("H12", 1.12), ("N18", 0.25), ("O02", 0.22), ("O03", 0.08)],
    "13": [("H13", 1.12), ("N18", 0.25), ("O02", 0.25), ("O03", 0.08)],
    "13.5": [("H02", 1.0)],
    "14": [("N18", 0.15), ("N17", 0.15), ("O02", 0.35), ("O03", 0.10)],
    "14.4": [("H13", 1.0)],
    "15": [("H26", 0.30), ("M09", 4.0), ("N10", 0.10), ("O02", 0.45), ("O03", 0.10)],
    "15.2": [("H14", 1.0)],
    "15.3": [("H15", 1.0)],
    "16": [("N18", 0.30), ("O02", 0.80)],
    "16.3": [("H02", 1.0)],
    "16.4": [("H16", 1.0)],
    "17": [("H17", 1.0), ("O02", 1.50), ("O04", 0.50), ("N22", 5.0)],
    "17.1": [("H18", 1.0)],
    "18": [("H19", 1.0), ("O02", 1.0)],
    "19": [("H20", 1.0), ("O02", 0.25)],
    "20": [("H21", 1.0), ("O02", 0.40), ("O03", 0.10)],
    "20.1": [("H22", 1.0)],
    "21": [("T01", 50.0)],
    "22": [("H23", 1.0), ("O06", 6.0), ("O03", 3.0), ("N10", 2.0)],
}.items()}

# ---------------------------------------------------------------- licitații 24.11.2026
# Universitatea Tehnică „Gheorghe Asachi” Iași – Lot 1: hidroizolația terasei căminului T20 (1.080 mp, 2 straturi SBS)
# echipă de 5 la cca. 300 km, 30 de zile: deplasarea (5 drumuri × 600 km) și cazarea (125 nopți) sunt în poziția 4
CAP_TU_D = "Lot 1 – Desfaceri"
CAP_TU_C = "Lot 1 – Lucrări de construcții-montaj"
TUIASI_T20 = {(CAP_TU_D, k): v for k, v in {
    "1": [("O02", 0.35), ("O03", 0.35)],
    "2": [("O02", 1.00), ("O03", 1.00)],
    "3": [("O02", 0.15), ("O03", 0.15)],
}.items()}
TUIASI_T20.update({(CAP_TU_C, k): v for k, v in {
    "1": [("H29", 25.0), ("O04", 1.50), ("O03", 0.80)],
    "2": [("H12", 1.15), ("N17", 0.30), ("N18", 0.25), ("O02", 0.35), ("O03", 0.15)],
    "3": [("N19", 30.0), ("O04", 1.00), ("O03", 0.60)],
    "4": [("O03", 0.12), ("O02", 0.03), ("H25", 0.02), ("H02", 0.15), ("N18", 0.02),
          ("R10", 0.12), ("R11", 2.80), ("R17", 0.0009)],
    "5": [("H17", 1.0), ("H16", 1.0), ("H20", 1.0), ("N10", 0.5), ("N18", 0.3), ("O02", 2.0), ("O03", 0.5)],
    "6": [("H27", 1.0), ("N18", 0.3), ("O02", 0.8), ("O03", 0.2)],
    # amorsare în două straturi
    "7": [("H24", 0.35), ("O02", 0.06), ("O03", 0.03)],
    # strat 1: membrană SBS ≥ 3 kg/mp fără ardezie, suprapuneri 10–15%; ridicarea materialelor pe terasă cu macaraua
    "8": [("H12", 1.15), ("N18", 0.25), ("O02", 0.22), ("O03", 0.08), ("T03", 0.0046), ("U04", 0.004)],
    # strat final: membrană SBS ≥ 4 mm cu ardezie
    "9": [("H13", 1.12), ("N18", 0.25), ("O02", 0.25), ("O03", 0.08), ("T03", 0.0058), ("U04", 0.004)],
    "10": [("H13", 0.30), ("N10", 0.30), ("N18", 0.20), ("O02", 0.60)],
    "11": [("H28", 1.0), ("N18", 0.20), ("O02", 0.60), ("O03", 0.10)],
    # glaf din tablă prevopsită 0,50 mm, desfășurată cca. 45 cm + suprapuneri 15 cm, 2 benzi de mastic PU la îmbinări
    "12": [("M20", 0.55), ("N10", 0.10), ("M09", 4.0), ("O02", 0.45), ("O03", 0.10)],
    # capac într-o apă din tablă cutată prevopsită, pe cadru metalic propriu (fără găurire în zidăria ghenei)
    "13": [("R01", 0.70), ("M43", 3.0), ("M44", 0.10), ("M09", 8.0), ("N10", 0.20), ("O06", 0.50), ("O02", 0.60),
           ("O03", 0.20), ("U05", 0.20)],
    "14": [("R15", 4.0), ("O03", 16.0)],
    "15": [("O03", 16.0), ("O02", 4.0), ("N26", 1.0)],
}.items()})

# ---------------------------------------------------------------- licitații 25.11.2026
# Școala Gimnazială Centrală Câmpina – hidroizolație acoperiș 30 mp + plafon hol 15 mp (echipă de 2–3, 110 km)
CAP_CP1 = "1. Reparația hidroizolației acoperișului – 30 mp"
CAP_CP2 = "2. Reparația plafonului holului – 15 mp"
CAP_CP3 = "3. Transport, deșeuri, predarea amplasamentului"
CAMPINA = {(CAP_CP1, k): v for k, v in {
    "1.1": [("O02", 0.25), ("O03", 0.30)],
    "1.2": [("N22", 3.0), ("O04", 0.15), ("O03", 0.10)],
    "1.3": [("N17", 0.30), ("O02", 0.05)],
    # membrană SBS cu ardezie, suprapuneri 10–15%, termosudare
    "1.4": [("H13", 1.12), ("N18", 0.30), ("O02", 0.35), ("O03", 0.10)],
    "1.5": [("K05", 1.0), ("N10", 0.10), ("O02", 0.30)],
}.items()}
CAMPINA.update({(CAP_CP2, k): v for k, v in {
    "2.1": [("H05", 1.20), ("O03", 0.15)],
    "2.2": [("O03", 0.60), ("O04", 0.20)],
    "2.3": [("K10", 0.30), ("O04", 0.20), ("O03", 0.10)],
    # tinci ≤ 1 cm ≈ 13 kg/mp mortar uscat
    "2.4": [("K06", 13.0), ("O04", 0.70), ("O03", 0.30)],
    "2.5": [("K07", 1.50), ("O04", 0.45)],
    "2.6": [("K08", 0.15), ("K09", 0.35), ("O07", 0.35)],
    "2.7": [("K12", 3.0), ("O03", 1.0)],
}.items()})
CAMPINA.update({(CAP_CP3, k): v for k, v in {
    # 3 drumuri Buzău – Câmpina dus-întors (3 × 220 km)
    "3.1": [("R11", 660.0)],
    "3.2": [("R15", 0.5), ("O03", 3.0)],
}.items()})

# Termocentrale Constanța – 600 mp membrană SBS 3 mm peste hidroizolația existentă, profile Z 150 ml, wassernas 75 ml
# echipă de 4 la 230 km, cca. 10 zile de lucru: 2 drumuri cu microbuzul și 40 de nopți de cazare (poziția 8)
CAP_TC = "Refacere hidroizolație – zona vestiare Secția Chimică"
TERMOCENTRALE = {(CAP_TC, k): v for k, v in {
    "1": [("O02", 0.04), ("O03", 0.05), ("N18", 0.03), ("N17", 0.05)],
    # profil Z: 0,65 m × 1,5 mm × 7,85 = 7,65 kg/ml (+1% pierderi), fasonare la abkant, 3 ancore/ml
    "2": [("K02", 7.73), ("K03", 1.0), ("K04", 3.0), ("O06", 0.15), ("O02", 0.15), ("O03", 0.10), ("U02", 0.10)],
    # wassernas din tablă prevopsită 0,50 mm, desfășurată cca. 35 cm + suprapuneri; 2 holșuruburi/ml
    "3": [("M20", 0.40), ("K04", 2.0), ("O02", 0.25), ("O03", 0.10), ("U02", 0.05)],
    "4": [("K01", 1.12), ("N18", 0.25), ("O02", 0.14), ("O03", 0.06)],
    "5": [("K05", 1.05), ("N10", 0.05), ("O02", 0.15)],
    "6": [("U04", 2.0), ("O03", 16.0)],
    "7": [("R15", 1.0), ("O03", 4.0)],
    "8": [("T01", 500.0), ("R11", 920.0), ("R10", 40.0)],
}.items()}
