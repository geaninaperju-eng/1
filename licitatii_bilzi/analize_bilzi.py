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
    (CAP_C4, "RPCI01C"): [("M30", 5.4), ("M31", 0.12), ("M03", 0.0065), ("M08", 0.35), ("M05", 0.12),
                          ("M06", 0.25), ("M32", 0.10), ("O02", 0.55), ("O01", 0.15), ("O03", 0.25),
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
