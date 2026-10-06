#!/usr/bin/env python3
"""Centralizatorul achizițiilor directe anterioare – C.A. Rosetti, dosar consultanță (CREDINVEST, contract 4470/24.07.2026).

Model: „03. Centralizator achizitii directe anterioare” din dosarul GAL Bozioru (varianta semnată).
Se adaugă la dosarul scanat și numerotat ca poziția 20, fila 47 (fără renumerotarea filelor 1–46).
"""
from pathlib import Path

import genereaza_dosare as g
from genereaza_dosare import B, P, compacteaza, new_doc, table

OUT = Path(__file__).resolve().parent / "output_etapa1" / "CA Rosetti"

d = {
    "uat": {"denumire": "Comuna C.A. Rosetti",
            "adresa": "Strada Principală, nr. 41, Sat C.A. Rosetti, Com. C.A. Rosetti, Jud. Buzău, CP 127120",
            "cif": "3662681", "telefon": "0238 731001", "email": "secretariat@primariacarosetti.ro"},
    "persoane": {"consilier_achizitii": None, "contabil": "STANCIU IULIANA", "reprezentant": "CRĂCIUN COSTEL"},
}
TITLU = ("„CREȘTEREA EFICIENȚEI ENERGETICE ȘI REDUCEREA COSTURILOR CU ENERGIA ELECTRICĂ PRIN REALIZAREA UNUI SISTEM "
         "FOTOVOLTAIC PENTRU AUTOCONSUM AFERENT STAȚIEI DE POMPARE A APEI POTABILE DIN COMUNA C.A. ROSETTI, JUDEȚUL BUZĂU”")

doc = new_doc()
g.header_uat(doc, d)
P(doc, "Nr. ", None, " / ", None, ".10.2026", after=6)
P(doc, B("CENTRALIZATORUL ACHIZIȚIILOR DIRECTE ANTERIOARE"), align="c", size=13)
P(doc, "Proiect ", TITLU, align="c", after=0)
P(doc, "Contract de finanțare nr. C 36020808218621003741 / 03.07.2026", align="c", after=8)
rows = [
    ["1", "Servicii elaborare documentații tehnice (SF și PT) – CPV 71241000-9", [None], ["Contract / comandă nr. ", None, " / 23.12.2025"],
     "16.539,12", [None], "I.1", "Neaplicabil – achiziție anterioară semnării contractului de finanțare, cheltuială neeligibilă, finanțată din bugetul local"],
    ["2", "Proiectare și execuție lucrări – sistem fotovoltaic pentru autoconsum aferent stației de pompare a apei potabile – CPV 45261215-4",
     "S.C. NEXT LEVEL BUSINESS S.R.L., CUI RO31945829", "Contract de execuție lucrări nr. 4249/14.07.2026 (achiziție directă SEAP DA40809147/13.07.2026)",
     "406.464,72", "491.822,31", "II.3.1", ["Avizat prin Formularul 2 nr. ", None, " / ", None]],
    ["", B("TOTAL achiziții directe anterioare"), "", "", B("423.003,84"), [None], "", ""],
]
table(doc, ["Nr. crt.", "Obiectul achiziției / cod CPV", "Operator economic", "Document justificativ (nr./data)",
            "Valoare fără TVA (lei)", "Valoare cu TVA (lei)", "Poziție PAP", "Stadiu avizare AFIR"], rows, size=8.5)
P(doc, "")
P(doc, "Achiziția care face obiectul prezentului dosar – „Servicii de consultanță în managementul proiectului” (CPV 79400000-8 Consultanță "
       "în afaceri și în management și servicii conexe; contract nr. 4470/24.07.2026 încheiat cu CREDINVEST CONSULTING S.R.L., achiziție "
       "directă SEAP DA40872973/23.07.2026, 48.500,00 lei fără TVA, poziția II.2.1 din PAP nr. 4151/09.07.2026, avizat prin Formularul 2 "
       "nr. 80/10.07.2026) – nu este similară, din punct de vedere al scopului, naturii și obiectului, cu achizițiile directe anterioare din "
       "cadrul proiectului; nu a fost divizată în scopul evitării aplicării procedurilor prevăzute de Legea nr. 98/2016, iar valoarea "
       "cumulată se încadrează în pragurile prevăzute la art. 7 alin. (5) din lege.", align="j")
P(doc, "")
g.semnaturi(doc, d)
g.footer_no(doc, "20")
for sec in doc.sections:
    sec.footer.paragraphs[0].runs[0].text = "Doc. 20 | Fila 47"

OUT.mkdir(parents=True, exist_ok=True)
p = OUT / "20. Centralizator achizitii directe anterioare - Rosetti consultanta (fila 47).docx"
doc.save(p)
compacteaza(p)
print(p.name, p.stat().st_size)
