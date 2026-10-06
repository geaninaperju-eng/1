#!/usr/bin/env python3
"""Completări pentru dosarul ETAPA I – Gălbinași, achiziție nacelă (procedură simplificată).

Modelul de structură: dosarul Etapa I Scutelnici (buldoexcavator), verificat de OJFIR Buzău la 29.09.2026.
Generează doar documentele care lipsesc sau trebuie corectate; restul există scanate și semnate.
Ieșire: output_etapa1/Galbinasi/
"""
from pathlib import Path

from docx.enum.text import WD_BREAK

import genereaza_dosare as g
from genereaza_dosare import B, P, compacteaza, new_doc, table

OUT = Path(__file__).resolve().parent / "output_etapa1" / "Galbinasi"

UAT = {
    "denumire": "PRIMĂRIA COMUNEI GĂLBINAȘI",
    "adresa": "Sat Gălbinași, Str. Unirii, nr. 71, Comuna Gălbinași, Județul Buzău, Cod poștal 127240",
    "tel": "Tel.: 0238 / 780.015, Fax: 0238 / 780.022",
}
OBIECT = "ACHIZIȚIE NACELĂ PENTRU DOTAREA SERVICIULUI DE ILUMINAT PUBLIC COMUNA GĂLBINAȘI"
CONTRACT_FIN = "C 36020808218621003751"
CPV = "42418000-9 – Utilaje de ridicare, de manipulare, de încărcare sau de descărcare (Rev.2)"
VALOARE = "443.421,60"
PRIMAR = "DUMITRU DRAGOMIR"
COMISIE = [
    ("CRISTEA CRISTINEL", "Viceprimar", "președinte al comisiei de evaluare"),
    ("SILIVESTRU ALINA-MAGDALENA", "Inspector superior – Compartimentul financiar-contabil, impozite și taxe", "membru al comisiei de evaluare"),
    ("BUZĂIANU-BARBU PETRONIA", "Consilier achiziții publice", "membru al comisiei de evaluare"),
    ("BRATU IONELA", "Consilier superior – Compartimentul financiar-contabil", "membru de rezervă al comisiei de evaluare"),
    ("GRIGORE TANTA", "Consilier superior – Compartimentul financiar-contabil", "membru de rezervă al comisiei de evaluare"),
]


def antet(doc):
    P(doc, "R O M Â N I A", align="c", after=0, size=10)
    P(doc, "JUDEȚUL BUZĂU", align="c", after=0, size=10)
    P(doc, B(UAT["denumire"]), align="c", after=0)
    P(doc, UAT["tel"], align="c", after=0, size=9)
    P(doc, "Sediu: " + UAT["adresa"], align="c", after=8, size=9)


def save(doc, name):
    OUT.mkdir(parents=True, exist_ok=True)
    p = OUT / f"{name}.docx"
    doc.save(p)
    compacteaza(p)
    return p


# ------------------------------------------------------------------ 00a Opis

def opis():
    doc = new_doc()
    antet(doc)
    P(doc, B("OPIS – DOSAR ACHIZIȚIE PUBLICĂ, ETAPA I (LANSAREA ACHIZIȚIEI)"), align="c", size=12, after=0)
    P(doc, "Procedură simplificată – „", OBIECT, "”", align="c", after=0, size=10)
    P(doc, "Contract de finanțare nr. ", CONTRACT_FIN, align="c", after=8, size=10)
    rows = [
        ("1", "Strategia de contractare nr. 9564/02.10.2026", ""),
        ("2", "Nota privind determinarea valorii estimate nr. 9565/02.10.2026", "2"),
        ("3", "Analiza rezonabilității valorii estimate nr. ______/02.10.2026, cu anexele: ofertele UTILBEN (nr. 208234/08.12.2025), "
              "ITALIA STAR (e-mail 17.12.2025), KSM (cod 0D3F7EB5CA/03.12.2025) și lista caracteristicilor tehnice din cererea de finanțare", ""),
        ("4", "Draft anunț de participare simplificat – export SEAP", "9"),
        ("5", "DUAE – export SEAP", ""),
        ("6", "Fișa de date a achiziției – export SEAP", "12"),
        ("7", "Formulare anexă pentru ofertanți", "7"),
        ("8", "Modelul contractului de furnizare produse", "16"),
        ("9", "Caietul de sarcini", "2"),
        ("10", "Dispoziția nr. 140/02.10.2026 privind aprobarea documentației de atribuire și numirea comisiei de evaluare", ""),
        ("11", "CV – Cristea Cristinel (președinte comisie)", ""),
        ("12", "CV – Silivestru Alina-Magdalena (membru comisie)", ""),
        ("13", "CV – Buzăianu-Barbu Petronia (membru comisie)", ""),
        ("14", "CV – Bratu Ionela (membru de rezervă)", ""),
        ("15", "CV – Grigore Tanta (membru de rezervă)", ""),
        ("16", "Declarația privind persoanele cu funcții de decizie nr. 9567/02.10.2026", "1"),
        ("17", "Declarația privind cuprinderea procedurii în Programul anual al achizițiilor publice 2026 nr. 9568/02.10.2026", "1"),
    ]
    table(doc, ["Nr.", "Denumire document", "Nr. file", "Pagina (de la – la)"],
          [[a, b, c, ""] for a, b, c in rows], widths=[1.2, 11.3, 1.8, 3.2], size=9.5)
    P(doc, "")
    P(doc, "Total: ____ file, numerotate de la 1 la ____.", size=10)
    P(doc, "Notă: numărul de file este completat acolo unde e cunoscut din exportul SEAP / documentul final; "
           "celelalte se completează după scanarea finală.", size=8.5)
    P(doc, "Reprezentant legal,", before=10, after=0)
    P(doc, B(PRIMAR), " – Primar", after=0)
    P(doc, "Semnătura și ștampila ______________")
    return save(doc, "00a. Opis dosar Etapa I")


# ------------------------------------------------------------------ 09 Caiet de sarcini

def caiet():
    doc = new_doc()
    antet(doc)
    P(doc, B("CAIET DE SARCINI"), align="c", size=13, after=8)

    def h(t):
        P(doc, B(t), before=6)

    def li(*t):
        P(doc, "– ", *t, align="j", after=1)

    h("1. INTRODUCERE")
    P(doc, "Comuna Gălbinași, în calitate de autoritate contractantă, intenționează să achiziționeze o nacelă în cadrul proiectului „",
      OBIECT, "”, finanțat prin contractul de finanțare nr. ", CONTRACT_FIN, " (PS PAC 2023-2027, DR-36 LEADER – GAL Ecoul Câmpiei Buzăului), "
      "destinată dotării Serviciului de iluminat public al Comunei Gălbinași.", align="j")
    P(doc, "Cod CPV: ", B(CPV), ".", align="j")
    h("2. GENERALITĂȚI")
    P(doc, "Autoritatea contractantă: Comuna Gălbinași, CIF 3724440.", align="j")
    P(doc, "Furnizorul va livra 1 (una) bucată nacelă montată pe autoșasiu, nouă și nefolosită, cu caracteristicile tehnice minime "
      "precizate la punctul 3. Cerințele din prezentul caiet de sarcini sunt minimale; sunt acceptate produse cu performanțe egale sau "
      "superioare. Orice referire la o anumită marcă, producător sau model se citește cu mențiunea „sau echivalent”.", align="j")
    h("3. OBIECTUL CONTRACTULUI – CARACTERISTICI TEHNICE MINIME")
    P(doc, "1 buc. nacelă pentru dotarea Serviciului de iluminat public al Comunei Gălbinași, cu următoarele caracteristici minime "
      "(corelate cu cererea de finanțare):", align="j")
    table(doc, ["Nr.", "Caracteristică", "Cerință minimă"], [
        ["1", "Tip echipament", "Nacelă montată pe autoșasiu (autoșasiu cu motor diesel)"],
        ["2", "Înălțime de lucru", "între 10 m și 15 m"],
        ["3", "Înălțime maximă a platformei", "între 8 m și 14 m"],
        ["4", "Extindere laterală maximă de lucru", "minimum 6,5 m cu sarcina de 200 kg în coș"],
        ["5", "Capacitate de încărcare a coșului", "între 150 kg și 200 kg (minimum 150 kg)"],
        ["6", "Rotație coș", "minimum 60° stânga + 60° dreapta"],
        ["7", "Rotație turelă", "360° (rotație continuă sau completă)"],
        ["8", "Coș nacelă", "din aluminiu sau din alt material cu greutate redusă și rezistent la coroziune, cu performanțe echivalente"],
        ["9", "Dotări de siguranță", "stabilizatori, sistem de coborâre de urgență, limitator de sarcină, comenzi din coș și de la sol"],
    ], widths=[1, 6, 10], size=10)
    P(doc, "Ofertantul va prezenta în propunerea tehnică fișa tehnică/pliantul producătorului din care să rezulte îndeplinirea fiecărei "
      "cerințe, cu precizarea valorilor efective ofertate.", align="j", before=4)
    h("4. CONDIȚII DE GARANȚIE ȘI SERVICE")
    li("termen de garanție: minimum 24 de luni de la data semnării procesului-verbal de recepție fără obiecțiuni;")
    li("în perioada de garanție, furnizorul asigură gratuit repararea/înlocuirea oricărui echipament sau componente defecte din motive "
       "constructive sau de fabricație, inclusiv manopera și piesele de schimb;")
    li("timp de intervenție în garanție: maximum 72 de ore de la notificarea scrisă a autorității contractante; remedierea în cel mult 15 zile;")
    li("service-ul în perioada de garanție se asigură de furnizor direct sau prin unități de service autorizate; ofertantul va descrie în "
       "propunerea tehnică modul de asigurare a service-ului în garanție și post-garanție;")
    li("în perioada post-garanție, service-ul se asigură contra cost, iar furnizorul își asumă angajamentul de asigurare a pieselor de schimb.")
    h("5. ALTE CONDIȚII SPECIFICE")
    li("manualul de utilizare și cel de întreținere în limba română;")
    li("instruirea gratuită a personalului de deservire, la livrare;")
    li("la livrare se predau: certificatul de conformitate/declarația de conformitate CE, cartea tehnică, certificatul de garanție, "
       "documentele necesare înmatriculării (CIV/omologare RAR) și, după caz, cele privind verificarea ISCIR a instalației de ridicat; "
       "costurile aferente acestor documente sunt incluse în prețul ofertat.")
    h("6. LIVRARE, RECEPȚIE ȘI PLATĂ")
    li("livrarea se face la sediul autorității contractante, cu transportul inclus în preț, în termen de maximum 30 de zile calendaristice de la "
       "data emiterii comenzii ferme;")
    li("recepția cantitativă și calitativă se efectuează la sediul autorității contractante, în termen de maximum 5 zile lucrătoare de la livrare, "
       "de către comisia de recepție numită de autoritatea contractantă, în prezența reprezentantului furnizorului, și se finalizează prin "
       "proces-verbal de recepție, după punerea în funcțiune și instruirea personalului;")
    li("plata se face prin ordin de plată, în termen de 30 de zile de la data primirii facturii, emise după semnarea procesului-verbal de "
       "recepție fără obiecțiuni; nu se acordă avans.")
    h("7. GRAFIC DE TIMP")
    P(doc, "Furnizorul va livra produsul în termen de cel mult 30 de zile calendaristice de la comanda fermă. Ofertele care prevăd un termen de "
      "livrare mai mare vor fi considerate neconforme.", align="j")
    h("8. STANDARDE APLICABILE")
    P(doc, "Produsul va respecta reglementările tehnice și standardele în vigoare (Directiva 2006/42/CE privind echipamentele tehnice, "
      "SR EN 280 pentru platforme de lucru mobile ridicătoare sau echivalent), în ordinea de precedență prevăzută la art. 156 din Legea nr. 98/2016.",
      align="j")
    P(doc, "")
    P(doc, "Întocmit,", after=0)
    P(doc, "Consilier achiziții publice", after=0)
    P(doc, "BUZĂIANU-BARBU PETRONIA")
    return save(doc, "09. Caiet de sarcini - CORECTAT")


# ------------------------------------------------------------------ 08a Clauze contract

def clauze_contract():
    doc = new_doc()
    P(doc, B("COMPLETĂRI ȘI CORECTURI LA MODELUL CONTRACTULUI DE FURNIZARE"), align="c", size=12, after=0)
    P(doc, "(document de lucru – textele se introduc în „Contract furnizare produse”, apoi se reîncarcă în SEAP)", align="c", size=9, after=8)
    rows = [
        ["Art. 4.3 – ajustarea prețului",
         "Se elimină posibilitatea de ajustare. Text propus: „Prețul contractului este ferm și nu se ajustează pe durata derulării contractului.” "
         "(conform strategiei de contractare pct. D3 și fișei de date II.3)."],
        ["Art. 9 – garanția de bună execuție",
         "Se completează: „Garanția de bună execuție este de 10% din prețul contractului fără TVA, respectiv 44.342,16 lei, și se constituie în "
         "termen de 5 zile lucrătoare de la semnarea contractului, prin una din formele prevăzute la art. 40 din H.G. nr. 395/2016. Garanția se "
         "restituie în termen de 14 zile de la data semnării procesului-verbal de recepție fără obiecțiuni.”"],
        ["Articol nou – Garanția produsului",
         "„Perioada de garanție acordată produsului este de ____ luni (minimum 24 de luni, conform ofertei) de la data semnării procesului-verbal "
         "de recepție. În perioada de garanție, furnizorul are obligația de a remedia pe cheltuiala sa, în maximum 72 de ore de la notificare "
         "intervenția și în cel mult 15 zile remedierea, orice defect apărut din motive constructive sau de fabricație. Perioada de garanție se "
         "prelungește cu durata imobilizării produsului.”"],
        ["Art. 18.6 – recepția",
         "„Recepția cantitativă și calitativă se efectuează la sediul achizitorului, în termen de maximum 5 zile lucrătoare de la livrare, de comisia "
         "de recepție a achizitorului, în prezența reprezentantului furnizorului, și se finalizează prin proces-verbal de recepție, după punerea "
         "în funcțiune și instruirea personalului.”"],
        ["Art. 18.7 / 27.3 – plata",
         "Se înlocuiește „în 5 zile de la data primirii sumelor de la finanțator” cu: „Achizitorul are obligația de a efectua plata în termen de 30 de "
         "zile de la data primirii facturii, emise după semnarea procesului-verbal de recepție fără obiecțiuni.” (corelat cu fișa de date IV.4.2 și "
         "caietul de sarcini; plata nu poate fi condiționată de rambursarea de la AFIR – Legea nr. 72/2013)."],
        ["Preambul / date achizitor",
         "Se unifică datele de contact cu antetul: tel. 0238/780.015, fax 0238/780.022 (în draft apare 0238/580.022)."],
    ]
    table(doc, ["Articol", "Text de introdus / modificare"], rows, widths=[3.8, 13.2], size=9.5)
    return save(doc, "08a. Clauze de completat in modelul de contract")


# ------------------------------------------------------------------ CV-uri comisie

def cv(nr, nume, functie, rol):
    doc = new_doc()
    P(doc, B("CURRICULUM VITAE"), align="c", size=13, after=8)
    rows = [
        ["Nume și prenume", nume],
        ["Funcția / compartimentul", functie],
        ["Calitatea în procedură", f"{rol[0].upper()}{rol[1:]} – Dispoziția nr. 140/02.10.2026, procedura simplificată „{OBIECT}”"],
        ["Adresa", "Primăria Comunei Gălbinași, Str. Unirii nr. 71, Comuna Gălbinași, Jud. Buzău"],
        ["Telefon / e-mail", "0238 / 780.015 / ______________________"],
        ["Naționalitate", "română"],
    ]
    t = doc.add_table(rows=0, cols=2)
    t.style = "Table Grid"
    for k, v in rows:
        c = t.add_row().cells
        c[0].text, c[1].text = k, v
        g.shade(c[0], "F2F2F2")
    P(doc, B("Experiență profesională"), before=8)
    table(doc, ["Perioada", "Angajator", "Funcția / principalele activități și responsabilități"],
          [["", "", ""] for _ in range(4)], widths=[3, 5, 9])
    P(doc, B("Educație și formare"), before=8)
    table(doc, ["Perioada", "Instituția", "Calificarea / diploma obținută"], [["", "", ""] for _ in range(3)], widths=[3, 5, 9])
    P(doc, B("Cursuri de specializare (de ex. achiziții publice, expert achiziții)"), before=8)
    table(doc, ["Perioada", "Organizator", "Certificat / atestat"], [["", "", ""] for _ in range(2)], widths=[3, 5, 9])
    P(doc, B("Competențe și abilități"), before=8)
    P(doc, "Limbi străine: ____________________   Competențe de utilizare a calculatorului: ____________________")
    P(doc, "Experiență în evaluarea ofertelor / achiziții publice: ______________________________________________")
    P(doc, "")
    P(doc, "Data: ____ / ____ / 2026", after=0)
    P(doc, "Semnătura: ______________________", after=0)
    P(doc, B(nume))
    return save(doc, f"{nr}. CV - {nume.title()}")


# ------------------------------------------------------------------ Note de corecturi

def note():
    doc = new_doc()
    P(doc, B("GĂLBINAȘI – NACELĂ – ETAPA I: CORECTURI ÎNAINTE DE TRANSMITEREA LA OJFIR"), size=12, after=0)
    P(doc, "Document intern de lucru (nu se include în dosar). Verificare făcută pe baza listei de verificare OJFIR Etapa I "
           "(Anexa 2) și a observațiilor primite la dosarul Scutelnici (29.09.2026).", size=9, after=8)
    secs = [
        ("A. Risc mare – de corectat obligatoriu", [
            "Caiet de sarcini: CPV greșit (43310000-9). Corect: 42418000-9, ca în strategie, notă, analiză, anunț și fișa de date. → înlocuit cu „09. Caiet de sarcini - CORECTAT”.",
            "Caiet de sarcini: valorile fixe „extindere laterală 6,5 m (200 kg) / 8 m (80 kg)” și „rotație coș 60°+60°” reproduc modelul UTILBEN Multitel 145 ALU → reformulate ca minime („minimum 6,5 m cu 200 kg”, „minimum 60°+60°”), care rămân corelate cu cererea de finanțare (punctul 4.2 din lista OJFIR).",
            "Caiet de sarcini: cerința „Furnizorul este importator (dealer) autorizat al producătorului în România” este restrictivă (punctele 5.3 și 5.6 din lista OJFIR) → eliminată și înlocuită cu descrierea modului de asigurare a service-ului autorizat.",
            "CV-urile semnate și datate ale membrilor comisiei (Cristea, Silivestru, Buzăianu-Barbu) și ale rezervelor (Bratu, Grigore) lipsesc. Este exact observația primită la Scutelnici (punctul 6.2) → modelele 11–15 se completează, se semnează și se datează.",
            "Contract: nu există clauză de garanție a produsului; garanția de bună execuție nu are suma completată (10% = 44.342,16 lei); recepția nu are termen; plata este condiționată de primirea banilor de la finanțator → vezi „08a. Clauze de completat in modelul de contract”.",
        ]),
        ("B. Corecturi în SEAP (anunț, fișa de date, DUAE)", [
            "Lista persoanelor cu funcții de decizie conține doar Dragomir, Cristea, Guteniuc și Botezatu → se adaugă Buzăianu-Barbu Petronia, Silivestru Alina-Magdalena, Bratu Ionela și Grigore Tanta (ca în declarația 9567 și dispoziția 140).",
            "Fișa de date IV.4.2 cere propunerea financiară în „Formularul 7” (care este declarația de conflict de interese) → corect: Formularul 8 – Formular de ofertă.",
            "Fișa de date IV.4.1 cere o declarație de însușire a clauzelor contractuale, dar formularul nu există → se adaugă formularul sau se elimină cerința.",
            "Anunțul cere „prestări de servicii similare”, iar fișa de date și DUAE cer „furnizări de produse similare” → se unifică la „furnizări de produse similare”.",
            "Valabilitatea garanției de participare (60 de zile) este mai scurtă decât valabilitatea ofertei (2 luni, până la 26.12.2026) → se stabilește „cel puțin egală cu perioada de valabilitate a ofertei” (minimum 61 de zile).",
            "Termenul de plată din fișa de date (30 de zile de la factură) trebuie să fie identic cu cel din caiet și din contract (corectate).",
            "Datele de contact sunt neunitare (tel. 0238/580.015 în anunț, 0238/780.015 în fișa de date; strada lipsește în anunț) → se aliniază cu antetul.",
            "Numărul contractului de finanțare se scrie cu prefixul „C” (C 36020808218621003751).",
        ]),
        ("C. Strategia de contractare nr. 9564 și nota nr. 9565", [
            "Strategia, pct. D2, conține text preluat dintr-un contract de lucrări („lucrările puse în operă”, „manopere”) și menționează o poliță de asigurare care nu există în contract → se reformulează pentru furnizare de produse.",
            "Strategia, pct. B1, invocă surse de preț nedocumentate (expoziții, catalogul SEAP, anunțuri de atribuire) → se păstrează doar sursele atașate (cele 3 oferte și bugetul din cererea de finanțare) sau se atașează dovezile.",
            "Strategia, pct. B5, enumeră art. 166 ca motiv de excludere → corect: art. 164, 165, 167 și 58–63 din Legea nr. 98/2016. Pct. B3: temeiul procedurii simplificate este art. 7 alin. (2) din Legea nr. 98/2016.",
            "Nota nr. 9565 citează eronat Codul administrativ („art. 129 alin. 29 lit) d”) → se verifică trimiterea.",
            "Dacă se corectează strategia sau nota, documentele se re-semnează (se poate păstra numărul de înregistrare cu mențiunea „revizia 1”).",
        ]),
        ("D. Analiza rezonabilității și anexele", [
            "Scanul analizei pare să nu conțină prima pagină (antet, număr de înregistrare, APROB, secțiunile I–III) → se rescanează complet și se completează numărul de înregistrare.",
            "Anexele analizei (cele 3 oferte și lista caracteristicilor tehnice) există doar în folderul COMPLETATE → se includ în dosar după analiză (documentul 3).",
            "Analiza marchează „C” (conform) la KSM (rotație 90°+90°) și Italia Star (extindere 9 m), ceea ce e coerent doar cu cerințele formulate ca minime → coerent cu caietul corectat.",
            "Ofertele sunt din decembrie 2025 și au valabilitatea expirată. OJFIR poate cere actualizarea lor; valoarea estimată rămâne cea din cererea de finanțare / PAP (443.421,60 lei).",
        ]),
        ("E. De decis", [
            "Termenul de livrare de 30 de zile este strâns (KSM indică 6–8 săptămâni). Dacă se mărește (de ex. la 60 de zile), se modifică unitar în caiet, fișa de date, anunț și contract.",
            "Titlul obiectului: „ACHIZIȚIE NACELĂ…” în toate documentele, dar „ACHIZIȚIA UNEI NACELE…” în declarația 9568 → se verifică titlul exact din contractul de finanțare și se unifică.",
            "Fișa navetă („0. Fisa naveta - Galbinasi.docx”) este corectă și conformă cu modelul Scutelnici; se completează numărul și data de înregistrare și se semnează.",
        ]),
    ]
    for titlu, items in secs:
        P(doc, B(titlu), before=8)
        for it in items:
            P(doc, "☐ ", it, align="j", after=2, size=10)
    return save(doc, "_CORECTURI inainte de transmitere - Galbinasi Etapa I")


if __name__ == "__main__":
    files = [opis(), caiet(), clauze_contract()]
    for i, (n, f, r) in enumerate(COMISIE, 11):
        files.append(cv(i, n, f, r))
    files.append(note())
    for f in files:
        print(f.name, f.stat().st_size)
