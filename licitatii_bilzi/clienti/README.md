# Bază de date – clienți potențiali BILZI STEEL PROFILE (acoperișuri / învelitori / șarpante)

Fișier: `clienti_potentiali.xlsx` (copie Google Sheet în Drive: folderul „BILZI STEEL PROFILE”).

## Foi
| Foaie | Conținut |
|---|---|
| **Clienti** | un rând per autoritate/entitate: scor prioritate 1–5, denumire, CUI, tip, județ, localitate, adresă, telefon, e-mail, web, nr. achiziții de acoperiș găsite, valoare totală, ultima achiziție (dată, obiect, valoare, modul), planuri cu acoperiș în județ, distanța aprox. față de Buzău |
| **Achizitii** | un rând per anunț / cumpărare directă / atribuire: data, modul, nr. anunț, tip procedură, obiect, organizator + CUI, județ, valoare, termen, câștigător/ofertant + CUI, link |
| **Concurenti** | firmele câștigătoare / ofertante la cumpărări directe și atribuiri de acoperiș în zonă: denumire, CUI, sediu, contact, nr. contracte, valoare totală, județe, clienți |
| **Planuri_judete** | pozițiile din planurile anuale de achiziții cu acoperiș/învelitori, pe județe (vezi limite) |

## Cum s-a construit
1. **Sursă:** www.licitatiipublice.ro, cont autentificat (variabilele de mediu `LICITATIIPUBLICE_USER` / `LICITATIIPUBLICE_PASS`; parola și cookie-urile nu sunt salvate în repo).
2. **Căutare** (`harvest.py`): modulele `planuriachizitii`, `achizitii`, `cumpararidirecte`, `atribuiri` ×
   cuvintele cheie *acoperis, invelitoare, sarpanta, tigla, jgheaburi, tinichigerie, hidroizolatie acoperis, tabla, materiale acoperis*,
   cu filtrele: „Afișează expirate” (`fltex`), data publicării = ultimele 24 de luni (`dp`), județele
   Buzău, Brăila, Galați, Vrancea, Prahova, Ialomița, Călărași, București, Ilfov, Dâmbovița, Brașov, Covasna, Bacău, Vaslui,
   Tulcea, Constanța, Giurgiu, Argeș, Teleorman, Neamț, Iași, Harghita (`lc`). Pauză ~1,2 s între cereri.
   Notă tehnică: filtrele trebuie trimise *după* criteriul de căutare (setarea criteriului resetează filtrele).
3. **Relevanță:** căutarea site-ului este full-text (prinde și investiții mari care doar menționează acoperișul), deci s-au păstrat
   doar rândurile al căror **titlu** conține: acoperiș, învelitoare, șarpantă, țiglă, jgheab, burlan, tinichigerie, streașină,
   hidroizolație, astereală, terasă, tablă cutată/zincată/tip țiglă, panouri sandwich.
4. **Detalii** (`detalii.py`): pagina fiecărui anunț relevant → organizator complet (CUI, adresă, telefon, e-mail, web),
   ofertant/câștigător, valoare, nr. anunț, tip procedură.
5. **Agregare** (`build_xlsx.py`): clienții sunt grupați după CUI (sau denumire, unde lipsește CUI).
   **Scor prioritate** (1–5) din puncte: distanța (≤60 km: 2; ≤130: 1,5; ≤200: 1; altfel 0,5) + frecvența (≥4 achiziții: 1,5; 2–3: 1; 1: 0,5)
   + valori potrivite pentru BILZI (cel puțin o achiziție de 20–150 mii lei: 1; 5–20 mii sau 150–500 mii: 0,5)
   + activitate în ultimele 12 luni (0,5) + există poziții de plan cu acoperiș în județ (0,25).
   Prag: ≥4,5 → 5; ≥3,5 → 4; ≥2,75 → 3; ≥2 → 2; restul 1.
   Distanța este rutieră, aproximativă, până la reședința de județ (pentru câteva localități din jurul Buzăului e mai precisă).

Rulare: `python3 harvest.py && python3 detalii.py && python3 build_xlsx.py` (cache în `date/`, ignorat de git; relansările continuă de unde au rămas).

## Perioada acoperită
Anunțuri publicate în ultimele ~24 de luni (octombrie 2024 – 1 octombrie 2026), inclusiv cele expirate; la atribuiri apar și câteva date de atribuire mai vechi.

## Limite
- **Planurile anuale de achiziții** cer o „extra-opțiune” plătită pe licitatiipublice.ro: titlul este trunchiat la ~20 de caractere și organizatorul e mascat.
  Ca urmare coloana „Are plan achiziții cu acoperiș” nu poate fi stabilită per entitate: foaia Planuri_judete arată pozițiile pe județ,
  iar în foaia Clienti apare numărul de poziții din județul clientului. Pentru detalii trebuie activată extra-opțiunea sau verificat direct în SEAP.
- **Detalii incomplete:** contul a început să refuze paginile de detaliu după ~1.550 de pagini descărcate (limită zilnică), așa că descărcarea s-a oprit.
  Pentru ~180 de rânduri (cele 88 de **atribuiri** și o parte din cumpărările directe) există doar datele din listă (titlu, organizator, județ, dată),
  fără valoare, CUI sau câștigător; aceste rânduri sunt marcate „(detaliu nedescărcat)”. Pentru ele, foaia Concurenti nu include câștigătorii.
  Relansarea `python3 detalii.py` (altă zi, cu pauze mai mari) completează doar ce lipsește.
- În modulul cumpărări directe, căutările „hidroizolatie acoperis”, „tabla” și „materiale acoperis” au fost întrerupte;
  rezultatele lor relevante se suprapun în mare parte cu „acoperis” / „tigla” / „jgheaburi”, dar pot lipsi câteva cumpărări.
- Pentru cuvintele foarte generale („tabla”, „materiale acoperis”) s-au parcurs doar cele mai recente 600, respectiv 200 de rezultate per modul.
- Valorile sunt cele afișate de site (estimată / preț maximal / valoare atribuită); sumele în EUR au fost convertite cu 5 lei/EUR.
  Unele anunțuri (consultări de piață) nu au valoare.
- Tipul entității este dedus din denumire, prin reguli simple; poate fi greșit la denumiri atipice.
- Filtrul pe județ al site-ului se aplică pe „Localizare”; organizatorii centrali (ministere, regii) pot avea sediul în București și lucrări în alte județe.
