# Căutare licitații pe licitatiipublice.ro

Necesită acces de rețea la `www.licitatiipublice.ro` (și `api.tender-service.com` pentru documentație).

- `./lp_search.sh "acoperis" achizitii 0` – lista de rezultate (HTML) pentru cuvântul cheie, modulul și pagina date.
  Modulul `achizitii` = anunțuri deschise (achiziții directe ADV, proceduri simplificate, licitații).
  Modulul `cumpararidirecte` = cumpărări din catalog care au DEJA un ofertant ales – nu sunt deschise, se ignoră.
- `python3 parse.py rezultat.html` – un rând JSON per anunț: link, titlu (trunchiat), termen, localizare, data publicării, valoare.
- Pagina publică a anunțului (`module/achizitii/vizualizare_anunt.jsp?ch=...`) arată tipul procedurii, termenul cu ora,
  valoarea și numele fișierelor din documentație. Textul integral, PDF-ul anunțului și descărcarea documentației
  cer autentificare (variabilele de mediu LICITATIIPUBLICE_USER / LICITATIIPUBLICE_PASS).
