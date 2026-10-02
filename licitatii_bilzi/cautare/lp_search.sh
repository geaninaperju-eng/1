#!/bin/bash
# usage: lp_search.sh "c1 words" [modul] [page]
B=https://www.licitatiipublice.ro/licitatiipublice
M=${2:-achizitii}; PG=${3:-0}
J=cj_search.txt; rm -f $J
A="Mozilla/5.0"
curl -s -m 30 -c $J -b $J -A "$A" "$B/module/$M/" -o /dev/null
curl -s -m 30 -c $J -b $J -A "$A" -X POST "$B/module/cautare/setare_valori_filtre.jsp" -d action=reset -o /dev/null
curl -s -m 30 -c $J -b $J -A "$A" -X POST "$B/module/cautare/setare_valori_ordonare.jsp" -d orderby=0 -d mode=desc -o /dev/null
curl -s -m 30 -c $J -b $J -A "$A" -X POST "$B/module/cautare/setare_criterii_cautare.jsp" --data-urlencode "c1=$1" -o /dev/null
curl -s -m 30 -c $J -b $J -A "$A" -X POST "$B/module/cautare/setare_valori_paginare.jsp" -d pg=$PG -o /dev/null
curl -s -m 30 -c $J -b $J -A "$A" -X POST "$B/module/$M/lista_ajax.jsp"
