"""Faza 1: liste de rezultate (toate paginile) pe module x cuvinte cheie, 22 județe, ultimele 24 luni, inclusiv expirate.
Rezultatul: date/lista.json (cache; relansarea continuă de unde a rămas)."""
import json, os, re, time, html, math, sys
import lp_client as L

JUDETE = ("Buzau Braila Galati Vrancea Prahova Ialomita Calarasi Bucuresti Ilfov Dambovita Brasov Covasna "
          "Bacau Vaslui Tulcea Constanta Giurgiu Arges Teleorman Neamt Iasi Harghita").split()
MODULE = ["planuriachizitii", "achizitii", "cumpararidirecte", "atribuiri"]
CUVINTE = ["acoperis", "invelitoare", "sarpanta", "tigla", "jgheaburi", "tinichigerie",
           "hidroizolatie acoperis", "tabla", "materiale acoperis"]
MONTHS = 24
# cuvinte foarte generale (căutarea e full-text): limităm la cele mai recente pagini
MAXPG = {"tabla": 12, "materiale acoperis": 4}
OUT = os.path.join(os.path.dirname(__file__), "date", "lista.json")

def parse_list(h, modul):
    out = []
    for m in re.finditer(r'property="view" aria-valuenow="([0-9a-f]+)"(.*?)(?=property="view" aria-valuenow=|\Z)', h, re.S):
        ch, body = m.group(1), m.group(2)
        tl = re.search(r'class="titlu[^"]*"[^>]*>(.*?)</a>', body, re.S)
        title = html.unescape(re.sub(r"<[^>]+>", "", tl.group(1))).strip() if tl else ""
        auth = re.search(r'class="authority">(.*?)</span>', body, re.S)
        t = L.text(body)
        g = lambda k: (re.search(re.escape(k) + r"\s*\n([^\n]+)", t) or [None, None])[1]
        out.append(dict(modul=modul, ch=ch, titlu=title,
                        organizator=html.unescape(auth.group(1)).strip() if auth else None,
                        localizare=g("Localizare:"), data_publicare=g("Data publicare:"),
                        termen=g("Termen limita:"), data_atribuire=g("Data atribuire:"),
                        data_estimata=g("Data estimata achizitie:"),
                        valoare=g("Val. estimata:") or g("Valoare atribuita:") or g("Pret maximal:")))
    return out

def main():
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    db = json.load(open(OUT)) if os.path.exists(OUT) else {"items": {}, "done": []}
    c = L.LP()
    lc = "+".join(JUDETE)
    for modul in MODULE:
        for kw in CUVINTE:
            key = f"{modul}|{kw}"
            if key in db["done"]:
                continue
            c._get(f"module/{modul}/")
            pg, total = 0, None
            while True:
                now = int(time.time() * 1000)
                c._post("module/cautare/setare_valori_paginare.jsp", dict(pg=pg))
                c._post("module/cautare/setare_criterii_cautare.jsp", dict(c1=kw))
                c._post("module/cautare/setare_valori_filtre.jsp",
                        dict(fltex="true", lc=lc, dp=f"{now - MONTHS * 30 * 86400000},{now}"))
                c._post("module/cautare/setare_valori_ordonare.jsp", dict(orderby=0, mode="desc"))
                h = c._post(f"module/{modul}/lista_ajax.jsp")
                if total is None:
                    m = re.search(r"\(\s*\d+-\d+\s*/\s*(\d+)\s*\)", h)
                    total = int(m.group(1)) if m else 0
                its = parse_list(h, modul)
                for it in its:
                    k = f"{modul}|{it['ch']}"
                    if k in db["items"]:
                        old = db["items"][k]
                        it["cuvinte"] = sorted(set(old["cuvinte"] + [kw]))
                        db["items"][k] = it
                    else:
                        it["cuvinte"] = [kw]
                        db["items"][k] = it
                print(f"{key} pg={pg} +{len(its)} / total {total}", flush=True)
                pg += 1
                if not its or pg >= min(math.ceil(total / 50), MAXPG.get(kw, 60)):
                    break
            db["done"].append(key)
            json.dump(db, open(OUT, "w"), ensure_ascii=False)
    print("gata", len(db["items"]))

if __name__ == "__main__":
    main()
