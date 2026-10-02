"""Faza 2: pagini de detaliu (organizator complet, ofertant/câștigător, CPV, descriere) pentru anunțurile relevante.
Cache: date/detalii.json."""
import json, os, re, sys
import lp_client as L

D = os.path.join(os.path.dirname(__file__), "date")
ROOF = re.compile(r"acoperi|invelitoar|invelitori|sarpant|tigl|jgheab|burlan|tinichigeri|streasin|"
                  r"hidroizola|astereal|lindab|tabla (cutata|zincata|vopsita|profilata|tip tigla)|"
                  r"tabla pentru acoperis|panouri sandwich|terasa", re.I)
KEYS = ["Cod fiscal:", "Judet:", "Localitate:", "Adresa:", "Telefon:", "Email:", "Web:"]

def parse_detail(h):
    t = L.text(h).split("\n")
    d = {}
    def after(k):
        try:
            i = t.index(k)
            return t[i + 1] if i + 1 < len(t) and not t[i + 1].endswith(":") else ""
        except ValueError:
            return None
    i0 = t.index("Informatii") if "Informatii" in t else 0
    d["titlu"] = t[i0 - 3] if i0 >= 3 else None
    for k in ["Tip procedura:", "Numar anunt:", "Pret maximal:", "Valoare estimata:", "Val. estimata:",
              "Valoare atribuita:", "Valoare:", "Data atribuire:", "Data limita:", "Termen limita:",
              "Data estimata achizitie:", "Descriere:", "Obiect:", "Tip contract:", "Localizare:", "Data publicare:"]:
        v = after(k)
        if v:
            d[k.rstrip(":")] = v
    d["cpv"] = "; ".join(sorted({x for x in t if re.match(r"^\d{8}(-\d)?\s*\(", x)}))
    ents = []
    for i, x in enumerate(t):
        if x == "Cod fiscal:" and i >= 2:
            e = {"sectiune": t[i - 2], "denumire": t[i - 1]}
            j = i
            while j < len(t) and not t[j].startswith("Vezi") and j < i + 20:
                if t[j] in KEYS:
                    nxt = t[j + 1] if j + 1 < len(t) else ""
                    e[t[j].rstrip(":")] = "" if nxt in KEYS or nxt.startswith("Vezi") else nxt
                j += 1
            ents.append(e)
    d["entitati"] = ents
    return d

def main(limit=None):
    lista = json.load(open(os.path.join(D, "lista.json")))["items"]
    out_p = os.path.join(D, "detalii.json")
    det = json.load(open(out_p)) if os.path.exists(out_p) else {}
    todo = [k for k, it in lista.items() if k not in det and
            (ROOF.search(it["titlu"] or "") or
             # titlurile din planuri sunt trunchiate: le descărcăm dacă au venit din cuvinte cheie specifice
             (it["modul"] == "planuriachizitii" and set(it["cuvinte"]) - {"tabla", "materiale acoperis"}))]
    print("de descărcat:", len(todo), flush=True)
    c = L.LP()
    for n, k in enumerate(todo[:limit]):
        modul, ch = k.split("|")
        try:
            det[k] = parse_detail(c.detail(modul, ch))
        except Exception as ex:
            print("eroare", k, ex, flush=True)
            continue
        if n % 25 == 0:
            print(n, k, det[k].get("titlu"), flush=True)
            json.dump(det, open(out_p, "w"), ensure_ascii=False)
    json.dump(det, open(out_p, "w"), ensure_ascii=False)
    print("gata", len(det))

if __name__ == "__main__":
    main(int(sys.argv[1]) if len(sys.argv) > 1 else None)
