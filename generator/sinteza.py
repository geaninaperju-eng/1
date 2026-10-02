"""Sinteza dosarelor generate: stadiu, rezonabilitate și ce trebuie verificat (document intern)."""
import json
from pathlib import Path

import docx
from docx.enum.section import WD_ORIENT
from docx.shared import Cm

import genereaza_dosare as g

ROOT = Path(__file__).resolve().parent
D = json.loads((ROOT / "data" / "proiecte.json").read_text(encoding="utf-8"))

def build(part, D):
  doc = g.new_doc()
  sec = doc.sections[0]
  sec.orientation = WD_ORIENT.LANDSCAPE
  sec.page_width, sec.page_height = Cm(29.7), Cm(21.0)
  g.P(doc, g.B("SINTEZĂ – DOSARE DE ACHIZIȚII PENTRU AVIZARE AFIR (proiecte GAL)"), size=13, after=0)
  g.P(doc, f"Partea {part} – {len(D)} dosare. Document intern de lucru – nu se depune.", size=10, after=6)
  rows = []
  for d in D:
    a = d["achizitie"]
    t = " ".join(p.text for p in docx.Document(next((ROOT / "output" / g.slug(d["folder"]) / g.slug(d["dosar"])).glob("06*"))).paragraphs)
    rez = "DA" if "ofertat: DA" in t else ("NU – de justificat" if "ofertat: NU" in t else "fără oferte")
    stadiu = (f"{a.get('cod_seap') or ''} {('contract ' + a['contract_nr_data']) if a.get('contract_nr_data') else ''}").strip() or "neinițiată"
    val = g.lei(a.get("valoare_contract_fara_tva") or a.get("valoare_estimata_fara_tva")) or "–"
    obs = d.get("observatii") or []
    obs = [obs] if isinstance(obs, str) else obs
    rows.append([d["folder"], d["dosar"].replace("Dosar achizitie ", ""), a["tip"], "simplificată" if a.get("procedura") == "simplificata" else "directă",
                 val, stadiu, rez, "\n".join("• " + o for o in obs)])
  g.table(doc, ["UAT", "Achiziție", "Tip", "Procedură", "Valoare fără TVA (lei)", "SEAP / contract", "Rezonabil", "De verificat / completat"],
        rows, widths=[2.6, 3.2, 1.6, 1.8, 2.2, 3.2, 1.8, 10.5], size=7.5)
  p = ROOT / "output" / f"00 SINTEZA dosare achizitii - partea {part}.docx"
  doc.save(p)
  g.compacteaza(p)
  print(p)


for old in (ROOT / "output").glob("00 SINTEZA*.docx"):
    old.unlink()
N = 8
for i in range(0, len(D), N):
    build(i // N + 1, D[i:i + N])
