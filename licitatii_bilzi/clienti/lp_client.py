"""Client minimal licitatiipublice.ro (autentificat, cu filtru expirate + perioadă)."""
import os, time, re, html, requests
B = "https://www.licitatiipublice.ro/licitatiipublice"
PAUSE = 1.2

class LP:
    def __init__(self):
        self.s = requests.Session()
        self.s.headers["User-Agent"] = "Mozilla/5.0"
        self.login()

    def _post(self, path, data=None):
        time.sleep(PAUSE)
        for i in range(4):
            try:
                r = self.s.post(f"{B}/{path}", data=data or {}, timeout=60)
                r.encoding = "utf-8"
                return r.text
            except requests.RequestException:
                time.sleep(2 ** (i + 1))
        raise RuntimeError("cerere eșuată: " + path)

    def _get(self, path):
        time.sleep(PAUSE)
        for i in range(4):
            try:
                r = self.s.get(f"{B}/{path}", timeout=60)
                r.encoding = "utf-8"
                return r.text
            except requests.RequestException:
                time.sleep(2 ** (i + 1))
        raise RuntimeError("cerere eșuată: " + path)

    def login(self):
        self._get("module/login/modal_login.jsp")
        r = self._post("module/login/validare_login.jsp", dict(
            numecont=os.environ["LICITATIIPUBLICE_USER"],
            parola=os.environ["LICITATIIPUBLICE_PASS"], checkedauthalive="true"))
        if '"succes":true' not in r:
            raise RuntimeError("autentificare eșuată")

    def search(self, word, modul, page=0, months=24):
        self._get(f"module/{modul}/")
        self._post("module/cautare/setare_valori_filtre.jsp", dict(action="reset"))
        now = int(time.time() * 1000)
        filt = dict(fltex="true", dp=f"{now - months * 30 * 86400000},{now}")
        self._post("module/cautare/setare_valori_filtre.jsp", filt)
        self._post("module/cautare/setare_valori_ordonare.jsp", dict(orderby=0, mode="desc"))
        self._post("module/cautare/setare_criterii_cautare.jsp", dict(c1=word))
        self._post("module/cautare/setare_valori_paginare.jsp", dict(pg=page))
        return self._post(f"module/{modul}/lista_ajax.jsp")

    def page(self, modul, pg):
        """pagina următoare cu aceleași criterii (după search)."""
        self._post("module/cautare/setare_valori_paginare.jsp", dict(pg=pg))
        return self._post(f"module/{modul}/lista_ajax.jsp")

    def detail(self, modul, ch):
        return self._get(f"module/{modul}/vizualizare_anunt.jsp?ch={ch}")

def text(h):
    t = html.unescape(re.sub(r"<script.*?</script>|<style.*?</style>", " ", h, flags=re.S))
    t = re.sub(r"<[^>]+>", "\n", t)
    return re.sub(r"\s*\n\s*", "\n", t).strip()
