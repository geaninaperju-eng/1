import re,html,sys,json
def items(path):
    t=open(path,encoding='utf-8',errors='ignore').read()
    parts=re.split(r'(?=<a[^>]*href="vizualizare_anunt\.jsp\?ch=)',t)
    out=[]
    for p in parts[1:]:
        m=re.match(r'<a[^>]*href="(vizualizare_anunt\.jsp\?ch=[^"&]+)[^"]*"',p)
        txt=html.unescape(re.sub(r'<[^>]+>','\n',p[:4000])); txt=re.sub(r'\s*\n\s*','\n',txt).strip()
        tl=re.search(r'title="([^"]+)"',p[:1500])
        g=lambda k: (re.search(k+r'\s*:?\s*\n?([^\n]+)',txt) or [None,None])[1]
        out.append(dict(link=m.group(1),title=html.unescape(tl.group(1)) if tl else txt.split('\n')[0],
            termen=g('Termen limita:'),loc=g('Localizare:'),pub=g('Data publicare:'),val=g('Val. estimata:')))
    return out
if __name__=='__main__':
    for f in sys.argv[1:]:
        for i in items(f): print(json.dumps(i,ensure_ascii=False))
