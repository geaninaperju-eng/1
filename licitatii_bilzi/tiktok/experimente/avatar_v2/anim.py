"""Motorul de animație al mascotei BILZI.

Primește o cronologie (replici cu voce, gesturi, poziții) și produce, pentru fiecare cadru,
parametrii pe care avatar.html îi aplică desenului. Mișcarea folosește arcuri amortizate
(inerție, depășire ușoară), mișcări de fond (respirație, transfer de greutate), clipit aleator,
încuviințări pe silabe și forme ale gurii după conținutul spectral al vocii.
"""
import array, math, random, subprocess

FPS = 30

# Poze ale brațului drept: (umăr, cot, încheietură, mână). 0° = brațul atârnă în jos;
# unghiurile negative ridică brațul spre exterior (dreapta ecranului).
POSES = {
    "idle":    (-8, -8, 4, "relax"),
    "wave":    (-128, -42, 0, "open"),
    "thumb":   (-52, -108, 8, "thumb"),
    "point":   (-130, -14, -6, "point"),
    "explain": (-34, -66, -24, "open"),
}


class Spring:
    """Arc de ordinul 2: urmărește ținta cu inerție (freq = viteză, zeta = amortizare)."""
    def __init__(self, x=0.0, freq=3.0, zeta=0.6):
        self.x, self.v, self.f, self.z = x, 0.0, freq, zeta

    def step(self, target, dt=1 / FPS):
        w = 2 * math.pi * self.f
        for _ in range(4):                      # sub-pași pentru stabilitate
            h = dt / 4
            a = w * w * (target - self.x) - 2 * self.z * w * self.v
            self.v += a * h
            self.x += self.v * h
        return self.x

    def kick(self, dv):
        self.v += dv


def noise(t, seed):
    """Zgomot lin (sumă de sinusuri cu faze aleatoare) în [-1, 1]."""
    r = random.Random(seed)
    return sum(math.sin(t * f + r.uniform(0, 6.28)) * a
               for f, a in ((0.71, 0.5), (1.33, 0.3), (2.47, 0.2)))


def voice_features(wav, total):
    """Pentru fiecare cadru: energie (deschiderea gurii) și raport de frecvențe înalte (gură lată)."""
    def band(flt):
        cmd = ["ffmpeg", "-v", "error", "-i", wav]
        if flt:
            cmd += ["-af", flt]
        cmd += ["-ac", "1", "-ar", "16000", "-f", "s16le", "-"]
        a = array.array("h", subprocess.check_output(cmd))
        hop = 16000 // FPS
        out = []
        for i in range(int(total * FPS) + 2):
            seg = a[i * hop:(i + 1) * hop]
            out.append(math.sqrt(sum(x * x for x in seg) / len(seg)) if len(seg) else 0.0)
        return out
    full, low, high = band(None), band("lowpass=f=900"), band("highpass=f=2500")
    peak = sorted(full)[int(len(full) * 0.97)] or 1.0
    opn, wide = [], []
    prev = 0.0
    for e, lo, hi in zip(full, low, high):
        v = min(1.0, (e / peak) ** 0.75 * 1.1)
        v = v if v > 0.1 else 0.0
        prev = v if v > prev else prev * 0.5 + v * 0.5
        opn.append(prev)
        wide.append(hi / (lo + hi + 1e-6))
    # normalizăm raportul pe segmentele cu voce
    voiced = sorted(w for w, o in zip(wide, opn) if o > 0.15) or [0.3]
    lo_r, hi_r = voiced[int(len(voiced) * 0.1)], voiced[int(len(voiced) * 0.9)]
    wide = [min(1.0, max(0.0, (w - lo_r) / (hi_r - lo_r + 1e-6))) for w in wide]
    return opn, wide


def onsets(opn):
    """Cadrele unde începe o silabă accentuată (pentru încuviințări și sprâncene)."""
    res, last = [], -99
    for i in range(1, len(opn)):
        if opn[i] > 0.45 and opn[i] - opn[i - 1] > 0.18 and i - last > 5:
            res.append((i, opn[i])); last = i
    return res


def animate(total, lines, place, opn, wide, seed=7, enter_at=0.1, variant=None, bubble_fn=None):
    """
    total  – durata (s)
    lines  – listă de dict(start, dur, text, gestures=[(t_rel, gest), ...], question=bool, final=bool)
    place  – funcție t -> (x, y, scale) poziția ancorei (bază, centru) în stratul randat
    opn, wide – trăsăturile vocii pe cadru (din voice_features)
    bubble_fn – funcție (line, t, show) -> dict pentru balon sau None
    Întoarce lista de cadre (dict) pentru avatar.html.
    """
    rnd = random.Random(seed)
    n = int(round(total * FPS))
    ons = dict(onsets(opn))

    # clipit aleator: intervale 2–5 s, uneori dublu
    blinks, t = [], 0.6 + rnd.uniform(0, 1.5)
    while t < total:
        blinks.append(t)
        if rnd.random() < 0.2:
            blinks.append(t + 0.32)
        t += rnd.uniform(2.0, 5.0)

    # sacade ale privirii
    sacc, t = [], 0.0
    while t < total:
        sacc.append((t, rnd.uniform(-0.35, 0.35), rnd.uniform(-0.25, 0.2)))
        t += rnd.uniform(0.7, 2.2)

    S = dict(
        enter=Spring(1.0, 2.1, 0.42), a1=Spring(POSES["idle"][0], 2.6, 0.55), a2=Spring(POSES["idle"][1], 3.2, 0.5),
        a3=Spring(0, 4.0, 0.45), nod=Spring(0, 4.5, 0.35), hr=Spring(0, 2.2, 0.6), turn=Spring(0, 2.5, 0.7),
        lx=Spring(0, 9, 0.9), ly=Spring(0, 9, 0.9), raise_=Spring(0, 4, 0.5), smile=Spring(0.35, 3, 0.8),
        mo=Spring(0, 14, 0.75), mw=Spring(1, 10, 0.8), rot=Spring(0, 1.8, 0.5), tail=Spring(0, 2.5, 0.25),
        beat=Spring(0, 5, 0.4),
    )
    frames = []
    hand = "relax"
    prev_x = None
    for fi in range(n):
        t = fi / FPS
        # --- replica activă și gestul curent
        cur = None
        for ln in lines:
            if ln["start"] - 0.3 <= t <= ln["start"] + ln["dur"] + (0.8 if not ln.get("final") else 99):
                cur = ln
        talking = cur is not None and cur["start"] <= t <= cur["start"] + cur["dur"]
        gest = "idle"
        if cur:
            rel = t - cur["start"]
            for tr, g in cur.get("gestures", []):
                if rel >= tr - 0.25:
                    gest = g
            if rel > cur["dur"] + 0.45 and not cur.get("final"):
                gest = "idle"
        p = POSES[gest]
        tgt1, tgt2, tgt3 = p[0], p[1], p[2]
        if gest == "wave":
            ph = 2 * math.pi * 1.7 * t
            tgt2 += 20 * math.sin(ph); tgt3 += 16 * math.sin(ph - 1.0); tgt1 += 4 * math.sin(ph - 0.4)
        if gest == "explain":
            tgt2 += S["beat"].x * 14; tgt3 += S["beat"].x * 10
        if gest == "idle":
            tgt1 += 2 * math.sin(2 * math.pi * t / 3.6)
        # mâna își schimbă forma la jumătatea drumului
        if abs(S["a1"].x - tgt1) < 35 or gest == "idle":
            hand = p[3]
        a1, a2, a3 = S["a1"].step(tgt1), S["a2"].step(tgt2), S["a3"].step(tgt3)

        # --- voce: deschidere, lățime, încuviințări
        o_t = opn[fi] if talking else 0.0
        w_t = 0.78 + 0.42 * wide[fi] if talking else 1.0
        if talking and o_t > 0.5 and wide[fi] < 0.25:
            w_t = 0.72                                    # sunete rotunde (o/u)
        mo, mw = S["mo"].step(o_t), S["mw"].step(w_t)
        if talking and fi in ons:
            k = ons[fi]
            S["nod"].kick(40 * k)
            S["beat"].kick(9 * k)
            if k > 0.75:
                S["raise_"].kick(5 * k)
        S["beat"].step(0)
        nod = S["nod"].step(0)

        # --- cap, privire, sprâncene
        look_x, look_y, turn_t = 0.0, 0.0, 0.18 * noise(t * 0.5, seed + 1)
        for st, sx, sy in sacc:
            if t >= st:
                look_x, look_y = sx, sy
        if talking:
            look_x *= 0.35; look_y *= 0.35          # când vorbește, privește spre cameră
        if gest == "point":
            look_x, look_y, turn_t = 0.85, -0.7, 0.55
        if gest == "wave":
            turn_t = -0.1
        lx, ly = S["lx"].step(look_x), S["ly"].step(look_y)
        turn = S["turn"].step(turn_t)
        hr = S["hr"].step(2.2 * noise(t * 0.6, seed + 2) + (3 if cur and cur.get("question") and talking else 0)
                          + (2.5 * math.sin(t * 6) * mo if talking else 0))
        raise_t = 0.25 if talking else 0.0
        if cur and cur.get("question") and talking:
            raise_t = 0.7
        if gest == "wave" and t - (cur["start"] if cur else 0) < 0.8:
            raise_t = 0.8
        braise = S["raise_"].step(raise_t)
        tilt = 0.5 if (cur and cur.get("question") and talking) else 0.0

        # --- zâmbet: mai mare după replică (poanta) și în repaus
        smile_t = 0.35
        if talking:
            smile_t = 0.15
        elif cur and t > cur["start"] + cur["dur"]:
            smile_t = 0.85
        sm = S["smile"].step(smile_t)

        # --- clipit
        close = 0.0
        for bt in blinks:
            d = t - bt
            if 0 <= d < 0.17:
                close = math.sin(math.pi * d / 0.17)
        if abs(S["turn"].v) > 1.2:
            close = max(close, 0.6)

        # --- corp: intrare, respirație, greutate, înclinare
        x, y, sc = place(t)
        if prev_x is not None and abs(x - prev_x) > 1:
            S["rot"].kick((x - prev_x) * -0.25)          # se apleacă în direcția mișcării
        prev_x = x
        ent = S["enter"].step(0.0 if t >= enter_at else 1.0)
        y += ent * 900
        stretch = max(-0.07, min(0.07, -S["enter"].v * 0.025))
        sy, sx = 1 + stretch, 1 - stretch * 0.6
        breath = 0.011 * (0.5 + 0.5 * math.sin(2 * math.pi * t / 3.4)) + (0.004 * mo if talking else 0)
        x += 5 * noise(t * 0.35, seed + 3)
        rot = S["rot"].step(0.9 * noise(t * 0.4, seed + 4) + (1.2 * math.sin(t * 4) * mo if talking else 0))
        lsway = 1.5 * math.sin(2 * math.pi * t / 3.4 + 0.6) + 2 * noise(t * 0.5, seed + 5)
        tail = S["tail"].step(-rot * 3 - hr)

        fr = dict(t=round(t, 3), variant=variant, x=x, y=y, sc=sc, sx=sx, sy=sy, rot=rot, breath=breath,
                  lsway=lsway, tailSwing=tail,
                  head=dict(r=hr, nod=nod, turn=turn, up=-ly * 0.3),
                  eyes=dict(lx=lx, ly=ly, close=close),
                  brow={"raise": braise, "tilt": tilt},
                  mouth=dict(o=mo, w=mw, s=sm),
                  arm=dict(a1=a1, a2=a2, a3=a3, hand=hand),
                  bubble=None)
        if bubble_fn and cur and not cur.get("final"):
            le = cur["start"] + cur["dur"]
            show = min(1.0, (t - cur["start"] + 0.15) / 0.3)
            if t > le + 0.35:
                show = max(0.0, 1 - (t - le - 0.35) / 0.25)
            if show > 0:
                fr["bubble"] = bubble_fn(cur, t, show)
        frames.append(fr)
    return frames
