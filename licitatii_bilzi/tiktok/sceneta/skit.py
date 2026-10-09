#!/usr/bin/env python3
"""Scenetă animată: Nea Bilzi pune un acoperiș (1080x1920, ~30 s).

Personajul are schelet (șold-genunchi-gleznă, umăr-cot-încheietură), iar mișcarea e compusă din
acțiuni (mers, urcat pe scară, săritură, bătut țigle, alunecare, atârnat de jgheab, cățărare,
șters pe frunte, deget mare) legate între ele prin tranziții line. Peste ele se adaugă respirația,
clipitul, privirea, încuviințările pe silabe și forma gurii după voce.
"""
import hashlib, json, math, os, random, shutil, subprocess, sys

HERE = os.path.dirname(os.path.abspath(__file__))
S = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(S, "avatar2"))
import anim  # noqa: E402  (voice_features, onsets, Spring, noise)

WORK = os.path.join(S, "work_skit")
OUT = sys.argv[1] if len(sys.argv) > 1 else WORK
os.makedirs(WORK, exist_ok=True)
FPS = 30
PIPER = os.path.join(S, "tts", "bin", "piper")
VOICE = os.path.join(S, "voice", "ro.onnx")
MUSIC = os.path.join(S, "music", "Life_of_Riley.mp3")

CS = 0.78                      # scara personajului
GROUND, ROOF = 1792, 1128      # linia tălpilor: iarbă / acoperiș
LADDER_X, GUTTER_Y = 912, 1160
PATCH = [(348, 1116), (444, 1116), (348, 1072), (444, 1072)]   # centrele țiglelor puse
HIP_X, SH_X, SH_Y, NECK_Y = 40, 78, -205, -236
THIGH, SHIN, UARM, FARM, SOLE = 100, 92, 90, 84, 23
HEAD_TOP = (326 - 70) * 0.86

LINES = dict(
    salut=("Salut! Eu sunt <b>Nea Bilzi</b>. Azi vă arăt cum se pune un acoperiș!",
           "Salut! Eu sunt Nea Bilzi, de la Bilzi Stil Profail. Azi vă arăt cum se pune un acoperiș!"),
    carte=("Țiglă cu țiglă, <b>ca la carte!</b>", None),
    frumos=("Ia uite ce frumos...", None),
    aoleu=("<b>Aoleu!</b>", "Aoleu!"),
    jgheab=("Stai liniștit... <b>jgheabul ține!</b>", "Stai liniștit... jgheabul ține!"),
    uff=("Uff! <b>Nici n-am transpirat.</b>", "Uff! Nici n-am transpirat."),
    final=("Vrei și tu acoperiș nou? <b>Sună-l pe Nea Bilzi!</b>", None),
)


# ---------------------------------------------------------------- utilitare
def run(cmd, **kw):
    r = subprocess.run(cmd, capture_output=True, text=True, **kw)
    if r.returncode:
        print(" ".join(map(str, cmd))[:1500]); print(r.stderr[-3000:]); sys.exit(1)
    return r


def dur(p):
    return float(subprocess.check_output(["ffprobe", "-v", "error", "-show_entries", "format=duration",
                                          "-of", "csv=p=0", p], text=True))


def tts(say):
    h = hashlib.md5(say.encode()).hexdigest()[:12]
    raw, out = os.path.join(WORK, f"v_{h}_raw.wav"), os.path.join(WORK, f"v_{h}.wav")
    if not os.path.exists(out):
        run([PIPER, "-m", VOICE, "-f", raw, "--length-scale", "0.92", "--sentence-silence", "0.12"], input=say)
        run(["ffmpeg", "-y", "-v", "error", "-i", raw, "-af",
             "silenceremove=start_periods=1:start_threshold=-45dB,areverse,"
             "silenceremove=start_periods=1:start_threshold=-45dB,areverse,"
             "highpass=f=80,equalizer=f=200:t=q:w=1:g=2,equalizer=f=3500:t=q:w=1:g=3,aresample=44100",
             "-ac", "1", out])
    return out, dur(out)


def ease(x):
    x = max(0.0, min(1.0, x)); return x * x * (3 - 2 * x)


def lerp(a, b, k):
    return a + (b - a) * k


def rot(v, deg):
    r = math.radians(deg); c, s = math.cos(r), math.sin(r)
    return (v[0] * c - v[1] * s, v[0] * s + v[1] * c)


# ---------------------------------------------------------------- poze
STAND = dict(lean=0, hr=0, aL1=22, aL2=138, aL3=0, hL="relax", hw=-28, aR1=8, aR2=8, aR3=0, hR="relax",
             lgL=(4, 0, 0, 0), lgR=(4, 0, 0, 0), hammer="hand", look=None, turn=None)


def P(**kw):
    d = dict(STAND); d.update(kw); return d


def fk(p, pelvis):
    """Cinematică directă: din unghiuri -> coordonatele articulațiilor."""
    lean = p["lean"]
    j = dict(pelvis=pelvis, lean=lean)
    j["neck"] = tuple(a + b for a, b in zip(pelvis, rot((0, NECK_Y * CS), lean)))
    for sd, side in (("L", -1), ("R", 1)):
        a, lift, k, foot = p["lg" + sd]
        hip = tuple(x + y for x, y in zip(pelvis, rot((side * HIP_X * CS, 0), lean * 0.3)))
        ra = math.radians(a)
        tl = THIGH * CS * (1 - 0.5 * lift)
        knee = (hip[0] + side * math.sin(ra) * tl, hip[1] + math.cos(ra) * tl)
        rb = math.radians(a + k)
        ankle = (knee[0] + side * math.sin(rb) * SHIN * CS, knee[1] + math.cos(rb) * SHIN * CS)
        j["leg" + sd] = dict(hip=hip, knee=knee, ankle=ankle, side=side, footRot=side * foot)
        a1, a2, a3 = p["a" + sd + "1"], p["a" + sd + "2"], p["a" + sd + "3"]
        sh = tuple(x + y for x, y in zip(pelvis, rot((side * SH_X * CS, SH_Y * CS), lean)))
        r1 = math.radians(a1 - side * lean)
        el = (sh[0] + side * math.sin(r1) * UARM * CS, sh[1] + math.cos(r1) * UARM * CS)
        r2 = math.radians(a1 + a2 - side * lean)
        wr = (el[0] + side * math.sin(r2) * FARM * CS, el[1] + math.cos(r2) * FARM * CS)
        hrot = -side * (a1 + a2 + a3 - side * lean)
        j["arm" + sd] = dict(sh=sh, el=el, wr=wr, hand=dict(p=wr, rot=hrot, sc=0.8 * CS, side=side, shape=p["h" + sd]))
    return j


def ground_pelvis(p, x, support):
    """Poziția bazinului astfel încât talpa cea mai de jos să stea pe linia de sprijin."""
    j = fk(p, (x, 0))
    low = max(j["legL"]["ankle"][1], j["legR"]["ankle"][1])
    return (x, support - low - SOLE * CS)


def blend_pose(a, b, k):
    out = {}
    for key, va in a.items():
        vb = b[key]
        if isinstance(va, (int, float)) and isinstance(vb, (int, float)):
            out[key] = lerp(va, vb, k)
        elif isinstance(va, tuple) and isinstance(vb, tuple):
            out[key] = tuple(lerp(x, y, k) for x, y in zip(va, vb))
        else:
            out[key] = vb if k > 0.5 else va
    return out


# ---------------------------------------------------------------- acțiuni
def walk(x0, x1, T, back=False, support=GROUND):
    def f(u):
        ph = 2 * math.pi * u / 0.62
        s = math.sin(ph)
        p = P(lgL=(4 + max(0, s) * 8, max(0, s) * 0.75, -max(0, s) * 10, 6),
              lgR=(4 + max(0, -s) * 8, max(0, -s) * 0.75, -max(0, -s) * 10, 6),
              aR1=10 + 16 * math.sin(ph + math.pi), aR2=12 + 6 * math.sin(ph), aL2=138 + 6 * s,
              lean=2.5 * math.sin(ph), hr=-2 * math.sin(ph))
        x = lerp(x0, x1, ease(u / T) if back else min(1, u / T) * 0.92 + ease(u / T) * 0.08)
        return p, ("ground", x, support, -abs(s) * 7)
    return f


def standing(x, support, poses):
    """poses: listă (t_rel, poză) – se trece lin între ele; poza poate fi funcție de u."""
    def f(u):
        cur = poses[0][1]
        for tr, ps in poses:
            if u >= tr:
                cur = ps
        p = cur(u) if callable(cur) else cur
        return p, ("ground", x, support, 0)
    return f


def wave_pose(u):
    w = math.sin(2 * math.pi * 1.7 * u)
    return P(aR1=130 + 4 * w, aR2=40 + 22 * w, aR3=14 * math.sin(2 * math.pi * 1.7 * u - 1), hR="open", hr=-3)


def explain_pose(u):
    b = math.sin(2 * math.pi * 1.3 * u)
    return P(aR1=36 + 6 * b, aR2=70 + 10 * b, aR3=-20, hR="open", hr=2 * b)


def thumb_pose(u):
    return P(aR1=56, aR2=104 + 4 * math.sin(u * 6), aR3=8, hR="thumb", hr=-2)


def climb(T, y0, y1):
    steps = 5
    def f(u):
        k = u / T
        n = k * steps
        st = int(min(steps - 1, n)); fr = n - st
        y = lerp(y0, y1, (st + ease(fr)) / steps)
        alt = 1 if st % 2 == 0 else -1
        reach = math.sin(math.pi * fr)
        p = P(aR1=150 + alt * 14 * reach, aR2=8, hR="relax", aL1=150 - alt * 14 * reach, aL2=8, hL="relax",
              lgL=(6, 0.55 * (reach if alt > 0 else 0), 4, 0), lgR=(6, 0.55 * (reach if alt < 0 else 0), 4, 0),
              hammer="belt", hr=2 * alt * reach)
        return p, ("pelvis", LADDER_X, y)
    return f


def hop(T, a, b):
    def f(u):
        k = ease(u / T)
        x, y = lerp(a[0], b[0], k), lerp(a[1], b[1], k) - 150 * 4 * (u / T) * (1 - u / T)
        p = P(aR1=145, aR2=20, hR="open", aL1=145, aL2=20, hL="open", lgL=(10, 0.6, 10, 10), lgR=(10, 0.6, 10, 10),
              hammer="belt", hr=-4)
        return p, ("pelvis", x, y)
    return f


STRIKES = [0.55, 1.15, 1.75, 2.35]
RAISED = dict(aL1=96, aL2=72, hw=-15)
DOWN = dict(aL1=38, aL2=18, hw=-158)


def hammering(x, T):
    def f(u):
        k = 1.0                                   # 1 = ciocanul sus, 0 = lovește
        for i, s0 in enumerate(STRIKES):
            d = u - s0
            if -0.42 <= d < 0:
                k = ease((d + 0.42) / 0.42) if i > 0 else 1.0
                break
            if 0 <= d < 0.07:
                k = 1 - d / 0.07
                break
            if 0.07 <= d < 0.18:
                k = 0.0
                break
            if i == len(STRIKES) - 1 and d >= 0.18:
                k = ease((d - 0.18) / 0.4)
        arm = {key: lerp(DOWN[key], RAISED[key], k) for key in RAISED}
        p = P(lgL=(16, 0.85, 26, 14), lgR=(16, 0.85, 26, 14), aR1=26, aR2=38, hR="relax", hr=6,
              look=(-0.85, 0.75), turn=-0.45, **arm)
        return p, ("ground", x, ROOF, 0)
    return f


def slip(T, x0):
    def f(u):
        k = u / T
        fl = 15 * u
        p = P(lean=14 * math.sin(fl * 0.9), aR1=145 + 35 * math.sin(fl), aR2=30 + 20 * math.sin(fl * 1.3), hR="open",
              aL1=140 + 35 * math.sin(fl + 1.6), aL2=30, hL="open", hammer="free",
              lgR=(40 * min(1, k * 3), 0.2, -10, 25), lgL=(6, 0, 6, -10), hr=-8 * math.sin(fl * 0.8), look=(0, -0.6))
        return p, ("ground", x0 + 30 * ease(k), ROOF + 10 * ease(k), 0)
    return f


HANG_POSE = dict(aL1=172, aL2=-2, hL="relax", aR1=172, aR2=-2, hR="relax", hammer="free")


def hang(T, gx):
    def f(u):
        sw = math.sin(2 * math.pi * 0.85 * u)
        p = P(lgL=(6 + 10 * sw, 0.1, 8 + 6 * sw, 10), lgR=(6 - 10 * sw, 0.1, 8 - 6 * sw, 10),
              lean=3 * sw, hr=-3 * sw, **HANG_POSE)
        return p, ("hang", gx, GUTTER_Y)
    return f


def fall(T, x0, gx):
    """De pe acoperiș până agățat de jgheab (bazinul coboară, mâinile urcă)."""
    def f(u):
        k = ease(u / T)
        p = P(lgL=(6, 0.3 * (1 - k), 8, 10), lgR=(20 * (1 - k), 0.3, 0, 10), hr=-6, look=(0, -0.7), **HANG_POSE)
        return p, ("hang", lerp(x0 + 30, gx, k), GUTTER_Y - 120 * (1 - k))
    return f


def pullup(T, gx, x1):
    def f(u):
        k = u / T
        if k < 0.55:
            q = ease(k / 0.55)
            p = P(aL1=172 - 20 * q, aL2=-2 + 130 * q, aR1=172 - 20 * q, aR2=-2 + 130 * q, hL="relax", hR="relax",
                  hammer="free", lgL=(6, 0.5 * q, 10, 10), lgR=(12, 0.3 * q, 10, 10), hr=4 * q)
            return p, ("hang", gx, GUTTER_Y + 0 * q - 0)  # se ridică pe brațe (unghiurile de mai sus)
        q = ease((k - 0.55) / 0.45)
        p = P(hammer="free", aL1=20, aL2=40, hL="relax", aR1=30, aR2=40, lgL=(10, 0.7 * (1 - q), 10, 0),
              lgR=(10, 0.4 * (1 - q), 10, 0))
        return p, ("ground", lerp(gx, x1, q), ROOF, -60 * (1 - q))
    return f


def wipe(T, x):
    def f(u):
        w = math.sin(2 * math.pi * 1.6 * u)
        p = P(aR1=122 + 10 * w, aR2=118, aR3=10, hR="open", aL1=48, aL2=-95, hL="relax", hammer="free",
              hr=-5 + 2 * w, look=(0.2, -0.2))
        return p, ("ground", x, ROOF, 0)
    return f


# ---------------------------------------------------------------- regia
def build():
    voice = {k: tts(v[1] or v[0].replace("<b>", "").replace("</b>", "")) for k, v in LINES.items()}
    D = {k: v[1] for k, v in voice.items()}
    seq, lines, sfx, ev = [], [], [], {}
    t = 0.0

    def add(fn, T, name=None):
        nonlocal t
        seq.append(dict(t0=t, t1=t + T, fn=fn, name=name))
        if name:
            ev[name] = t
        t += T

    def say(key, at, gest=None):
        lines.append(dict(key=key, start=at, dur=D[key], text=LINES[key][0], wav=voice[key][0]))

    X_TALK, X_KNEEL, X_BACK, X_HANG = 470, 560, 690, 740
    add(walk(-220, X_TALK, 2.3), 2.3, "walkin")
    say("salut", t + 0.2)
    T1 = D["salut"] + 0.9
    add(standing(X_TALK, GROUND, [(0, wave_pose), (min(2.0, T1 * 0.45), explain_pose), (T1 - 0.6, P())]), T1)
    add(walk(X_TALK, LADDER_X, 1.5), 1.5)
    p0 = ground_pelvis(P(), LADDER_X, GROUND)
    top = ground_pelvis(P(), LADDER_X, 1150)
    add(climb(2.8, p0[1], top[1] - 20), 2.8, "climb")
    roofp = ground_pelvis(P(), X_KNEEL, ROOF)
    add(hop(0.75, (LADDER_X, top[1] - 20), roofp), 0.75, "hop")
    add(standing(X_KNEEL, ROOF, [(0, P())]), 0.35)
    add(hammering(X_KNEEL, 2.75), 2.75, "hammer")
    for i, s0 in enumerate(STRIKES):
        sfx.append(("toc", ev["hammer"] + s0))
    say("carte", t + 0.15)
    T2 = D["carte"] + 0.7
    add(standing(X_KNEEL, ROOF, [(0, thumb_pose), (T2 - 0.3, P())]), T2)
    add(walk(X_KNEEL, X_BACK, 1.0, back=True, support=ROOF), 1.0)
    say("frumos", t + 0.05)
    T3 = D["frumos"] + 0.25
    add(standing(X_BACK, ROOF, [(0, P(aR1=50, aR2=-100, hR="relax", look=(-0.7, 0.5), turn=-0.4, hr=5))]), T3)
    say("aoleu", t + 0.1)
    add(slip(0.95, X_BACK), 0.95, "slip")
    add(fall(0.4, X_BACK, X_HANG), 0.4, "fall")
    sfx.append(("thud", ev["slip"] + 1.15))
    say("jgheab", t + 0.5)
    T4 = D["jgheab"] + 1.4
    add(hang(T4, X_HANG), T4, "hang")
    add(pullup(1.2, X_HANG, X_HANG - 40), 1.2, "pullup")
    say("uff", t + 0.25)
    T5 = D["uff"] + 0.8
    add(wipe(T5, X_HANG - 40), T5, "wipe")
    say("final", t + 0.15)
    T6 = D["final"] + 3.0
    add(standing(X_HANG - 40, ROOF, [(0, thumb_pose), (D["final"] + 0.4, lambda u: dict(wave_pose(u), hammer="free",
                                                                                      aL1=48, aL2=-95))]), T6, "final")
    return seq, lines, sfx, ev, t


def main():
    seq, lines, sfx, ev, total = build()
    n = int(round(total * FPS))
    # pista de voce
    vin, parts = [], []
    for i, ln in enumerate(lines):
        vin += ["-i", ln["wav"]]
        ms = int(ln["start"] * 1000)
        parts.append(f"[{i}:a]adelay={ms}|{ms}[d{i}]")
    parts.append("".join(f"[d{i}]" for i in range(len(lines))) +
                 f"amix=inputs={len(lines)}:normalize=0,apad,atrim=duration={total:.3f}[v]")
    vwav = os.path.join(WORK, "voice.wav")
    run(["ffmpeg", "-y", "-v", "error", *vin, "-filter_complex", ";".join(parts), "-map", "[v]",
         "-ac", "1", "-ar", "44100", vwav])
    opn, wide = anim.voice_features(vwav, total)
    ons = dict(anim.onsets(opn))

    rnd = random.Random(5)
    blinks, tb = [], 0.8
    while tb < total:
        blinks.append(tb)
        if rnd.random() < 0.2:
            blinks.append(tb + 0.3)
        tb += rnd.uniform(2.0, 4.5)
    sacc, ts = [], 0.0
    while ts < total:
        sacc.append((ts, rnd.uniform(-0.3, 0.3), rnd.uniform(-0.2, 0.15))); ts += rnd.uniform(0.7, 2.0)

    Sp = anim.Spring
    sp = dict(nod=Sp(0, 4.5, 0.35), lx=Sp(0, 9, 0.9), ly=Sp(0, 9, 0.9), turn=Sp(0, 2.5, 0.7), raise_=Sp(0, 4, 0.5),
              smile=Sp(0.35, 3, 0.8), mo=Sp(0, 14, 0.75), mw=Sp(1, 10, 0.8), px=Sp(0, 6, 0.7), py=Sp(0, 6, 0.7))
    frames = []
    hammer_free = None     # (x, y, vx, vy, rot, vr) după ce scapă ciocanul
    prev_pose, prev_pel = None, None
    seg_idx = -1
    blend_from = None
    for fi in range(n):
        t = fi / FPS
        si = max(i for i, sg in enumerate(seq) if sg["t0"] <= t + 1e-9)
        sg = seq[si]
        u = t - sg["t0"]
        p, root = sg["fn"](u)
        # bazinul
        if root[0] == "ground":
            pel = ground_pelvis(p, root[1], root[2])
            pel = (pel[0], pel[1] + root[3])
        elif root[0] == "pelvis":
            pel = (root[1], root[2])
        else:  # agățat: mâinile pe jgheab
            j0 = fk(p, (0, 0))
            wm = ((j0["armL"]["wr"][0] + j0["armR"]["wr"][0]) / 2, (j0["armL"]["wr"][1] + j0["armR"]["wr"][1]) / 2)
            pel = (root[1] - wm[0], root[2] - wm[1] + 14)
        # tranziție lină la schimbarea acțiunii
        if si != seg_idx:
            if prev_pose is not None:
                blend_from = (prev_pose, prev_pel, t)
            seg_idx = si
        if blend_from:
            bp, bpel, bt = blend_from
            k = ease((t - bt) / 0.24)
            if k < 1:
                p = blend_pose(bp, p, k)
                pel = (lerp(bpel[0], pel[0], k), lerp(bpel[1], pel[1], k))
            else:
                blend_from = None
        prev_pose, prev_pel = p, pel
        j = fk(p, pel)

        # ciocanul: în mână / la brâu / liber (cade)
        if hammer_free is not None and p["hammer"] != "free":
            p = dict(p, hammer="free", aL1=48, aL2=-95, hL="relax")
            j = fk(p, pel)
        wl = j["armL"]["wr"]
        if p["hammer"] == "hand":
            hrot = p["hw"]
            j["hammer"] = dict(p=wl, rot=hrot)
        elif p["hammer"] == "belt":
            j["hammer"] = dict(p=(pel[0] - 58 * CS, pel[1] + 30 * CS), rot=168)
        else:
            if hammer_free is None:
                hammer_free = [wl[0], wl[1], -260, -520, frames[-1]["j"]["hammer"]["rot"] if frames else 0, -540]
            hx, hy, vx, vy, hrr, vr = hammer_free
            if hy < GROUND - 8:
                dt = 1 / FPS
                vy += 1900 * dt; hx += vx * dt; hy += vy * dt; hrr += vr * dt
                if hy >= GROUND - 8:
                    hy, vx, vy, vr = GROUND - 8, 0, 0, 0; hrr = -100
                hammer_free = [hx, hy, vx, vy, hrr, vr]
            j["hammer"] = dict(p=(hammer_free[0], hammer_free[1]), rot=hammer_free[4])

        # replica curentă
        cur = None
        for ln in lines:
            if ln["start"] - 0.2 <= t <= ln["start"] + ln["dur"] + 0.7:
                cur = ln
        talking = cur is not None and cur["start"] <= t <= cur["start"] + cur["dur"]
        o_t = opn[fi] if talking else 0.0
        w_t = (0.78 + 0.42 * wide[fi]) if talking else 1.0
        if talking and o_t > 0.5 and wide[fi] < 0.25:
            w_t = 0.72
        mo, mw = sp["mo"].step(o_t), sp["mw"].step(w_t)
        if talking and fi in ons:
            sp["nod"].kick(40 * ons[fi])
            if ons[fi] > 0.75:
                sp["raise_"].kick(5 * ons[fi])
        nod = sp["nod"].step(0)

        lx_t, ly_t = 0, 0
        for st, sx, sy in sacc:
            if t >= st:
                lx_t, ly_t = sx, sy
        if talking:
            lx_t *= 0.3; ly_t *= 0.3
        if p.get("look"):
            lx_t, ly_t = p["look"]
        turn_t = p.get("turn") if p.get("turn") is not None else 0.15 * anim.noise(t * 0.5, 3)
        lx, ly, turn = sp["lx"].step(lx_t), sp["ly"].step(ly_t), sp["turn"].step(turn_t)
        name = sg["name"] or ""
        raise_t = 0.3 if talking else 0
        tilt = 0.0
        if name in ("slip", "fall"):
            raise_t, tilt = 1.0, 0.0
        if name == "hang":
            raise_t, tilt = 0.7, -0.45
        braise = sp["raise_"].step(raise_t)
        smile_t = 0.15 if talking else 0.4
        if cur and not talking and t > cur["start"] + cur["dur"]:
            smile_t = 0.85
        if name == "hammer":
            smile_t = 0.0
        if name in ("slip", "fall"):
            smile_t = 0.0
        sm = sp["smile"].step(smile_t)
        if name in ("slip", "fall") and not talking:
            mo = max(mo, 0.7); mw = 0.7                       # gura „O” de spaimă
        close = 0.0
        for bt in blinks:
            d = t - bt
            if 0 <= d < 0.17:
                close = math.sin(math.pi * d / 0.17)
        if name == "wipe":
            close = max(close, 0.55)

        # țiglele puse
        tiles = []
        for i, s0 in enumerate(STRIKES):
            d = t - (ev["hammer"] + s0 + 0.05)
            tiles.append(0 if d < 0 else min(1.0, d / 0.25))
        # efecte
        fx = {}
        for i, s0 in enumerate(STRIKES):
            d = t - (ev["hammer"] + s0)
            if 0 <= d < 0.5:
                cx, cy = PATCH[i]
                fx["toc"] = dict(x=cx - 10, y=cy - 70 - d * 60, s=0.6 + 0.6 * min(1, d / 0.12), r=-8 + i * 6,
                                 a=1 - max(0, (d - 0.3) / 0.2))
                fx["dust"] = [(cx + math.cos(a) * (14 + d * 150), cy + 10 + math.sin(a) * (6 + d * 60) - d * 40,
                               6 * (1 - d * 1.6) + 1, max(0, 0.9 - d * 1.8)) for a in (2.6, 3.0, 3.5, 0.3, -0.2)]
        if name in ("slip", "fall"):
            hx, hy = j["neck"][0], j["neck"][1] - HEAD_TOP * CS - 50
            fx["excl"] = dict(x=hx + 70, y=hy, s=1 + 0.15 * math.sin(t * 30))
            fx["motion"] = [(pel[0] + dx, pel[1] - 260 + dy, pel[0] + dx, pel[1] - 320 + dy)
                            for dx, dy in ((-120, 0), (120, 20), (-90, 160), (110, 180))]
        if name in ("hang", "wipe"):
            d = (t - ev[name]) % 1.1
            fx["sweat"] = dict(x=j["neck"][0] + 75, y=j["neck"][1] - 150 + d * 70, s=1.1, a=max(0, 1 - d / 1.1))

        # porumbelul: ciugulește; la alunecare zboară și revine
        pg = dict(x=0, y=0, head=-3 * max(0, math.sin(t * 5)) ** 8 * 3, wing=0)
        ds = t - ev["slip"]
        if 0 <= ds < 2.2:
            up = math.sin(math.pi * min(1, ds / 2.2))
            pg = dict(x=-30 * up, y=-110 * up, head=0, wing=-40 * abs(math.sin(ds * 28)))

        # balon
        bub = None
        if cur:
            le = cur["start"] + cur["dur"]
            show = min(1.0, (t - cur["start"] + 0.15) / 0.25)
            if t > le + 0.45:
                show = max(0.0, 1 - (t - le - 0.45) / 0.25)
            if show > 0:
                head_top = j["neck"][1] - HEAD_TOP * CS - 30
                bub = dict(html=cur["text"], show=show, ax=j["neck"][0], ay=min(head_top, 1180))
        card = 0.0
        fin = [ln for ln in lines if ln["key"] == "final"][0]
        if t > fin["start"] + fin["dur"] + 0.5:
            card = (t - fin["start"] - fin["dur"] - 0.5) / 0.35

        breath = 0.012 * (0.5 + 0.5 * math.sin(2 * math.pi * t / 3.4)) + (0.004 * mo if talking else 0)
        frames.append(dict(t=round(t, 3), cs=CS, j=j, breath=breath,
                           head=dict(r=p["hr"] + (2.5 * math.sin(t * 6) * mo if talking else 0), nod=nod, turn=turn, up=-ly * 0.3),
                           eyes=dict(lx=lx, ly=ly, close=close), brow={"raise": braise, "tilt": tilt},
                           mouth=dict(o=mo, w=mw, s=sm), tiles=tiles, pigeon=pg, fx=fx, bubble=bub, card=card))

    fjson = os.path.join(WORK, "frames.json")
    with open(fjson, "w", encoding="utf-8") as f:
        json.dump(frames, f, ensure_ascii=False, default=lambda o: list(o))
    meta = dict(total=total, lines=[(l["key"], l["start"], l["dur"]) for l in lines], sfx=sfx)
    with open(os.path.join(WORK, "meta.json"), "w") as f:
        json.dump(meta, f)
    print(f"{len(frames)} cadre, {total:.1f}s")
    return frames, total, sfx, vwav


if __name__ == "__main__":
    main()
