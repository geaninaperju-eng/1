#!/usr/bin/env python3
"""Generează câte un video TikTok (1080x1920) pentru fiecare proiect din portofoliul BILZI,
prezentat de mascota „Nea Bilzi” (personaj animat, voce sintetizată, replici cu umor)."""
import array, hashlib, json, math, os, shutil, subprocess, sys, wave

S = os.path.dirname(os.path.abspath(__file__))
IMG = os.path.join(S, "img")
WORK = os.path.join(S, "work")
AV = os.path.join(S, "avatar")
OUT = sys.argv[1]
ONLY = sys.argv[2:]          # opțional: slug-uri de randat
os.makedirs(WORK, exist_ok=True)
os.makedirs(OUT, exist_ok=True)

W, H, FPS = 1080, 1920, 30
FW, FH = 1080, 1350          # fereastra foto (4:5)
FY = 556                     # sub siglă și banda de text
OY = 900                     # stratul cu personajul (1080x1020) se suprapune de la y=900
LOGO = os.path.join(S, "logo", "badge.png")
LOGO_W = 220
MUSIC = os.path.join(S, "music")
PIPER = os.path.join(S, "tts", "bin", "piper")
VOICE = os.path.join(S, "voice", "ro.onnx")
XF = 0.5                     # durata tranziției
MIN_CLIP = 2.4               # durata minimă a unei poze
LEAD = 0.35                  # personajul începe să vorbească la 0,35 s după intrarea pozei
BOLD = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
REG = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
ACCENT = "0xE30613"          # roșu Bilka

# Replici: text = ce apare în balon (HTML, <b> = portocaliu), say = ce pronunță vocea
# (scris fonetic unde e nevoie), g = gest: wave / thumb / idle.
def L(text, say=None, g="thumb"):
    return dict(text=text, say=say or text.replace("<b>", "").replace("</b>", ""), g=g)

INTRO_SAY = "Bilzi Stil Profail"
PROIECTE = [
    dict(slug="01_Ciocarlia_Bilka_Balcanic_8004",
         titlu="Casă în Ciocârlia",
         model="Țiglă metalică Bilka Balcanic 8004 MAT",
         loc="Ciocârlia, Ialomița",
         muzica=("Life_of_Riley.mp3", 0),
         poze=[("proiect1-img1.jpg", L("Salut! Eu sunt <b>Nea Bilzi</b>. Azi vă duc la Ciocârlia!",
                                        "Salut! Eu sunt Nea Bilzi, de la " + INTRO_SAY + ". Azi vă duc la Ciocârlia!", "wave")),
               ("proiect1-img2-1.jpg", L("Țiglă Bilka Balcanic, <b>roșie ca ardeii din grădină</b>.")),
               ("proiect1-img4-1.jpg", None),
               ("proiect1-img3-1.jpg.webp", L("Vecinul s-a uitat peste gard și a zis doar atât: <b>vreau și eu!</b>")),
               ("proiect1-img5-1.jpg.webp", None),
               ("proiect1-img6-1-rotated.jpg.webp", L("Văzut de sus e și mai frumos. Porumbeii i-au dat <b>cinci stele!</b>"))],
         final=L("Vrei și tu un acoperiș ca ăsta? <b>Sună-l pe Nea Bilzi!</b>",
                 "Vrei și tu un acoperiș ca ăsta? Sună-l pe Nea Bilzi! Venim cu drag.", "wave")),
    dict(slug="02_Afumati_Bilka_Clasic_8019",
         titlu="Casă în Afumați",
         model="Țiglă metalică Bilka Clasic 8019 Grande Mat",
         loc="Afumați, Ilfov",
         muzica=("Wallpaper.mp3", 0),
         poze=[("proiect2-img1.jpg", L("Salutare! <b>Nea Bilzi</b> aici. Azi suntem la Afumați, lângă București.",
                                        "Salutare! Nea Bilzi aici, de la " + INTRO_SAY + ". Azi suntem la Afumați, lângă București.", "wave")),
               ("proiect2-img2-1.jpg.webp", L("Bilka Clasic, gri antracit. <b>Elegant ca un costum de nuntă.</b>")),
               ("proiect2-img3.jpg", L("Doar că ăsta <b>nu se pătează</b> la prima horă!")),
               ("proiect2-img3-1.jpg.webp", L("Multe colțuri, multe pante. Le-am luat pe toate <b>la rând, cu răbdare</b>.")),
               ("proiect2-img4.jpg.webp", None)],
         final=L("Și casa ta merită o haină nouă. <b>Sună-ne și venim!</b>", None, "wave")),
    dict(slug="03_Lilieci_Camin_Cultural_Bilka_Clasic_8017",
         titlu="Cămin cultural Lilieci",
         model="Țiglă metalică Bilka Clasic 8017 MAT",
         loc="Lilieci, Ialomița",
         muzica=("Inspired.mp3", 0),
         poze=[("bilka-materiale-livrare.jpg.webp", L("Salut, sunt <b>Nea Bilzi</b>! A venit marfa: Bilka, direct la poartă.",
                                                     "Salut, sunt Nea Bilzi, de la " + INTRO_SAY + ". A venit marfa: Bilka, direct la poartă!", "wave")),
               ("proiect3-img1.jpg.webp", L("Căminul cultural din Lilieci, cu <b>acoperiș nou-nouț</b>.")),
               ("proiect3-img4.jpg", L("De acum, la hora satului nu mai plouă. <b>Doar cu confetti!</b>")),
               ("proiect3-img3-1.jpg.webp", None),
               ("proiect3-img4-1.jpg.webp", L("Toată lumea zâmbește, <b>satul dansează</b>."))],
         final=L("Ai o clădire care plânge când plouă? <b>Sună-l pe Nea Bilzi!</b>", None, "wave")),
    dict(slug="04_Bilka_Iberic_Negru_Pluvial_Alb",
         titlu="Casă nouă, de la zero",
         model="Bilka Iberic negru · sistem pluvial alb",
         loc="Montaj complet acoperiș",
         muzica=("Funkorama.mp3", 12),
         poze=[("proiect4-img1-1.jpg.webp", L("Salut, sunt <b>Nea Bilzi</b>! Aici am pornit de la zero, la o casă nouă.",
                                               "Salut, sunt Nea Bilzi, de la " + INTRO_SAY + ". Aici am pornit de la zero, la o casă nouă.", "wave")),
               ("proiect4-img3.jpg.webp", L("Colegul verifică fiecare țiglă. Nu, <b>nu face TikTok</b>, muncește!",
                                            "Colegul verifică fiecare țiglă. Nu, nu face tic toc, muncește!")),
               ("proiect4-img2.jpg.webp", L("Bilka Iberic neagră, cu <b>jgheaburi albe</b>.")),
               ("proiect4-img4.jpg.webp", L("Negru cu alb, ca un smoching. <b>Casa e gata de gală!</b>"))],
         final=L("Vrei și tu o casă la patru ace? <b>Sună-ne, venim cu drag!</b>", None, "wave")),
]


def run(cmd, **kw):
    r = subprocess.run(cmd, capture_output=True, text=True, **kw)
    if r.returncode:
        print(" ".join(cmd)); print(r.stderr[-3000:]); sys.exit(1)
    return r


def size(path):
    o = subprocess.check_output(["ffprobe", "-v", "error", "-select_streams", "v:0",
                                 "-show_entries", "stream=width,height", "-of", "csv=p=0", path], text=True)
    w, h = map(int, o.strip().split(",")[:2])
    return w, h


def duration(path):
    return float(subprocess.check_output(["ffprobe", "-v", "error", "-show_entries", "format=duration",
                                          "-of", "csv=p=0", path], text=True))


def txt(name, s):
    p = os.path.join(WORK, name)
    with open(p, "w", encoding="utf-8") as f:
        f.write(s)
    return p


def tts(say):
    """Sintetizează o replică (cu cache) și întoarce (cale wav, durată)."""
    h = hashlib.md5(say.encode()).hexdigest()[:12]
    raw, out = os.path.join(WORK, f"v_{h}_raw.wav"), os.path.join(WORK, f"v_{h}.wav")
    if not os.path.exists(out):
        run([PIPER, "-m", VOICE, "-f", raw, "--length-scale", "0.92", "--sentence-silence", "0.15"], input=say)
        # voce puțin mai caldă și mai prezentă; eliminăm liniștea de la capete
        run(["ffmpeg", "-y", "-v", "error", "-i", raw, "-af",
             "silenceremove=start_periods=1:start_threshold=-45dB,areverse,"
             "silenceremove=start_periods=1:start_threshold=-45dB,areverse,"
             "highpass=f=80,equalizer=f=200:t=q:w=1:g=2,equalizer=f=3500:t=q:w=1:g=3,"
             "aresample=44100", "-ac", "1", out])
    return out, duration(out)


def clip(src, dst, idx, dur):
    """O poză -> clip 1080x1920 cu fundal blurat și mișcare lentă."""
    w, h = size(src)
    n = int(dur * FPS)
    bg = (f"[0:v]scale={W}:{H}:force_original_aspect_ratio=increase,crop={W}:{H},"
          f"boxblur=40:3,eq=brightness=-0.18:saturation=0.8[bg]")
    if w / h > FW / FH * 1.15:
        a, b = (0.15, 0.85) if idx % 2 == 0 else (0.85, 0.15)
        fg = (f"[0:v]scale=-2:{FH*2}:flags=lanczos,"
              f"crop={FW*2}:{FH*2}:x='(iw-ow)*({a}+({b}-{a})*t/{dur:.3f})':y=0,"
              f"scale={FW}:{FH}:flags=lanczos[fg]")
    else:
        fg = (f"[0:v]scale={FW*3}:{FH*3}:force_original_aspect_ratio=increase:flags=lanczos,"
              f"crop={FW*3}:{FH*3},"
              f"zoompan=z='1+0.10*in/{n}':x='iw/2-iw/zoom/2':y='ih/2-ih/zoom/2':d=1:s={FW}x{FH}:fps={FPS}[fg]")
    flt = f"{bg};{fg};[bg][fg]overlay=0:{FY}:shortest=1,format=yuv420p[v]"
    run(["ffmpeg", "-y", "-v", "error", "-loop", "1", "-framerate", str(FPS), "-t", f"{dur:.3f}", "-i", src,
         "-filter_complex", flt, "-map", "[v]", "-t", f"{dur:.3f}", "-r", str(FPS),
         "-c:v", "libx264", "-preset", "medium", "-crf", "16", dst])


def endcard(dst, p, dur):
    """Card final cu datele de contact; jos rămâne loc pentru personaj."""
    l1 = txt("e1.txt", "BILZI STEEL PROFILE")
    l2 = txt("e2.txt", "Montaj acoperișuri · Țiglă metalică")
    l3 = txt("e3.txt", "Cere o ofertă!")
    l4 = txt("e4.txt", "0720 244 244")
    l5 = txt("e5.txt", "www.acoperis-mag.ro")
    y = 345
    last = os.path.join(IMG, p["poze"][0][0])
    flt = (f"[0:v]scale={W}:{H}:force_original_aspect_ratio=increase,crop={W}:{H},"
           f"boxblur=30:3,eq=brightness=-0.35:saturation=0.6,"
           f"drawbox=x=90:y={y+330}:w={W-180}:h=150:color={ACCENT}@1:t=fill,"
           f"drawtext=fontfile={BOLD}:textfile={l1}:fontsize=74:fontcolor=white:x=(w-tw)/2:y={y},"
           f"drawtext=fontfile={REG}:textfile={l2}:fontsize=44:fontcolor=white@0.9:x=(w-tw)/2:y={y+105},"
           f"drawtext=fontfile={BOLD}:textfile={l3}:fontsize=56:fontcolor=white:x=(w-tw)/2:y={y+240},"
           f"drawtext=fontfile={BOLD}:textfile={l4}:fontsize=84:fontcolor=white:x=(w-tw)/2:y={y+360},"
           f"drawtext=fontfile={REG}:textfile={l5}:fontsize=48:fontcolor=white:x=(w-tw)/2:y={y+515},"
           f"fade=t=in:st=0:d=0.4,format=yuv420p[v]")
    run(["ffmpeg", "-y", "-v", "error", "-loop", "1", "-framerate", str(FPS), "-t", f"{dur:.3f}", "-i", last,
         "-filter_complex", flt, "-map", "[v]", "-t", f"{dur:.3f}", "-r", str(FPS),
         "-c:v", "libx264", "-preset", "medium", "-crf", "16", dst])


def envelope(wav_path, total):
    """Deschiderea gurii pe cadru, din energia vocii (0..1)."""
    raw = subprocess.check_output(["ffmpeg", "-v", "error", "-i", wav_path, "-ac", "1", "-ar", "16000",
                                   "-f", "s16le", "-"])
    a = array.array("h", raw)
    hop = 16000 // FPS
    env = []
    for i in range(int(total * FPS) + 1):
        seg = a[i * hop:(i + 1) * hop]
        env.append(math.sqrt(sum(x * x for x in seg) / len(seg)) if len(seg) else 0.0)
    peak = sorted(env)[int(len(env) * 0.97)] or 1.0
    out, prev = [], 0.0
    for e in env:
        v = min(1.0, (e / peak) ** 0.8 * 1.15)
        v = v if v > 0.12 else 0.0
        prev = v if v > prev else prev * 0.55 + v * 0.45   # deschidere rapidă, închidere lină
        out.append(prev)
    return out


def ease_back(x):
    c1 = 1.70158; c3 = c1 + 1
    return 1 + c3 * (x - 1) ** 3 + c1 * (x - 1) ** 2


def video(p):
    slug = p["slug"]
    print("->", slug)
    # 1) replici + durate
    lines, durs = [], []
    for f, ln in p["poze"]:
        if ln:
            wav, d = tts(ln["say"])
            lines.append(dict(ln, wav=wav, dur=d))
            durs.append(max(MIN_CLIP, LEAD + d + 0.55))
        else:
            lines.append(None)
            durs.append(MIN_CLIP)
    fwav, fd = tts(p["final"]["say"])
    final = dict(p["final"], wav=fwav, dur=fd)
    end_dur = max(3.5, 0.45 + fd + 1.6)

    starts, t = [], 0.0
    for d in durs:
        starts.append(t); t += d - XF
    end_start = t
    total = end_start + end_dur

    timeline = []   # (start, line)
    for st, ln in zip(starts, lines):
        if ln:
            timeline.append((st + LEAD, ln))
    timeline.append((end_start + 0.45, final))

    # 2) pista de voce + plicul pentru gură
    vin, vparts = [], []
    for i, (st, ln) in enumerate(timeline):
        vin += ["-i", ln["wav"]]
        ms = int(st * 1000)
        vparts.append(f"[{i}:a]adelay={ms}|{ms}[d{i}]")
    vparts.append("".join(f"[d{i}]" for i in range(len(timeline))) +
                  f"amix=inputs={len(timeline)}:normalize=0,apad,atrim=duration={total:.3f}[v]")
    voice = os.path.join(WORK, f"{slug}_voice.wav")
    run(["ffmpeg", "-y", "-v", "error", *vin, "-filter_complex", ";".join(vparts), "-map", "[v]",
         "-ac", "1", "-ar", "44100", voice])
    env = envelope(voice, total)

    # 3) cadrele personajului
    frames = []
    n_frames = int(round(total * FPS))
    for fi in range(n_frames):
        tt = fi / FPS
        # poziție: colț stânga-jos în timpul pozelor, centru-mare pe cardul final
        k = min(1.0, max(0.0, (tt - end_start) / 0.6))
        k = k * k * (3 - 2 * k)
        ax = 235 + (540 - 235) * k
        ay = 1100.0
        sc = 1.2 + (1.5 - 1.2) * k
        if tt < 0.7:
            ay += (1 - ease_back(tt / 0.7)) * 800
        cur = None
        for st, ln in timeline:
            if st - 0.25 <= tt <= st + ln["dur"] + 0.6 or (ln is final and tt >= st - 0.25):
                cur = (st, ln)
        gesture, arm, bubble, talk = "idle", 0.0, None, False
        if cur:
            st, ln = cur
            le = st + ln["dur"]
            gesture = ln["g"]
            arm = min(1.0, (tt - st + 0.25) / 0.3) if tt < le else max(0.0, 1 - (tt - le) / 0.5)
            if ln is final:
                arm = min(1.0, (tt - st + 0.25) / 0.3)       # pe final continuă să facă cu mâna
            talk = st <= tt <= le
            if ln is not final:
                show = min(1.0, (tt - st + 0.15) / 0.3)
                if tt > le + 0.35:
                    show = max(0.0, 1 - (tt - le - 0.35) / 0.25)
                if show > 0:
                    bubble = dict(html=ln["text"], show=show, x=40, y=318)
        blink = 0.1 if (tt % 3.3) < 0.12 else 1.0
        frames.append(dict(t=tt, ax=ax, ay=ay, sc=sc, mouth=env[min(fi, len(env) - 1)] if talk else 0.0,
                           blink=blink, gesture=gesture, armBlend=arm, talk=talk, bubble=bubble))
    fjson = os.path.join(WORK, f"{slug}_frames.json")
    with open(fjson, "w", encoding="utf-8") as f:
        json.dump(frames, f, ensure_ascii=False)
    fdir = os.path.join(WORK, f"{slug}_ov")
    shutil.rmtree(fdir, ignore_errors=True)
    run(["node", os.path.join(AV, "render.js"), fjson, fdir],
        env=dict(os.environ, NODE_PATH="/opt/node22/lib/node_modules"))

    # 4) clipurile foto + cardul final
    clips = []
    for i, ((f, _), d) in enumerate(zip(p["poze"], durs)):
        c = os.path.join(WORK, f"{slug}_{i}.mp4")
        clip(os.path.join(IMG, f), c, i, d)
        clips.append((c, d))
    e = os.path.join(WORK, f"{slug}_end.mp4")
    endcard(e, p, end_dur)
    clips.append((e, end_dur))

    # 5) montaj final
    inputs, parts = [], []
    for c, _ in clips:
        inputs += ["-i", c]
    trans = ["slideleft", "smoothup", "fade", "slideright", "circleopen", "smoothleft"]
    cur, t = "[0:v]", 0.0
    for i in range(1, len(clips)):
        t += clips[i - 1][1] - XF
        nxt = f"[x{i}]"
        parts.append(f"{cur}[{i}:v]xfade=transition={trans[(i-1) % len(trans)]}:duration={XF}:offset={t:.3f}{nxt}")
        cur = nxt

    titlu = txt(f"{slug}_t.txt", p["titlu"])
    model = txt(f"{slug}_m.txt", p["model"])
    loc = txt(f"{slug}_l.txt", p["loc"])
    on = f"enable='lt(t,{end_start:.2f})'"
    top_y, bot_y = 263, 415
    over = (
        f"drawtext=fontfile={BOLD}:textfile={titlu}:fontsize=66:fontcolor=white:"
        f"shadowcolor=black@0.6:shadowx=3:shadowy=3:x=(w-tw)/2:y={top_y+55}:{on},"
        f"drawtext=fontfile={BOLD}:textfile={model}:fontsize=38:fontcolor=white:box=1:boxcolor={ACCENT}@0.95:"
        f"boxborderw=18:x=(w-tw)/2:y={bot_y}:{on},"
        f"drawtext=fontfile={REG}:textfile={loc}:fontsize=40:fontcolor=white:"
        f"shadowcolor=black@0.6:shadowx=2:shadowy=2:x=(w-tw)/2:y={bot_y+80}:{on}"
    )
    n = len(clips)
    parts.append(f"{cur}{over}[t]")
    parts.append(f"[{n}:v]scale={LOGO_W}:-1:flags=lanczos[lg]")
    parts.append(f"[t][lg]overlay=36:150[t2]")
    parts.append(f"[t2][{n+3}:v]overlay=0:{OY}:shortest=1,format=yuv420p[v]")
    # audio: muzica scade automat când vorbește Nea Bilzi
    mf, mstart = p["muzica"]
    parts.append(f"[{n+1}:a]atrim=start={mstart}:duration={total:.3f},asetpts=PTS-STARTPTS,"
                 f"loudnorm=I=-16:TP=-2,volume=0.6,aresample=44100,aformat=channel_layouts=stereo[mus]")
    parts.append(f"[{n+2}:a]loudnorm=I=-15:TP=-2,aresample=44100,aformat=channel_layouts=stereo,asplit=2[vo][sc]")
    parts.append(f"[mus][sc]sidechaincompress=threshold=0.03:ratio=6:attack=30:release=450[duck]")
    parts.append(f"[duck][vo]amix=inputs=2:normalize=0,loudnorm=I=-14:TP=-1.5:LRA=11,"
                 f"afade=t=in:d=0.3,afade=t=out:st={total-1.5:.3f}:d=1.5,aresample=44100[a]")
    dst = os.path.join(OUT, f"BILZI_TikTok_{slug}.mp4")
    run(["ffmpeg", "-y", "-v", "error", *inputs,
         "-loop", "1", "-i", LOGO, "-i", os.path.join(MUSIC, mf), "-i", voice,
         "-framerate", str(FPS), "-i", os.path.join(fdir, "f_%05d.png"),
         "-filter_complex", ";".join(parts), "-map", "[v]", "-map", "[a]",
         "-t", f"{total:.3f}", "-r", str(FPS),
         "-c:v", "libx264", "-preset", "slow", "-crf", "19", "-profile:v", "high", "-pix_fmt", "yuv420p",
         "-c:a", "aac", "-b:a", "192k", "-movflags", "+faststart", dst])
    shutil.rmtree(fdir, ignore_errors=True)
    print(f"   {dst}  ({total:.1f}s)")


for p in PROIECTE:
    if not ONLY or p["slug"] in ONLY:
        video(p)
