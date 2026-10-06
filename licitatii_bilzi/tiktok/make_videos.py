#!/usr/bin/env python3
"""Generează câte un video TikTok (1080x1920) pentru fiecare proiect din portofoliul BILZI."""
import os, subprocess, sys

S = os.path.dirname(os.path.abspath(__file__))
IMG = os.path.join(S, "img")
WORK = os.path.join(S, "work")
OUT = sys.argv[1]
os.makedirs(WORK, exist_ok=True)
os.makedirs(OUT, exist_ok=True)

W, H, FPS = 1080, 1920, 30
FW, FH = 1080, 1350          # fereastra foto (4:5), centrată
FY = 556                     # sub siglă și banda de text
LOGO = os.path.join(S, "logo", "badge.png")   # sigla pe fundal alb, colțuri rotunjite
LOGO_W = 220                 # lățimea siglei din colț
MUSIC = os.path.join(S, "music")
CLIP = 3.4                   # durata unei poze
XF = 0.5                     # durata tranziției
END = 3.0                    # card final
BOLD = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
REG = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
ACCENT = "0xE30613"          # roșu Bilka

PROIECTE = [
    dict(slug="01_Ciocarlia_Bilka_Balcanic_8004",
         hook="ACOPERIȘ NOU\nVĂZUT DIN DRONĂ",
         titlu="Casă în Ciocârlia",
         model="Țiglă metalică Bilka Balcanic 8004 MAT",
         loc="Ciocârlia, Ialomița",
         muzica=("Life_of_Riley.mp3", 0),
         poze=["proiect1-img1.jpg", "proiect1-img2-1.jpg", "proiect1-img4-1.jpg",
               "proiect1-img3-1.jpg.webp", "proiect1-img5-1.jpg.webp",
               "proiect1-img6-1-rotated.jpg.webp"]),
    dict(slug="02_Afumati_Bilka_Clasic_8019",
         hook="ELEGANȚĂ ÎN\nGRI ANTRACIT",
         titlu="Casă în Afumați",
         model="Țiglă metalică Bilka Clasic 8019 Grande Mat",
         loc="Afumați, Ilfov",
         muzica=("Wallpaper.mp3", 0),
         poze=["proiect2-img1.jpg", "proiect2-img2-1.jpg.webp", "proiect2-img3.jpg",
               "proiect2-img3-1.jpg.webp", "proiect2-img4.jpg.webp"]),
    dict(slug="03_Lilieci_Camin_Cultural_Bilka_Clasic_8017",
         hook="CĂMIN CULTURAL\nACOPERIȘ ÎNLOCUIT",
         titlu="Cămin cultural Lilieci",
         model="Țiglă metalică Bilka Clasic 8017 MAT",
         loc="Lilieci, Ialomița",
         muzica=("Inspired.mp3", 0),
         poze=["bilka-materiale-livrare.jpg.webp", "proiect3-img1.jpg.webp",
               "proiect3-img4.jpg", "proiect3-img3-1.jpg.webp", "proiect3-img4-1.jpg.webp"]),
    dict(slug="04_Bilka_Iberic_Negru_Pluvial_Alb",
         hook="NEGRU + ALB\nCONTRAST PERFECT",
         titlu="Casă nouă, de la zero",
         model="Bilka Iberic negru · sistem pluvial alb",
         loc="Montaj complet acoperiș",
         muzica=("Funkorama.mp3", 12),
         poze=["proiect4-img1-1.jpg.webp", "proiect4-img3.jpg.webp",
               "proiect4-img2.jpg.webp", "proiect4-img4.jpg.webp"]),
]


def run(cmd):
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode:
        print(" ".join(cmd)); print(r.stderr[-3000:]); sys.exit(1)


def size(path):
    o = subprocess.check_output(["ffprobe", "-v", "error", "-select_streams", "v:0",
                                 "-show_entries", "stream=width,height", "-of", "csv=p=0", path], text=True)
    w, h = map(int, o.strip().split(",")[:2])
    return w, h


def txt(name, s):
    p = os.path.join(WORK, name)
    with open(p, "w", encoding="utf-8") as f:
        f.write(s)
    return p


def clip(src, dst, idx):
    """O poză -> clip 1080x1920 cu fundal blurat și mișcare lentă."""
    w, h = size(src)
    n = int(CLIP * FPS)
    bg = (f"[0:v]scale={W}:{H}:force_original_aspect_ratio=increase,crop={W}:{H},"
          f"boxblur=40:3,eq=brightness=-0.18:saturation=0.8[bg]")
    if w / h > FW / FH * 1.15:
        # peisaj: panoramare stânga<->dreapta pe o fereastră 4:5
        a, b = (0.1, 0.9) if idx % 2 == 0 else (0.9, 0.1)
        fg = (f"[0:v]scale=-2:{FH*2}:flags=lanczos,"
              f"crop={FW*2}:{FH*2}:x='(iw-ow)*({a}+({b}-{a})*t/{CLIP})':y=0,"
              f"scale={FW}:{FH}:flags=lanczos[fg]")
    else:
        # portret / aproape 4:5: zoom lent spre centru
        fg = (f"[0:v]scale={FW*3}:{FH*3}:force_original_aspect_ratio=increase:flags=lanczos,"
              f"crop={FW*3}:{FH*3},"
              f"zoompan=z='1+0.10*in/{n}':x='iw/2-iw/zoom/2':y='ih/2-ih/zoom/2':d=1:s={FW}x{FH}:fps={FPS}[fg]")
    flt = f"{bg};{fg};[bg][fg]overlay=0:{FY}:shortest=1,format=yuv420p[v]"
    run(["ffmpeg", "-y", "-v", "error", "-loop", "1", "-framerate", str(FPS), "-t", str(CLIP), "-i", src,
         "-filter_complex", flt, "-map", "[v]", "-t", str(CLIP), "-r", str(FPS),
         "-c:v", "libx264", "-preset", "medium", "-crf", "16", dst])


def endcard(dst, p):
    l1 = txt("e1.txt", "BILZI STEEL PROFILE")
    l2 = txt("e2.txt", "Montaj acoperișuri · Țiglă metalică")
    l3 = txt("e3.txt", "Cere o ofertă!")
    l4 = txt("e4.txt", "0720 244 244")
    l5 = txt("e5.txt", "www.acoperis-mag.ro")
    last = os.path.join(IMG, p["poze"][0])
    flt = (f"[0:v]scale={W}:{H}:force_original_aspect_ratio=increase,crop={W}:{H},"
           f"boxblur=30:3,eq=brightness=-0.35:saturation=0.6,"
           f"drawbox=x=90:y=1030:w={W-180}:h=150:color={ACCENT}@1:t=fill,"
           f"drawtext=fontfile={BOLD}:textfile={l1}:fontsize=78:fontcolor=white:x=(w-tw)/2:y=700,"
           f"drawtext=fontfile={REG}:textfile={l2}:fontsize=44:fontcolor=white@0.9:x=(w-tw)/2:y=810,"
           f"drawtext=fontfile={BOLD}:textfile={l3}:fontsize=56:fontcolor=white:x=(w-tw)/2:y=950,"
           f"drawtext=fontfile={BOLD}:textfile={l4}:fontsize=84:fontcolor=white:x=(w-tw)/2:y=1060,"
           f"drawtext=fontfile={REG}:textfile={l5}:fontsize=48:fontcolor=white:x=(w-tw)/2:y=1220,"
           f"fade=t=in:st=0:d=0.4,format=yuv420p[v]")
    run(["ffmpeg", "-y", "-v", "error", "-loop", "1", "-framerate", str(FPS), "-t", str(END), "-i", last,
         "-filter_complex", flt, "-map", "[v]", "-t", str(END), "-r", str(FPS),
         "-c:v", "libx264", "-preset", "medium", "-crf", "16", dst])


def video(p):
    print("->", p["slug"])
    clips = []
    for i, f in enumerate(p["poze"]):
        c = os.path.join(WORK, f"{p['slug']}_{i}.mp4")
        clip(os.path.join(IMG, f), c, i)
        clips.append((c, CLIP))
    e = os.path.join(WORK, f"{p['slug']}_end.mp4")
    endcard(e, p)
    clips.append((e, END))

    # lanț de tranziții xfade
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
    total = t + clips[-1][1]
    body_end = total - END  # momentul când începe cardul final

    hook = txt(f"{p['slug']}_hook.txt", p["hook"])
    titlu = txt(f"{p['slug']}_t.txt", p["titlu"])
    brand = txt(f"{p['slug']}_b.txt", "BILZI STEEL PROFILE")
    model = txt(f"{p['slug']}_m.txt", p["model"])
    loc = txt(f"{p['slug']}_l.txt", p["loc"])
    on = f"enable='lt(t,{body_end:.2f})'"
    hook_on = "enable='between(t,0.2,2.9)'"
    alpha_hook = "alpha='if(lt(t,0.5),(t-0.2)/0.3,if(gt(t,2.5),(2.9-t)/0.4,1))'"
    top_y = 263
    bot_y = 415
    over = (
        # bandă sus: brand + titlu proiect
        f"drawtext=fontfile={BOLD}:textfile={titlu}:fontsize=66:fontcolor=white:"
        f"shadowcolor=black@0.6:shadowx=3:shadowy=3:x=(w-tw)/2:y={top_y+55}:{on},"
        # bandă jos: model țiglă (pe fundal roșu) + localitate
        f"drawtext=fontfile={BOLD}:textfile={model}:fontsize=38:fontcolor=white:box=1:boxcolor={ACCENT}@0.95:"
        f"boxborderw=18:x=(w-tw)/2:y={bot_y}:{on},"
        f"drawtext=fontfile={REG}:textfile={loc}:fontsize=40:fontcolor=white:"
        f"shadowcolor=black@0.6:shadowx=2:shadowy=2:x=(w-tw)/2:y={bot_y+80}:{on},"
        # hook mare în primele secunde, peste poză
        f"drawtext=fontfile={BOLD}:textfile={hook}:fontsize=78:line_spacing=16:fontcolor=white:"
        f"box=1:boxcolor=black@0.45:boxborderw=30:text_align=center:x=(w-tw)/2:y=(h-th)/2:{alpha_hook}:{hook_on}"
    )
    n = len(clips)
    parts.append(f"{cur}{over}[t]")
    # sigla permanentă în colțul stânga-sus
    parts.append(f"[{n}:v]scale={LOGO_W}:-1:flags=lanczos[lg]")
    parts.append(f"[t][lg]overlay=36:150,format=yuv420p[v]")
    # muzică: tăiată la durata videoului, normalizată, fade in/out
    mf, start = p["muzica"]
    parts.append(f"[{n+1}:a]atrim=start={start}:duration={total:.3f},asetpts=PTS-STARTPTS,"
                 f"loudnorm=I=-14:TP=-1.5:LRA=11,afade=t=in:d=0.4,"
                 f"afade=t=out:st={total-1.8:.3f}:d=1.8,aresample=44100[a]")
    dst = os.path.join(OUT, f"BILZI_TikTok_{p['slug']}.mp4")
    run(["ffmpeg", "-y", "-v", "error", *inputs,
         "-loop", "1", "-i", LOGO, "-i", os.path.join(MUSIC, mf),
         "-filter_complex", ";".join(parts), "-map", "[v]", "-map", "[a]",
         "-t", f"{total:.3f}", "-r", str(FPS),
         "-c:v", "libx264", "-preset", "slow", "-crf", "19", "-profile:v", "high", "-pix_fmt", "yuv420p",
         "-c:a", "aac", "-b:a", "192k", "-movflags", "+faststart", dst])
    print(f"   {dst}  ({total:.1f}s)")


for p in PROIECTE:
    video(p)
