#!/usr/bin/env python3
"""Video demonstrativ cu mascota BILZI: cele 4 variante se prezintă pe rând (1080x1920)."""
import hashlib, json, os, shutil, subprocess, sys

HERE = os.path.dirname(os.path.abspath(__file__))
S = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import anim  # noqa: E402

WORK = os.path.join(S, "work_demo")
OUT = sys.argv[1]
os.makedirs(WORK, exist_ok=True)
PIPER = os.path.join(S, "tts", "bin", "piper")
VOICE = os.path.join(S, "voice", "ro.onnx")
LOGO = os.path.join(S, "logo", "badge.png")
MUSIC = os.path.join(S, "music", "Life_of_Riley.mp3")
BOLD = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
REG = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
FPS = anim.FPS
W, H = 1080, 1920

SEGMENTE = [
    dict(variant=None, nume="Nea Bilzi", sub="mustăcios, cu ciocan",
         text="Salut! Eu sunt <b>Nea Bilzi</b>. Pun acoperișuri ca la carte!",
         say="Salut! Eu sunt Nea Bilzi. Pun acoperișuri ca la carte!",
         gest=[(0, "wave"), (1.9, "explain"), (-0.8, "thumb")]),
    dict(variant="tanar", nume="Meșterul Bilzi", sub="tânăr, cu ruletă",
         text="Eu sunt <b>Meșterul Bilzi</b>. Măsor de două ori, montez o dată. <b>Și bine!</b>",
         say="Eu sunt Meșterul Bilzi. Măsor de două ori, montez o dată. Și bine!",
         gest=[(0, "wave"), (1.7, "explain"), (-0.9, "thumb")]),
    dict(variant="batran", nume="Maistrul Bilzi", sub="cu experiență, barbă albă",
         text="Eu sunt <b>Maistrul Bilzi</b>. Am pus atâtea acoperișuri, că mă salută <b>și porumbeii!</b>",
         say="Eu sunt Maistrul Bilzi. Am pus atâtea acoperișuri, că mă salută și porumbeii!",
         gest=[(0, "wave"), (1.8, "explain"), (-1.3, "point")]),
    dict(variant="mesterita", nume="Meșterița Bilzi", sub="meșteriță, cu ruletă",
         text="Eu sunt <b>Meșterița Bilzi</b>. Da, eu urc pe acoperiș. <b>Colegii țin scara!</b>",
         say="Eu sunt Meșterița Bilzi. Da, eu urc pe acoperiș. Colegii țin scara!",
         gest=[(0, "wave"), (1.9, "explain"), (-0.9, "thumb")], feminin=True),
]
LEAD, TAIL = 0.75, 1.7


def run(cmd, **kw):
    r = subprocess.run(cmd, capture_output=True, text=True, **kw)
    if r.returncode:
        print(" ".join(map(str, cmd))[:2000]); print(r.stderr[-3000:]); sys.exit(1)
    return r


def dur(p):
    return float(subprocess.check_output(["ffprobe", "-v", "error", "-show_entries", "format=duration",
                                          "-of", "csv=p=0", p], text=True))


def tts(say, feminin=False):
    h = hashlib.md5((say + str(feminin)).encode()).hexdigest()[:12]
    raw, out = os.path.join(WORK, f"v_{h}_raw.wav"), os.path.join(WORK, f"v_{h}.wav")
    if not os.path.exists(out):
        run([PIPER, "-m", VOICE, "-f", raw, "--length-scale", "0.92", "--sentence-silence", "0.12"], input=say)
        pitch = "rubberband=pitch=1.3:formant=shifted," if feminin else ""
        run(["ffmpeg", "-y", "-v", "error", "-i", raw, "-af",
             pitch + "silenceremove=start_periods=1:start_threshold=-45dB,areverse,"
             "silenceremove=start_periods=1:start_threshold=-45dB,areverse,"
             "highpass=f=80,equalizer=f=200:t=q:w=1:g=2,equalizer=f=3500:t=q:w=1:g=3,aresample=44100",
             "-ac", "1", out])
    return out, dur(out)


def ease(x):
    x = max(0.0, min(1.0, x))
    return x * x * (3 - 2 * x)


def main():
    frames, voice_parts, titles = [], [], []
    t0 = 0.0
    for i, sg in enumerate(SEGMENTE):
        wav, d = tts(sg["say"], sg.get("feminin", False))
        seg_total = LEAD + d + TAIL
        # pista de voce a segmentului (cu liniște înainte/după) -> trăsături pentru gură
        swav = os.path.join(WORK, f"seg{i}.wav")
        ms = int(LEAD * 1000)
        run(["ffmpeg", "-y", "-v", "error", "-i", wav, "-af",
             f"adelay={ms},apad,atrim=duration={seg_total:.3f}", "-ac", "1", "-ar", "44100", swav])
        opn, wide = anim.voice_features(swav, seg_total)
        gestures = [((g if g >= 0 else d + g), name) for g, name in sg["gest"]]
        line = dict(start=LEAD, dur=d, text=sg["text"], gestures=gestures,
                    question=sg["say"].strip().endswith("?"))

        def place(t, seg_total=seg_total):
            out = ease((t - (seg_total - 0.45)) / 0.45)          # iese pe jos la final
            return 540, 1990 + out * 1100, 2.0

        def bubble(ln, t, show):
            return dict(html=ln["text"], show=show, x=70, y=700, tail=420, fs=52)

        fr = anim.animate(seg_total, [line], place, opn, wide, seed=11 + i * 7, enter_at=0.12,
                          variant=sg["variant"], bubble_fn=bubble)
        frames += fr
        voice_parts.append(swav)
        titles.append((t0, t0 + seg_total, i, sg))
        t0 += len(fr) / FPS
    total = len(frames) / FPS

    fjson = os.path.join(WORK, "frames.json")
    with open(fjson, "w", encoding="utf-8") as f:
        json.dump(frames, f, ensure_ascii=False)
    fdir = os.path.join(WORK, "ov")
    shutil.rmtree(fdir, ignore_errors=True)
    print(f"randez {len(frames)} cadre…", flush=True)
    run(["node", os.path.join(HERE, "render.js"), fjson, fdir, str(W), str(H)],
        env=dict(os.environ, NODE_PATH="/opt/node22/lib/node_modules"))

    # voce completă
    voice = os.path.join(WORK, "voice.wav")
    lst = os.path.join(WORK, "voice.txt")
    with open(lst, "w") as f:
        f.writelines(f"file '{p}'\n" for p in voice_parts)
    run(["ffmpeg", "-y", "-v", "error", "-f", "concat", "-safe", "0", "-i", lst, "-c", "copy", voice])

    # fundal: crem, cerc turcoaz în spatele mascotei, siglă, titluri pe segmente
    draw = []
    for a, b, i, sg in titles:
        tn = os.path.join(WORK, f"t{i}.txt"); open(tn, "w", encoding="utf-8").write(f"{i+1} / 4  ·  {sg['nume']}")
        ts = os.path.join(WORK, f"s{i}.txt"); open(ts, "w", encoding="utf-8").write(sg["sub"])
        en = f"enable='between(t,{a:.2f},{b:.2f})'"
        draw.append(f"drawtext=fontfile={BOLD}:textfile={tn}:fontsize=64:fontcolor=0x0F534F:x=(w-tw)/2:y=335:{en}")
        draw.append(f"drawtext=fontfile={REG}:textfile={ts}:fontsize=40:fontcolor=0x5A5A5A:x=(w-tw)/2:y=420:{en}")
    bg = (f"color=c=0xF3EFE6:s={W}x{H}:r={FPS}:d={total:.3f},format=rgb24,"
          f"geq=r='if(lt(hypot(X-540,Y-1330),470),18,if(lt(hypot(X-540,Y-1330),490),230,243))':"
          f"g='if(lt(hypot(X-540,Y-1330),470),127,if(lt(hypot(X-540,Y-1330),490),115,239))':"
          f"b='if(lt(hypot(X-540,Y-1330),470),126,if(lt(hypot(X-540,Y-1330),490),26,230))'")
    flt = (f"{bg}[bg0];[bg0]" + ",".join(draw) + "[bg1];"
           f"[1:v]scale=220:-1[lg];[bg1][lg]overlay=36:150[bg2];"
           f"[bg2][0:v]overlay=0:0:shortest=1,format=yuv420p[v];"
           f"[3:a]atrim=duration={total:.3f},asetpts=PTS-STARTPTS,loudnorm=I=-24:TP=-3,volume=0.8,"
           f"aformat=channel_layouts=stereo,aresample=44100[mus];"
           f"[2:a]loudnorm=I=-15:TP=-2,aformat=channel_layouts=stereo,aresample=44100,asplit=2[vo][sc];"
           f"[mus][sc]sidechaincompress=threshold=0.03:ratio=6:attack=30:release=450[duck];"
           f"[duck][vo]amix=inputs=2:normalize=0,loudnorm=I=-14:TP=-1.5,"
           f"afade=t=out:st={total-1.2:.3f}:d=1.2[a]")
    dst = os.path.join(OUT, "Mascota_BILZI_demo_4_variante.mp4")
    run(["ffmpeg", "-y", "-v", "error", "-framerate", str(FPS), "-i", os.path.join(fdir, "f_%05d.png"),
         "-loop", "1", "-i", LOGO, "-i", voice, "-i", MUSIC,
         "-filter_complex", flt, "-map", "[v]", "-map", "[a]", "-t", f"{total:.3f}", "-r", str(FPS),
         "-c:v", "libx264", "-preset", "slow", "-crf", "19", "-pix_fmt", "yuv420p",
         "-c:a", "aac", "-b:a", "192k", "-movflags", "+faststart", dst])
    shutil.rmtree(fdir, ignore_errors=True)
    print(dst, f"({total:.1f}s)")


if __name__ == "__main__":
    main()
