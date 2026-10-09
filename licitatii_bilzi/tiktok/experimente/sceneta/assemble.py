#!/usr/bin/env python3
"""Randează toate cadrele scenetei (3 procese în paralel) și le montează cu voce, muzică și efecte."""
import json, os, shutil, subprocess, sys

HERE = os.path.dirname(os.path.abspath(__file__))
S = os.path.dirname(HERE)
WORK = os.path.join(S, "work_skit")
OUT = sys.argv[1]
FPS = 30
GROUND = 1792

frames = json.load(open(os.path.join(WORK, "frames.json")))
meta = json.load(open(os.path.join(WORK, "meta.json")))
total = meta["total"]
fdir = os.path.join(WORK, "frames")
shutil.rmtree(fdir, ignore_errors=True)
env = dict(os.environ, NODE_PATH="/opt/node22/lib/node_modules")
procs = [subprocess.Popen(["node", os.path.join(HERE, "render.js"), os.path.join(WORK, "frames.json"), fdir,
                           str(k), "3"], env=env) for k in range(3)]
for p in procs:
    assert p.wait() == 0

# efecte sonore sintetizate
def synth(name, expr, d):
    path = os.path.join(WORK, f"sfx_{name}.wav")
    subprocess.run(["ffmpeg", "-y", "-v", "error", "-f", "lavfi", "-i",
                    f"aevalsrc='{expr}':s=44100:d={d}", "-ac", "1", path], check=True)
    return path

toc = synth("toc", "0.9*sin(2*PI*240*t)*exp(-t*38)+0.5*(random(0)*2-1)*exp(-t*70)", 0.18)
thud = synth("thud", "0.9*sin(2*PI*85*t)*exp(-t*16)+0.35*(random(0)*2-1)*exp(-t*45)", 0.35)
whoosh = synth("whoosh", "0.5*(random(0)*2-1)*sin(PI*t/0.45)", 0.45)
events = [(toc, t) for kind, t in meta["sfx"] if kind == "toc"]
# momentul exact în care ciocanul atinge iarba
landed = next((i for i, f in enumerate(frames) if f["j"].get("hammer") and f["j"]["hammer"]["p"][1] >= GROUND - 9
               and i > 30), None)
if landed:
    events.append((thud, landed / FPS))
slip = next(s for k, s, d in meta["lines"] if k == "aoleu")
events.append((whoosh, slip + 0.05))

voice = os.path.join(WORK, "voice.wav")
inputs = ["-framerate", str(FPS), "-i", os.path.join(fdir, "f_%05d.jpg"),
          "-i", voice, "-i", os.path.join(S, "music", "Life_of_Riley.mp3")]
parts = []
for i, (path, t) in enumerate(events):
    inputs += ["-i", path]
    ms = int(t * 1000)
    parts.append(f"[{3+i}:a]adelay={ms}|{ms},volume=0.9,aformat=channel_layouts=stereo,aresample=44100[e{i}]")
parts.append("".join(f"[e{i}]" for i in range(len(events))) +
             f"amix=inputs={len(events)}:normalize=0,apad,atrim=duration={total:.3f}[fx]")
parts.append(f"[2:a]atrim=duration={total:.3f},asetpts=PTS-STARTPTS,loudnorm=I=-22:TP=-3,volume=0.85,"
             f"aformat=channel_layouts=stereo,aresample=44100[mus]")
parts.append("[1:a]loudnorm=I=-15:TP=-2,aformat=channel_layouts=stereo,aresample=44100,asplit=2[vo][sc]")
parts.append("[mus][sc]sidechaincompress=threshold=0.03:ratio=6:attack=30:release=450[duck]")
parts.append(f"[duck][vo][fx]amix=inputs=3:normalize=0,loudnorm=I=-14:TP=-1.5,"
             f"afade=t=in:d=0.3,afade=t=out:st={total-1.5:.3f}:d=1.5[a]")
dst = os.path.join(OUT, "Nea_Bilzi_pe_acoperis.mp4")
subprocess.run(["ffmpeg", "-y", "-v", "error", *inputs, "-filter_complex", ";".join(parts),
                "-map", "0:v", "-map", "[a]", "-t", f"{total:.3f}", "-r", str(FPS),
                "-c:v", "libx264", "-preset", "slow", "-crf", "19", "-pix_fmt", "yuv420p",
                "-c:a", "aac", "-b:a", "192k", "-movflags", "+faststart", dst], check=True)
shutil.rmtree(fdir, ignore_errors=True)
print(dst, f"{total:.1f}s")
