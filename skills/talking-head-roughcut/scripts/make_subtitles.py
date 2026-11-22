#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""按《字幕规范》给粗剪成片生成/压制字幕。

用法（在粗剪输出目录或任意目录）：
    python make_subtitles.py --plan plan.json --out subtitles.ass
    python make_subtitles.py --plan plan.json --burn 待剪视频_粗剪.mp4 --out-video 待剪视频_字幕版.mp4

规范：
- 不显示句读标点；顿号换成一个半角空格；保留数字/英文里的 . - % /
- 按句读和语义断行，一行约 8–16 字；不拆词、不拆数字单位
- 每条字幕延续到下一条开始，最后一条撑到该句结束（不吃尾音）
- 只做字幕，不改音画
"""
import argparse
import json
import os
import re
import subprocess
import sys
import tempfile

PUNCT = "，。？！：；、,.?!:;\"'“”‘’（）()【】《》…—·"

# 句子间允许的最短间隔（秒），用于避免字幕压到下句开头
TAIL_PAD = 0.02
HEAD_PAD = 0.10


def clean(txt: str) -> str:
    txt = txt.replace("、", " ")
    txt = re.sub("[" + re.escape("，。？！：；,?!:;\"'“”‘’（）()【】《》…—·") + "]", "", txt)
    txt = re.sub(r"[ \t]+", " ", txt).strip()
    return txt


def split_phrases(txt: str, limit: int = 20):
    parts = [p for p in re.split(r"(?<=[，。？！：；])", txt) if p.strip()]
    out = []
    for p in parts:
        body = p
        while len(clean(body)) > limit:
            cut = None
            for i in range(8, len(body) - 4):
                if body[i] in "、 " and not (body[i - 1].isascii() and body[i - 1].isalnum()):
                    cut = i
                    break
            if cut is None:
                cut = len(body) // 2
                while cut < len(body) - 2 and body[cut].isascii() and body[cut].isalnum():
                    cut += 1
            out.append(body[: cut + 1])
            body = body[cut + 1:]
        if body.strip():
            out.append(body)
    return out


def ass_time(t: float) -> str:
    t = max(0.0, t)
    h = int(t // 3600)
    m = int((t % 3600) // 60)
    s = t % 60
    return f"{h}:{m:02d}:{s:05.2f}"


def build(plan_path: str, limit: int = 20):
    plan = json.load(open(plan_path, encoding="utf-8-sig"))
    speed = plan.get("speed", 1.0)
    segs, cum, t = plan["segments"], [], 0.0
    for s in segs:
        cum.append(t)
        t += (s["e"] - s["s"]) / speed

    def to_rough(src):
        for s, c in zip(segs, cum):
            d = (s["e"] - s["s"]) / speed
            if c - 0.001 <= src <= c + d + 0.001:
                return c + (src - s["s"]) / speed
        return None

    lines = []
    sentences = plan["sentences"]
    for st in plan["chain"]:
        a, b = to_rough(st["s"]), to_rough(st["e"])
        if a is None or b is None or b <= a:
            continue
        text = sentences[st["i"]] if st["i"] < len(sentences) else ""
        phrases = [clean(p) for p in split_phrases(text, limit)]
        phrases = [p for p in phrases if p]
        if not phrases:
            continue
        weights = [len(p.replace(" ", "")) or 1 for p in phrases]
        total = sum(weights)
        t0, t1 = a + HEAD_PAD, b - TAIL_PAD
        if t1 <= t0:
            continue
        cur = t0
        for i, (p, w) in enumerate(zip(phrases, weights)):
            dur = (t1 - t0) * w / total
            end = cur + dur if i < len(phrases) - 1 else t1
            lines.append((round(cur, 3), round(end, 3), p))
            cur = end
    return lines


ASS_HEAD = """[Script Info]
ScriptType: v4.00+
PlayResX: {w}
PlayResY: {h}
WrapStyle: 0
ScaledBorderAndShadow: yes

[V4+ Styles]
Format: Name, Fontname, Fontsize, PrimaryColour, SecondaryColour, OutlineColour, BackColour, Bold, Italic, Underline, StrikeOut, ScaleX, ScaleY, Spacing, Angle, BorderStyle, Outline, Shadow, Alignment, MarginL, MarginR, MarginV, Encoding
Style: Default,{font},{fs},&H00FFFFFF,&H00FFFFFF,&H00141418,&H64000000,-1,0,0,0,100,100,1,0,1,{ol},2,2,60,60,{mv},1

[Events]
Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text
"""


def write_ass(lines, out, w, h, fs, mv, font, outline):
    with open(out, "w", encoding="utf-8") as f:
        f.write(ASS_HEAD.format(w=w, h=h, fs=fs, mv=mv, font=font, ol=outline))
        for a, b, txt in lines:
            f.write(f"Dialogue: 0,{ass_time(a)},{ass_time(b)},Default,,0,0,0,,{txt}\n")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--plan", required=True)
    ap.add_argument("--out", default="subtitles.ass")
    ap.add_argument("--width", type=int, default=1080)
    ap.add_argument("--height", type=int, default=1920)
    ap.add_argument("--font-size", type=int, default=62)
    ap.add_argument("--margin-v", type=int, default=95)
    ap.add_argument("--font", default="Microsoft YaHei")
    ap.add_argument("--outline", type=float, default=4)
    ap.add_argument("--limit", type=int, default=20)
    ap.add_argument("--burn", default="")
    ap.add_argument("--out-video", default="")
    ap.add_argument("--crf", type=int, default=16)
    args = ap.parse_args()

    lines = build(args.plan, args.limit)
    write_ass(lines, args.out, args.width, args.height, args.font_size, args.margin_v, args.font, args.outline)
    print(f"subtitles {len(lines)} -> {args.out}")

    if args.burn:
        out_video = args.out_video or (os.path.splitext(args.burn)[0] + "_字幕版.mp4")
        work_dir = os.path.dirname(os.path.abspath(args.plan))
        ass_rel = os.path.relpath(os.path.abspath(args.out), work_dir).replace("\\", "/")
        vf = f"ass={ass_rel}"
        cmd = [os.environ.get("KOU_BO_FFMPEG", "ffmpeg"), "-v", "error", "-y", "-i", os.path.abspath(args.burn),
               "-vf", vf, "-c:v", "libx264", "-preset", "veryfast", "-crf", str(args.crf),
               "-c:a", "copy", "-movflags", "+faststart", os.path.abspath(out_video)]
        subprocess.run(cmd, check=True, cwd=work_dir)
        print("burned ->", out_video)


if __name__ == "__main__":
    main()
