#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""口播视频粗剪 (talking-head rough cut).

匹配原素材与口播稿 -> 对齐校准 -> 去掉重复录制/卡壳/长气口 -> 切点换景别 -> 按原分辨率、原码率输出。

用法示例:
    python roughcut.py --video "待剪视频.MOV" --script "口播稿.docx" --dry-run
    python roughcut.py --video "待剪视频.MOV" --script "口播稿.docx"
    python roughcut.py --from-plan "xxx_粗剪/plan.json"

依赖: ffmpeg/ffprobe、faster-whisper + 本地模型、numpy、python-docx(不需要,docx 用 zip 解析)。
路径通过 scripts/environment.json 或环境变量 KOU_BO_FFMPEG / KOU_BO_FFPROBE / KOU_BO_MODEL 提供。
"""
from __future__ import annotations

import argparse
import bisect
import csv
import json
import os
import re
import shutil
import subprocess
import sys
import wave
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

HERE = Path(__file__).resolve().parent
ENV_FILE = HERE / "environment.json"
CJK_RE = re.compile(r"[\u3400-\u9fffA-Za-z0-9]")
FILLERS = set("嗯呃额")
# 句尾语气词 + 模糊限定词：切到它们等于改语义（「1000-2000左右」变确切数字），必须留全
TAIL_TOKENS = {
    "呢", "啊", "吧", "啦", "嘛", "呀", "了", "的",
    "左右", "上下", "大概", "差不多", "之类", "这样", "上下吧",
}
LAUGH = re.compile(r"^[哈呵嘿嘻]+$")


# --------------------------------------------------------------------------- #
# 环境
# --------------------------------------------------------------------------- #
def load_env() -> dict:
    env = {}
    if ENV_FILE.exists():
        env = json.loads(ENV_FILE.read_text("utf-8"))
    mapping = {
        "ffmpeg": "KOU_BO_FFMPEG",
        "ffprobe": "KOU_BO_FFPROBE",
        "model": "KOU_BO_MODEL",
    }
    for key, var in mapping.items():
        if os.environ.get(var):
            env[key] = os.environ[var]
    env.setdefault("ffmpeg", shutil.which("ffmpeg") or "")
    if not env.get("ffprobe"):
        ff = env.get("ffmpeg") or ""
        cand = Path(ff).with_name("ffprobe.exe") if ff else None
        env["ffprobe"] = str(cand) if cand and cand.exists() else (shutil.which("ffprobe") or "")
    for key in ("ffmpeg", "ffprobe"):
        if not env.get(key) or not Path(env[key]).exists():
            raise SystemExit(
                f"找不到 {key}。请在 {ENV_FILE} 里填写，或设置环境变量 {mapping[key]}。"
            )
    if not env.get("model") or not Path(env["model"]).exists():
        raise SystemExit(f"找不到语音模型目录。请在 {ENV_FILE} 里填写 model，或设置 KOU_BO_MODEL。")
    return env


# --------------------------------------------------------------------------- #
# 媒体探测
# --------------------------------------------------------------------------- #
@dataclass
class Media:
    path: str
    width: int
    height: int
    fps: float
    duration: float
    v_bitrate: int
    a_bitrate: int
    sample_rate: int
    channels: int
    pix_fmt: str


def probe_media(ffprobe: str, path: str) -> Media:
    raw = subprocess.run(
        [ffprobe, "-v", "error", "-print_format", "json", "-show_format", "-show_streams", path],
        capture_output=True,
    )
    if raw.returncode != 0:
        raise SystemExit("ffprobe 读取失败: " + raw.stderr.decode("utf-8", "replace")[-800:])
    data = json.loads(raw.stdout.decode("utf-8", "replace"))
    vs = next(s for s in data["streams"] if s["codec_type"] == "video")
    aus = next((s for s in data["streams"] if s["codec_type"] == "audio"), None)

    def frac(s: str) -> float:
        try:
            a, b = s.split("/")
            return float(a) / float(b) if float(b) else 0.0
        except Exception:
            return 0.0

    fps = frac(vs.get("r_frame_rate", "0/0")) or frac(vs.get("avg_frame_rate", "0/0")) or 30.0
    duration = float(data["format"].get("duration") or vs.get("duration") or 0)
    vbr = int(float(vs.get("bit_rate") or 0)) or int(float(data["format"].get("bit_rate") or 0))
    abr = int(float(aus.get("bit_rate") or 0)) if aus else 128000
    return Media(
        path=path,
        width=int(vs["width"]),
        height=int(vs["height"]),
        fps=round(fps, 4),
        duration=duration,
        v_bitrate=vbr,
        a_bitrate=abr or 128000,
        sample_rate=int(aus.get("sample_rate") or 44100) if aus else 44100,
        channels=int(aus.get("channels") or 2) if aus else 2,
        pix_fmt=vs.get("pix_fmt") or "yuv420p",
    )


def extract_wav(ffmpeg: str, video: str, wav_path: str) -> None:
    cmd = [ffmpeg, "-hide_banner", "-loglevel", "error", "-y", "-i", video,
           "-vn", "-ac", "1", "-ar", "16000", "-c:a", "pcm_s16le", wav_path]
    r = subprocess.run(cmd, capture_output=True)
    if r.returncode != 0:
        raise SystemExit("抽取音轨失败: " + r.stderr.decode("utf-8", "replace")[-800:])


def read_wav_mono16(path: str):
    import numpy as np
    with wave.open(path, "rb") as w:
        frames = w.readframes(w.getnframes())
    return np.frombuffer(frames, dtype=np.int16).astype("float32") / 32768.0


# --------------------------------------------------------------------------- #
# 口播稿
# --------------------------------------------------------------------------- #
W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def docx_paragraphs(path: str) -> list[str]:
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    out = []
    for p in root.iter(W_NS + "p"):
        txt = "".join(t.text or "" for t in p.iter(W_NS + "t")).strip()
        if txt:
            out.append(txt)
    return out


def read_text_file(path: str) -> list[str]:
    p = Path(path)
    if p.suffix.lower() == ".docx":
        return docx_paragraphs(path)
    raw = p.read_text("utf-8", errors="replace")
    return [x.strip() for x in raw.splitlines() if x.strip()]


def norm(text: str) -> str:
    t = "".join(CJK_RE.findall(text)).lower()
    return t.translate(str.maketrans("0123456789", "零一二三四五六七八九"))


def pick_script_paragraphs(paras: list[str], mode: str) -> list[str]:
    """默认取正文段：长度 >= 最长段 70% 的段落（跳过标签/话题段）。"""
    cand = [p for p in paras if len(norm(p)) >= 8 and norm(p).count("#") * 2 < len(norm(p))]
    if not cand:
        raise SystemExit("口播稿里没有可用正文。")
    if mode == "all":
        return cand
    if mode.isdigit():
        idx = int(mode)
        if not 1 <= idx <= len(cand):
            raise SystemExit(f"--paragraph 取值 1..{len(cand)}")
        return [cand[idx - 1]]
    longest = max(len(norm(p)) for p in cand)
    return [p for p in cand if len(norm(p)) >= 0.7 * longest]


def segment_script(text: str) -> list[str]:
    """按句子/分句切分台词，保证每段 20~45 字，便于逐句对齐和定位重录。"""
    units, cur = [], ""
    for ch in text:
        if ch in "\r\n":
            if norm(cur):
                units.append(cur.strip())
            cur = ""
            continue
        cur += ch
        if ch in "。！？!?；;":
            if norm(cur):
                units.append(cur.strip())
            cur = ""
        elif ch in "，,、：:" and len(norm(cur)) >= 16:
            units.append(cur.strip())
            cur = ""
    if norm(cur):
        units.append(cur.strip())
    units = [u for u in units if norm(u)]
    # 太短的尾巴（如「是不花什么米的。」）并回上一句，避免碎句对不上
    merged: list[str] = []
    for u in units:
        if merged and len(norm(u)) < 8 and len(norm(merged[-1])) + len(norm(u)) <= 30:
            merged[-1] = merged[-1].rstrip() + u
        else:
            merged.append(u)
    return merged


# --------------------------------------------------------------------------- #
# 语音识别
# --------------------------------------------------------------------------- #
def transcribe(model, audio, prompt: str, condition: bool = False) -> tuple[list[dict], list[dict]]:
    segments, _info = model.transcribe(
        audio,
        language="zh",
        beam_size=5,
        word_timestamps=True,
        vad_filter=False,
        condition_on_previous_text=condition,
        initial_prompt=prompt,
    )
    words, segs = [], []
    for s in segments:
        ws = []
        for w in (s.words or []):
            wd = {"w": w.word, "s": float(w.start), "e": float(w.end), "p": float(w.probability)}
            ws.append(wd)
            words.append(wd)
        segs.append({"s": float(s.start), "e": float(s.end), "text": s.text, "words": ws})
    return segs, words


def build_chars(words: list[dict]) -> list[dict]:
    chars = []
    for w in words:
        cs = [c.lower() for c in w["w"] if CJK_RE.match(c)]
        n = len(cs)
        if n == 0:
            continue
        d = w["e"] - w["s"]
        for k, c in enumerate(cs):
            chars.append({"c": c, "s": w["s"] + d * k / n, "e": w["s"] + d * (k + 1) / n})
    return chars


def text_between(chars: list[dict], s: float, e: float) -> str:
    return "".join(c["c"] for c in chars if c["s"] >= s - 0.03 and c["e"] <= e + 0.03)


def align_pairs(a: str, b: str) -> list[tuple[int, int]]:
    """Smith-Waterman 局部对齐：在 b 里找与 a 最像的一段，匹配 +2 / 错配 -1 / gap -2。"""
    n, m = len(a), len(b)
    if n == 0 or m == 0:
        return []
    score = [[0] * (m + 1) for _ in range(n + 1)]
    best, bi, bj = 0, 0, 0
    for i in range(1, n + 1):
        ai = a[i - 1]
        row, prev = score[i], score[i - 1]
        for j in range(1, m + 1):
            s = max(0, prev[j - 1] + (3 if ai == b[j - 1] else -1), prev[j] - 2, row[j - 1] - 2)
            row[j] = s
            if s > best:
                best, bi, bj = s, i, j
    i, j, pairs = bi, bj, []
    while i > 0 and j > 0 and score[i][j] > 0:
        if score[i][j] == score[i - 1][j - 1] + (3 if a[i - 1] == b[j - 1] else -1):
            if a[i - 1] == b[j - 1]:
                pairs.append((i - 1, j - 1))
            i -= 1
            j -= 1
        elif score[i][j] == score[i - 1][j] - 2:
            i -= 1
        else:
            j -= 1
    pairs.reverse()
    return pairs


def find_candidates(sent_norm: str, chars: list[dict], min_sim: float = 0.5) -> list[dict]:
    """在整段识别文本里找这句台词的所有出现位置（含多次重录）。"""
    L = len(sent_norm)
    if L < 3 or not chars:
        return []
    n = 1 if L < 8 else (3 if L >= 9 else 2)
    grams = [sent_norm[i:i + n] for i in range(L - n + 1)]
    if not grams:
        return []
    index: dict[str, list[int]] = {}
    for i in range(len(chars) - n + 1):
        g = "".join(chars[i + k]["c"] for k in range(n))
        index.setdefault(g, []).append(i)
    M = len(chars)
    diff = [0] * (M + 2)
    for gi, g in enumerate(grams):
        for p in index.get(g, ()):
            lo = max(0, p + n - L)
            hi = min(M - 1, p)
            if lo <= hi:
                diff[lo] += 1
                diff[hi + 1] -= 1
    counts, c = [0] * M, 0
    for i in range(M):
        c += diff[i]
        counts[i] = c
    peak = max(counts) if counts else 0
    if peak < max(2, int(0.45 * len(grams))):
        return []
    thr = max(int(0.6 * peak), int(0.45 * len(grams)))
    positions = [i for i, v in enumerate(counts) if v >= thr]
    groups, cur = [], [positions[0]]
    for p in positions[1:]:
        if p - cur[-1] <= max(2, L // 3):
            cur.append(p)
        else:
            groups.append(cur)
            cur = [p]
    groups.append(cur)

    out = []
    for g in groups:
        gmin, gmax = min(g), max(g)
        lo = max(0, gmin - L)
        hi = min(M, gmax + L + 1)
        sub = chars[lo:hi]
        pairs = align_pairs(sent_norm, "".join(x["c"] for x in sub))
        if not pairs:
            continue
        js = [j for _, j in pairs]
        # 匹配段必须落在候选群附近，避免对到相邻的相似句
        if max(js) < (gmin - lo) or min(js) > (gmax - lo):
            continue
        cov = len(pairs) / L
        span_len = max(js) - min(js) + 1
        sim = 2.0 * len(pairs) / (L + span_len)
        if sim < min_sim or cov < 0.45:
            continue
        out.append({
            "s": max(0.0, sub[min(js)]["s"]),
            "e": sub[max(js)]["e"],
            "cov": round(cov, 3),
            "sim": round(sim, 3),
        })
    out.sort(key=lambda x: (x["s"], x["e"]))
    merged = []
    for c in out:
        if merged and c["s"] < merged[-1]["e"] - 0.2 and c["e"] > merged[-1]["s"] + 0.2:
            if c["cov"] > merged[-1]["cov"]:
                merged[-1] = c
        else:
            merged.append(c)
    return merged


def choose_chain(cands: list[list[dict]]) -> list[dict]:
    """在所有候选里选一条时间递增、覆盖台词最多、同覆盖率下尽量靠后的链（=只保留最后一次录制）。"""
    states: list[dict] = []
    for i, occs in enumerate(cands):
        for occ in occs:
            best_prev, best_key = None, None
            for st in states:
                if st["i"] < i and st["e"] <= occ["s"] + 0.06:
                    key = (st["count"], st["tsum"], st["qual"])
                    if best_key is None or key > best_key:
                        best_key, best_prev = key, st
            if best_prev is None:
                count, tsum, qual = 1, occ["s"], occ["sim"]
            else:
                count = best_prev["count"] + 1
                tsum = best_prev["tsum"] + occ["s"]
                qual = best_prev["qual"] + occ["sim"]
            st = {"i": i, "s": occ["s"], "e": occ["e"], "cov": occ["cov"], "sim": occ["sim"],
                  "count": count, "tsum": tsum, "qual": qual, "prev": best_prev}
            states.append(st)
    if not states:
        return []
    best = max(states, key=lambda st: (st["count"], st["tsum"], st["qual"]))
    chain = []
    while best is not None:
        chain.append(best)
        best = best["prev"]
    chain.reverse()
    return chain


# --------------------------------------------------------------------------- #
# 剪辑方案
# --------------------------------------------------------------------------- #
def _frame_rms(audio, sr: int, win: float = 0.02):
    """20ms 帧的 RMS 包络。"""
    import numpy as np
    hop = max(1, int(win * sr))
    n = len(audio) // hop
    frames = audio[:n * hop].reshape(n, hop)
    return np.sqrt((frames ** 2).mean(axis=1)), win


def _longest_silence(rms, win: float, t0: float, t1: float, thr: float):
    """在 [t0,t1] 里找最长的一段低于阈值的静音，返回 (start, end)。"""
    i0, i1 = max(0, int(t0 / win)), min(len(rms), int(t1 / win))
    best, cur = None, None
    for i in range(i0, i1):
        if rms[i] < thr:
            if cur is None:
                cur = i
        elif cur is not None:
            if best is None or (i - cur) > (best[1] - best[0]):
                best = (cur, i)
            cur = None
    if cur is not None and (best is None or (i1 - cur) > (best[1] - best[0])):
        best = (cur, i1)
    if best is None:
        return None
    return (best[0] * win, best[1] * win)


def speech_runs(audio, sr: int, win: float = 0.02) -> list[tuple[float, float]]:
    """能量 VAD：返回语音区间（合并 <80ms 小缝，丢弃 <50ms 碎片）。"""
    import numpy as np
    rms, win = _frame_rms(audio, sr, win)
    speech = rms[rms > 0.02]
    thr = max(0.012, float(np.median(speech)) * 0.2) if len(speech) else 0.02
    runs, cur = [], None
    for i, v in enumerate(rms >= thr):
        if v and cur is None:
            cur = i
        elif not v and cur is not None:
            runs.append([cur, i])
            cur = None
    if cur is not None:
        runs.append([cur, len(rms)])
    merged = []
    for r in runs:
        if merged and (r[0] - merged[-1][1]) * win < 0.12:
            merged[-1][1] = r[1]
        else:
            merged.append(r)
    return [(round(s * win, 3), round(e * win, 3)) for s, e in merged if (e - s) * win >= 0.05]


def _tail_word(e: float, words):
    """句子最后一个词；只有它属于语气词/模糊限定词时才返回，否则 None。"""
    cands = [w for w in words if w["e"] <= e + 0.06]
    if not cands:
        return None
    w = max(cands, key=lambda x: x["e"])
    return w if w["w"].strip(" ，。！？、,.!?") in TAIL_TOKENS else None


def extend_tail(e: float, runs, words, pad: float = 0.03, max_ext: float = 0.5) -> float:
    """句尾语气词/「左右」这类模糊限定词必须完整保留：贴着能量包络再往后延一点。

    Whisper 的词尾普遍标得早（「左右」标 204.27s，实际说到 204.5s）。按词边界切会把
    「左右」切掉，数字就变成确切的「1000-2000」，语义变了——用户明确要求必须保留。
    """
    if _tail_word(e, words) is None:
        return e
    for r in runs:
        if r[0] - 0.02 <= e <= r[1] + 0.02:
            return min(r[1] + pad, e + max_ext)
    # 词尾落在两段语音之间的短缝里（「左右」常被切成「左」「右」两段）：贴着后一段走
    after = [r for r in runs if r[0] > e and r[0] - e <= 0.3]
    if after:
        return min(after[0][1] + pad, e + max_ext)
    return e


def snap_span(s: float, e: float, runs, pad: float = 0.03,
              max_back: float = 0.30, max_fwd: float = 0.45) -> tuple[float, float]:
    """把句子首尾贴到语音段边缘：起点取 s 所在段，终点取 e 所在段，保证尾音完整。

    首尾都限幅：起点最多回退 max_back，终点最多后延 max_fwd。
    不限幅就会把紧贴着的重复句/口误整段"吸"进这一句——实测事故：前一句把后面重复的
    「那海南这边的园区地址费用呢」吞掉，用户听到重复，而且中间 3 秒静音也没被剪掉。
    """
    if not runs:
        return s, e
    inside = [r for r in runs if r[0] - 0.02 <= s <= r[1] + 0.02]
    if inside:
        ns = max(0.0, max(inside[0][0] - pad, s - max_back))
    else:
        after = [r for r in runs if r[0] >= s]
        ns = after[0][0] - pad if after and after[0][0] - s <= 0.5 else s

    inside_e = [r for r in runs if r[0] - 0.02 <= e <= r[1] + 0.02]
    if inside_e:
        ne = inside_e[-1][1] + pad
    else:
        before = [r for r in runs if r[1] < e]
        after = [r for r in runs if r[0] > e]
        cands = []
        if before:
            cands.append((e - before[-1][1], before[-1][1] + pad))
        if after:
            cands.append((after[0][1] - e, after[0][1] + pad))
        cands = [c for c in cands if c[0] <= 0.6]
        if not cands:
            ne = e
        else:
            fwd = [c for c in cands if c[0] > 0]
            ne = min(fwd)[1] if fwd and min(fwd)[0] <= 0.25 else min(cands)[1]
    return ns, ne


def trim_stammer_at(s: float, e: float, words, rms, win: float, thr: float,
                    trims: list[dict]) -> float:
    """句首说了半句停住再重来：被拉长的短字里藏着停顿，从停顿结束处重新起剪。"""
    inside = [w for w in words if w["s"] >= s - 0.02 and w["e"] <= e + 0.02]
    for w in inside:
        if w["s"] - s > 2.6:
            break
        chars = sum(1 for ch in w["w"] if CJK_RE.match(ch))
        if w["e"] - w["s"] < 0.9 or chars > 2:
            continue
        sil = _longest_silence(rms, win, w["s"] + 0.05, w["e"] - 0.05, thr)
        if not sil or sil[1] - sil[0] < 0.3:
            continue
        new_s = round(sil[1] - 0.05, 3)
        if new_s - s < 0.25 or e - new_s < 0.4:
            break
        cut_words = [x for x in inside if x["e"] <= sil[1]]
        if not cut_words:
            break
        pause = new_s - max((x["e"] for x in cut_words), default=sil[0])
        trims.append({"orig": round(s, 3), "kept": new_s, "kind": "stammer",
                      "cut": [round(s, 3), new_s],
                      "text": "剪掉 %.2f–%.2fs：%s（停顿 %.2fs）" % (
                          s, new_s, "".join(x["w"] for x in cut_words), pause)})
        return new_s
    return s


def _bigrams(t: str) -> set:
    t = "".join(ch for ch in t if CJK_RE.match(ch))
    return {t[i:i + 2] for i in range(len(t) - 1)} if len(t) > 1 else ({t} if t else set())


def trim_duplicate_tails(spans: list[dict], chain: list[dict], words: list[dict],
                         sent_texts: list[str], pad: float = 0.03) -> None:
    """前一句把「后面重复的半句」吞进去了：把尾巴收回到句尾词，让重复片段落进缺口被整段删掉。

    实测事故：18 号句的尾巴被吸附到后面重复的「那海南这边的园区地址费用呢」之后，
    用户听到重复，而且 3 秒静音也没被剪。判定条件：句尾多出来的这段语音，与下一句的
    2-gram 重合率 ≥ 0.5 才算"重复"，否则只是正常尾音（如「讲清楚」），必须留着。
    """
    for i in range(len(spans) - 1):
        a, b = spans[i], spans[i + 1]
        if b["s"] <= a["e"]:
            continue
        extra = [w for w in words if w["s"] >= chain[i]["e"] - 0.02 and w["e"] <= a["e"] + 0.02]
        if not extra:
            continue
        be = _bigrams("".join(w["w"] for w in extra))
        bn = _bigrams(sent_texts[b["si"]] if b["si"] < len(sent_texts) else "")
        if not be or not bn:
            continue
        if len(be & bn) / len(be) >= 0.5:
            a["e"] = min(a["e"], round(chain[i]["e"] + pad, 3))


def merge_spans(spans, words, sent_texts, gap_cut: float):
    """短停顿直接并进同一段（不切）；长停顿里若只有短暂衔接词（如「那以后呢」）也并进来。"""
    out = []
    for sp in spans:
        if out:
            prev = out[-1]
            gap = sp["s"] - prev["e"]
            if gap <= gap_cut:
                prev["e"] = max(prev["e"], sp["e"])
                continue
        out.append(dict(sp))
    return out


def split_blocks(pieces, runs, gap_cut: float, keep: float, trims: list[dict]):
    """块内 ≥gap_cut 的静音切掉，只留 keep（长气口）。"""
    out = []
    for p in pieces:
        gaps = []
        for r1, r2 in zip(runs, runs[1:]):
            gs, ge = r1[1], r2[0]
            if ge - gs >= gap_cut and gs >= p["s"] - 0.02 and ge <= p["e"] + 0.02:
                gaps.append((gs, ge))
        start = p["s"]
        for gs, ge in gaps:
            cut_end = round(gs + keep / 2, 3)
            resume = round(ge - keep / 2, 3)
            if cut_end <= start or resume >= p["e"]:
                continue
            out.append({"s": start, "e": cut_end, "si": p["si"], "note": "长气口"})
            trims.append({"orig": round(ge - gs, 3), "kept": round(keep, 3), "kind": "inner",
                          "cut": [cut_end, resume], "text": ""})
            start = resume
        out.append({"s": start, "e": p["e"], "si": p["si"], "note": p.get("note", "")})
    return out


def detect_exclusions(words: list[dict], spans: list[tuple[float, float]]) -> list[dict]:
    """找孤立填充词与同词重复卡壳，返回要挖掉的区间。"""
    out = []
    n = len(words)
    for idx, w in enumerate(words):
        text = norm(w["w"])
        if not text:
            continue
        inside = any(w["s"] >= a - 0.01 and w["e"] <= b + 0.01 for a, b in spans)
        if not inside:
            continue
        prev_gap = w["s"] - words[idx - 1]["e"] if idx > 0 else 9.0
        next_gap = words[idx + 1]["s"] - w["e"] if idx + 1 < n else 9.0
        if len(text) == 1 and text in FILLERS and prev_gap >= 0.12 and next_gap >= 0.12:
            out.append({"s": w["s"] - 0.05, "e": w["e"] + 0.05, "note": f"填充词「{w['w'].strip()}」"})
        elif LAUGH.match(text) and prev_gap >= 0.2 and next_gap >= 0.2:
            out.append({"s": w["s"] - 0.05, "e": w["e"] + 0.05, "note": "笑场"})
        elif idx > 0:
            pt = norm(words[idx - 1]["w"])
            if pt == text and len(text) >= 2 and prev_gap < 0.6:
                out.append({"s": words[idx - 1]["s"] - 0.05, "e": words[idx - 1]["e"] + 0.05,
                            "note": f"重复卡壳「{text}」"})
    return out


def subtract(spans: list[dict], exclusions: list[dict]) -> list[dict]:
    pieces = []
    for sp in spans:
        cur = [{"s": sp["s"], "e": sp["e"], "si": sp["si"], "note": sp.get("note", "")}]
        for ex in exclusions:
            nxt = []
            for p in cur:
                if ex["e"] <= p["s"] or ex["s"] >= p["e"]:
                    nxt.append(p)
                    continue
                if ex["s"] > p["s"]:
                    nxt.append({"s": p["s"], "e": ex["s"], "si": p["si"], "note": p["note"]})
                if ex["e"] < p["e"]:
                    nxt.append({"s": ex["e"], "e": p["e"], "si": p["si"], "note": p["note"]})
            cur = nxt
        pieces.extend(cur)
    return [p for p in pieces if p["e"] - p["s"] > 0.08]


def trim_gaps(pieces: list[dict], gap_cut: float, keep_sentence: float, keep_inner: float,
              trims: list[dict], words: list[dict]) -> list[dict]:
    """两个语音段之间超过 gap_cut 的静音，只保留 target，其余切掉。"""
    for a, b in zip(pieces, pieces[1:]):
        if b["s"] <= a["e"]:
            continue
        g = b["s"] - a["e"]
        if g <= gap_cut:
            continue
        target = keep_inner if ("长气口" in a["note"] or "长气口" in b["note"]) else keep_sentence
        half = min(target / 2, g / 2)
        cut_s, cut_e = a["e"] + half, b["s"] - half
        dropped = "".join(w["w"] for w in words if w["s"] < cut_e - 0.02 and w["e"] > cut_s + 0.02 and w["e"] <= cut_e - 0.05)
        a["e"] += half
        b["s"] -= half
        trims.append({"orig": round(g, 3), "kept": round(2 * half, 3), "kind": "gap",
                      "cut": [round(cut_s, 3), round(cut_e, 3)], "text": dropped.strip()})
    return pieces


def inside_word(t: float, words: list[dict], margin: float = 0.02) -> dict | None:
    starts = [w["s"] for w in words]
    i = bisect.bisect_right(starts, t) - 1
    for k in (i, i + 1):
        if 0 <= k < len(words):
            w = words[k]
            if w["s"] + margin < t < w["e"] - margin:
                return w
    return None


def snap_to_words(t: float, words: list[dict], margin: float = 0.02) -> float:
    """切点若落进某个字里，挪到最近的一侧词边界外侧，避免剪到字。"""
    w = inside_word(t, words, margin)
    if w is None:
        return t
    return max(0.0, w["s"] - 0.02) if (t - w["s"]) <= (w["e"] - t) else w["e"] + 0.02


def quantize(t: float, fps: float) -> float:
    return round(t * fps) / fps


def build_segments(chain: list[dict], words: list[dict], media: Media, cfg: argparse.Namespace,
                   audio=None, sr: int = 16000, sent_texts: list[str] | None = None) -> dict:
    import numpy as np
    trims: list[dict] = []
    runs = speech_runs(audio, sr) if audio is not None else []
    if audio is not None:
        rms, win = _frame_rms(audio, sr)
        speech = rms[rms > 0.02]
        thr = max(0.012, float(np.median(speech)) * 0.2) if len(speech) else 0.02
    else:
        rms, win, thr = np.array([]), 0.02, 0.02

    spans = []
    for st in chain:
        s, e = st["s"], st["e"]
        if audio is not None:
            s = trim_stammer_at(s, e, words, rms, win, thr, trims)
            s, e = snap_span(s, e, runs)
            e = extend_tail(e, runs, words)
        spans.append({"s": s, "e": e, "si": st["i"], "note": ""})
    if sent_texts:
        trim_duplicate_tails(spans, chain, words, sent_texts)
        spans = merge_spans(spans, words, sent_texts, cfg.gap_cut)
    spans[0]["s"] = max(0.0, spans[0]["s"] - cfg.lead)
    spans[-1]["e"] = min(media.duration, spans[-1]["e"] + cfg.tail)

    exclusions = detect_exclusions(words, [(x["s"], x["e"]) for x in spans])
    pieces = subtract(spans, exclusions)
    if audio is not None:
        pieces = split_blocks(pieces, runs, cfg.gap_cut, cfg.keep_inner, trims)
    pieces = trim_gaps(pieces, cfg.gap_cut, cfg.keep_sentence, cfg.keep_inner, trims, words)

    merged = []
    for p in sorted(pieces, key=lambda x: x["s"]):
        if merged and p["s"] <= merged[-1]["e"] + 0.04:
            merged[-1]["e"] = max(merged[-1]["e"], p["e"])
            if p["note"]:
                merged[-1]["note"] = (merged[-1]["note"] + " " + p["note"]).strip()
        else:
            merged.append(dict(p))

    segs = []
    for p in merged:
        s = quantize(p["s"], media.fps)
        e = quantize(p["e"], media.fps)
        if e - s < 0.2:
            continue
        segs.append({"s": s, "e": e, "si": p["si"], "note": p["note"]})
    for a, b in zip(segs, segs[1:]):
        if b["s"] < a["e"]:
            b["s"] = a["e"]


    cycle = [float(x) for x in cfg.zoom_cycle.split(",") if x.strip()]
    for i, seg in enumerate(segs):
        seg["zoom"] = 1.0 if cfg.no_zoom else cycle[i % len(cycle)]
    return {"segments": segs, "exclusions": exclusions, "trims": trims}


# --------------------------------------------------------------------------- #
# 渲染
# --------------------------------------------------------------------------- #
def build_filter(segs: list[dict], media: Media, fps: float, speed: float = 1.0) -> str:
    n = len(segs)
    W, H = media.width, media.height
    parts = [f"[0:v]split={n}" + "".join(f"[vs{i}]" for i in range(n)),
             f"[0:a]asplit={n}" + "".join(f"[as{i}]" for i in range(n))]
    for i, seg in enumerate(segs):
        z = float(seg.get("zoom", 1.0))
        v = f"[vs{i}]trim=start={seg['s']:.6f}:end={seg['e']:.6f},setpts=PTS-STARTPTS"
        if z > 1.0001:
            w = int(round(W * z / 2)) * 2
            h = int(round(H * z / 2)) * 2
            v += f",scale={w}:{h}:flags=lanczos,crop={W}:{H}:(iw-ow)/2:(ih-oh)/2"
        v += f",setsar=1[v{i}]"
        parts.append(v)
        d = seg["e"] - seg["s"]
        a = f"[as{i}]atrim=start={seg['s']:.6f}:end={seg['e']:.6f},asetpts=PTS-STARTPTS"
        if d > 0.03:
            a += f",afade=t=in:st=0:d=0.005,afade=t=out:st={d - 0.005:.6f}:d=0.005"
        a += f"[a{i}]"
        parts.append(a)
    parts.append("".join(f"[v{i}][a{i}]" for i in range(n)) + f"concat=n={n}:v=1:a=1[cv][ca]")
    if abs(speed - 1.0) > 1e-6:
        parts.append(f"[cv]setpts=PTS/{speed:.6f},fps={fps:g}[outv]")
        parts.append(f"[ca]atempo={speed:.6f}[outa]")
    else:
        parts.append(f"[cv]fps={fps:g}[outv]")
        parts.append("[ca]anull[outa]")
    return ";\n".join(parts)


def render(ffmpeg: str, media: Media, segs: list[dict], out_path: str, preset: str, work: Path,
           two_pass: bool = True, speed: float = 1.0) -> None:
    filters = build_filter(segs, media, media.fps, speed)
    script = work / "filters.txt"
    script.write_text(filters, encoding="utf-8")
    vb = media.v_bitrate or 8_000_000
    base = [
        ffmpeg, "-hide_banner", "-loglevel", "warning", "-stats", "-y",
        "-i", media.path,
        "-filter_complex", filters,
        "-map", "[outv]", "-map", "[outa]",
        "-c:v", "libx264", "-preset", preset,
        "-b:v", str(vb), "-maxrate", str(vb), "-bufsize", str(vb * 2),
        "-pix_fmt", "yuv420p", "-r", f"{media.fps:g}",
        "-c:a", "aac", "-b:a", str(media.a_bitrate), "-ar", str(media.sample_rate), "-ac", str(media.channels),
        "-movflags", "+faststart",
    ]
    passlog = str(work / "x264pass")
    print("渲染中（" + ("%d 段" % len(segs)) + "，" + ("两遍" if two_pass else "单遍") + "）…", flush=True)
    if two_pass:
        p1 = base + ["-pass", "1", "-passlogfile", passlog, "-an", "-f", "null", "NUL"]
        r1 = subprocess.run(p1, capture_output=True)
        if r1.returncode != 0:
            raise SystemExit("渲染失败(第1遍):\n" + r1.stderr.decode("utf-8", "replace")[-2000:])
        cmd = base + ["-pass", "2", "-passlogfile", passlog, out_path]
    else:
        cmd = base + [out_path]
    r = subprocess.run(cmd, capture_output=True)
    if r.returncode != 0:
        err = r.stderr.decode("utf-8", "replace")
        raise SystemExit("渲染失败:\n" + err[-2500:])
    for leftover in work.glob("x264pass-*"):
        try:
            leftover.unlink()
        except OSError:
            pass


# --------------------------------------------------------------------------- #
# 报告
# --------------------------------------------------------------------------- #
def write_reports(work: Path, media: Media, chain: list[dict], sentences: list[dict],
                  cands: list[list[dict]], segs: list[dict], exclusions: list[dict],
                  stats: dict, out_path: str) -> None:
    out_probe = None
    if out_path and Path(out_path).exists():
        try:
            out_probe = probe_media(load_env()["ffprobe"], out_path)
        except Exception:
            out_probe = None

    with (work / "剪点表.csv").open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["序号", "素材入点", "素材出点", "时长", "景别", "对应台词", "备注"])
        for i, seg in enumerate(segs, 1):
            si = seg["si"]
            text = sentences[si]["text"] if si < len(sentences) else ""
            w.writerow([i, f"{seg['s']:.3f}", f"{seg['e']:.3f}", f"{seg['e'] - seg['s']:.3f}",
                        f"x{seg.get('zoom', 1.0):.3f}", text, seg.get("note", "")])

    total = out_probe.duration if out_probe else stats["out_duration"]
    cut_pct = (1 - total / media.duration) * 100
    lines = [
        "# 口播粗剪报告",
        "",
        f"- 素材：`{media.path}`",
        f"- 成片：`{out_path}`",
        f"- 原片时长：{media.duration:.2f}s → 粗剪时长：{total:.2f}s"
        f"（删掉 {media.duration - total:.2f}s，压缩 {cut_pct:.1f}%）",
        f"- 分辨率/帧率：{media.width}x{media.height} / {media.fps:g}fps（与原素材一致）",
        f"- 倍速：{stats.get('speed', 1.0):g}x（atempo 不变调）",
        f"- 码率：原 {media.v_bitrate / 1e6:.2f} Mbps → 目标 {stats['out_bitrate'] / 1e6:.2f} Mbps"
        + (f"（ffprobe 实测 {out_probe.v_bitrate / 1e6:.2f} Mbps）" if out_probe else ""),
        "",
        "## 一、台词匹配校准",
        "",
        f"- 台词分句：{len(sentences)} 句；采用：{stats['matched']} 句"
        f"（{stats['matched'] / max(1, len(sentences)) * 100:.1f}%）",
        f"- 平均相似度：{stats['avg_sim']:.2f}；平均字符覆盖：{stats['avg_cov'] * 100:.1f}%",
        "",
        "| # | 台词 | 识别结果 | 相似度 | 覆盖 | 候选 | 采用入点 |",
        "|---|---|---|---|---|---|---|",
    ]
    for i, s in enumerate(sentences):
        chosen = next((c for c in chain if c["i"] == i), None)
        occ = cands[i] if i < len(cands) else []
        brief = s["text"] if len(s["text"]) <= 24 else s["text"][:23] + "…"
        asr = s.get("asr", "")
        if chosen:
            lines.append(f"| {i + 1} | {brief} | {asr} | {chosen['sim']:.2f} | "
                         f"{chosen['cov'] * 100:.0f}% | {len(occ)} | {chosen['s']:.2f}s |")
        else:
            lines.append(f"| {i + 1} | {brief} | {asr} | — | — | {len(occ)} | 未采用 |")

    lines += ["", "## 二、重录（同一段只保留最后一次）", ""]
    if stats["retakes"]:
        for r in stats["retakes"]:
            lines.append(f"- `{r['text']}`：完整出现 {r['total']} 次，保留第 {r['kept']} 次"
                         f"（{r['kept_time']:.2f}s），丢掉更早的 {r['dropped']}")
    if stats.get("stammer"):
        for d in stats["stammer"]:
            lines.append(f"- 起句卡壳重来：{d['text']}")
    if stats.get("dropped_gaps"):
        for d in stats["dropped_gaps"]:
            lines.append(f"- 丢弃片段 {d['cut'][0]:.2f}–{d['cut'][1]:.2f}s（{d['text'][:48]}）")
    if stats.get("dropped_takes"):
        for d in stats["dropped_takes"]:
            lines.append(f"- 卡壳重来片段（未采用）：{d['text']} @ {d['s']:.2f}–{d['e']:.2f}s，覆盖 {d['cov'] * 100:.0f}%")
    if not stats["retakes"] and not stats.get("dropped_takes") and not stats.get("dropped_gaps"):
        lines.append("- 未发现同一句台词的重复录制。")

    lines += ["", "## 三、气口与卡壳", "",
              f"- 剪掉的长停顿（>{stats['gap_cut']:.2f}s）：{stats['gaps_removed']} 处，共 {stats['gap_seconds']:.2f}s；"
              f"断句停顿保留约 {stats['keep_sentence']:.2f}s，句内长气口保留约 {stats['keep_inner']:.2f}s。",
              f"- 剪掉的填充词/重复卡壳：{len(exclusions)} 处。"]
    for ex in exclusions[:20]:
        lines.append(f"  - {ex['note']} @ {ex['s']:.2f}s")
    if stats["weak_sentences"]:
        lines += ["", "## 四、需要人工确认", ""]
        for w in stats["weak_sentences"]:
            lines.append(f"- `{w['text']}`：相似度 {w['sim']:.2f}、覆盖 {w['cov'] * 100:.0f}%，识别为「{w['asr']}」")
    lines += ["", "## 五、切景（换景别）", "",
              f"- 共 {len(segs)} 段，按 {stats['zoom_cycle']} 交替缩放；每个气口切点换一次景别，切点均在词间静音处。"]
    if out_probe:
        lines += ["", "## 六、成片校验", "",
                  f"- {out_probe.width}x{out_probe.height} / {out_probe.fps:g}fps / "
                  f"{out_probe.duration:.2f}s / 视频 {out_probe.v_bitrate / 1e6:.2f} Mbps / "
                  f"音频 {out_probe.a_bitrate / 1000:.0f} kbps"]

    (work / "粗剪报告.md").write_text("\n".join(lines), encoding="utf-8")



# --------------------------------------------------------------------------- #
# 主流程
# --------------------------------------------------------------------------- #
def main() -> None:
    ap = argparse.ArgumentParser(description="口播视频粗剪：匹配口播稿，去重录/卡壳/长气口，原分辨率原码率输出")
    ap.add_argument("--video", help="原素材视频")
    ap.add_argument("--script", help="口播稿 .docx/.txt/.md")
    ap.add_argument("--out-dir", help="输出目录，默认 <视频目录>/<视频名>_粗剪")
    ap.add_argument("--paragraph", default="auto", help="取第几段正文，或 auto（默认）/all")
    ap.add_argument("--gap-cut", type=float, default=0.5, help="多长的停顿算长气口，默认 0.5s")
    ap.add_argument("--keep-sentence", type=float, default=0.12, help="断句停顿保留时长，默认 0.12s")
    ap.add_argument("--keep-inner", type=float, default=0.10, help="句内长气口保留时长，默认 0.10s")
    ap.add_argument("--lead", type=float, default=0.12, help="片头预留，默认 0.12s")
    ap.add_argument("--tail", type=float, default=0.2, help="片尾预留，默认 0.2s")
    ap.add_argument("--zoom-cycle", default="1.0,1.09", help="交替景别，默认 1.0,1.09")
    ap.add_argument("--no-zoom", action="store_true", help="不换景别（纯硬切）")
    ap.add_argument("--min-cov", type=float, default=0.8, help="台词覆盖率达到多少才算匹配上，默认 0.8")
    ap.add_argument("--min-sim", type=float, default=0.72, help="相似度门槛，默认 0.72")
    ap.add_argument("--calibrate", dest="calibrate", action="store_true", default=False,
                    help="整段重识别一遍做校准（慢约一倍，默认关）")
    ap.add_argument("--no-calibrate", dest="calibrate", action="store_false")
    ap.add_argument("--overrides", help="人工指定台词对应的素材时间，JSON: {\"3\": [12.3, 15.1]}")
    ap.add_argument("--dry-run", action="store_true", help="只出方案和剪点表，不渲染")
    ap.add_argument("--from-plan", help="用已有 plan.json 直接渲染")
    ap.add_argument("--force-asr", action="store_true", help="忽略缓存的识别结果")
    ap.add_argument("--preset", default="medium", help="x264 预设，默认 medium")
    ap.add_argument("--speed", type=float, default=1.0, help="整体倍速，如 1.1；视频 setpts + 音频 atempo 不变调")
    ap.add_argument("--single-pass", action="store_true", help="单遍编码（更快，但码率会略低于原片）")
    cfg = ap.parse_args()

    env = load_env()
    if cfg.from_plan:
        plan = json.loads(Path(cfg.from_plan).read_text("utf-8"))
        media = Media(**plan["media"])
        out_path = plan["out_path"]
        work = Path(cfg.from_plan).parent
        render(env["ffmpeg"], media, plan["segments"], out_path, cfg.preset, work, not cfg.single_pass,
               float(plan.get("speed", 1.0)))
        print("完成：" + out_path)
        return

    if not cfg.video or not cfg.script:
        ap.error("需要 --video 和 --script（或用 --from-plan）")
    video, script = str(Path(cfg.video).resolve()), str(Path(cfg.script).resolve())
    for p in (video, script):
        if not Path(p).exists():
            raise SystemExit("文件不存在: " + p)

    media = probe_media(env["ffprobe"], video)
    stem = Path(video).stem
    work = Path(cfg.out_dir) if cfg.out_dir else Path(video).parent / f"{stem}_粗剪"
    work.mkdir(parents=True, exist_ok=True)
    out_path = str(work / f"{stem}_粗剪.mp4")

    print(f"素材 {media.width}x{media.height} {media.fps:g}fps {media.duration:.2f}s "
          f"{media.v_bitrate / 1e6:.2f}Mbps")

    # 口播稿
    paras = read_text_file(script)
    body = pick_script_paragraphs(paras, cfg.paragraph)
    script_text = "".join(body)
    sentences = [{"text": t, "norm": norm(t)} for t in segment_script(script_text)]
    if not sentences:
        raise SystemExit("口播稿分句为空。")

    # 识别
    from faster_whisper import WhisperModel
    wav_path = str(work / "audio16k.wav")
    asr_json = work / "asr.json"
    if cfg.force_asr or not asr_json.exists():
        if not Path(wav_path).exists():
            extract_wav(env["ffmpeg"], video, wav_path)
        model = WhisperModel(env["model"], device="cpu", compute_type="int8",
                             cpu_threads=min(16, os.cpu_count() or 8))
        prompt = "以下是普通话口播视频的内容，请使用简体中文，带标点符号。"
        segs, words = transcribe(model, read_wav_mono16(wav_path), prompt)
        asr_json.write_text(json.dumps({"segments": segs, "words": words}, ensure_ascii=False), encoding="utf-8")
    else:
        data = json.loads(asr_json.read_text("utf-8"))
        words = data["words"]
    chars = build_chars(words)
    print(f"识别完成：{len(words)} 个词 / {len(chars)} 个字")

    # 第一遍对齐
    cands = [find_candidates(s["norm"], chars) for s in sentences]
    chars_for = [chars] * len(sentences)

    # 可选校准：整段再识别一次（用完整口播稿做提示），逐句取覆盖更高的结果
    if cfg.calibrate:
        model2 = WhisperModel(env["model"], device="cpu", compute_type="int8",
                              cpu_threads=min(16, os.cpu_count() or 8))
        print("校准：换一种解码方式整段重识别一遍…", flush=True)
        _s2, words2 = transcribe(model2, read_wav_mono16(wav_path),
                                 "以下是普通话口播视频的内容，请使用简体中文，带标点符号。", condition=True)
        chars2 = build_chars(words2)
        cands2 = [find_candidates(s["norm"], chars2) for s in sentences]
        improved = 0
        for i in range(len(sentences)):
            b1 = max(cands[i], key=lambda x: x["cov"]) if cands[i] else None
            b2 = max(cands2[i], key=lambda x: x["cov"]) if cands2[i] else None
            if b2 and (b1 is None or b2["cov"] > b1["cov"] + 0.05):
                cands[i] = cands2[i]
                chars_for[i] = chars2
                improved += 1
        print(f"校准：{improved} 句改用第二遍结果")

    # 人工指定
    if cfg.overrides:
        ov = json.loads(Path(cfg.overrides).read_text("utf-8"))
        for key, span in ov.items():
            i = int(key) - 1
            if 0 <= i < len(sentences):
                cands[i].insert(0, {"s": float(span[0]), "e": float(span[1]), "cov": 1.0, "sim": 1.0, "manual": True})

    # 每句先筛强候选；没有强候选时保留弱候选并标记
    strong, weak = [], []
    for i, occ in enumerate(cands):
        st = [c for c in occ if c["cov"] >= cfg.min_cov and c["sim"] >= cfg.min_sim]
        strong.append(st)
        if not st and occ:
            wk = [c for c in occ if c["cov"] >= 0.45]
            if wk:
                weak.append((i, max(wk, key=lambda x: x["cov"])))
    chosen_cands = [list(x) for x in strong]
    for i, c in weak:
        chosen_cands[i] = [dict(c, weak=True)]

    chain = choose_chain(chosen_cands)

    if not chain:
        raise SystemExit("一句台词都没匹配上：检查口播稿正文是否选对，或调低 --min-cov。")
    if len(chain) < 0.75 * len(sentences):
        raise SystemExit(
            f"只对上 {len(chain)}/{len(sentences)} 句，对齐率过低。先看报告的对照表，"
            f"确认口播稿正文选对（--paragraph），或把漏句用 --overrides 指到素材时间后再渲染。")

    # 每句的识别文本（取覆盖最高的候选区间），供报告核对
    for i, s in enumerate(sentences):
        if i < len(cands) and cands[i]:
            best = max(cands[i], key=lambda x: x["cov"])
            s["asr"] = text_between(chars_for[i], best["s"], best["e"])

    weak_sentences = []
    for st in chain:
        if st.get("cov", 0) < cfg.min_cov:
            s = sentences[st["i"]]
            weak_sentences.append({"text": s["text"], "sim": st["sim"], "cov": st["cov"],
                                   "asr": s.get("asr", "")})

    if not Path(wav_path).exists():
        extract_wav(env["ffmpeg"], video, wav_path)
    audio16 = read_wav_mono16(wav_path)
    built = build_segments(chain, words, media, cfg, audio16, 16000,
                           [s["norm"] for s in sentences])
    segs = built["segments"]
    if not segs:
        raise SystemExit("剪辑方案为空。")

    out_duration = sum(s["e"] - s["s"] for s in segs) / max(0.1, cfg.speed)
    matched = len(chain)
    stats = {
        "out_duration": out_duration,
        "out_bitrate": media.v_bitrate + media.a_bitrate,
        "cut_pct": (1 - out_duration / media.duration) * 100,
        "matched": matched,
        "avg_sim": sum(st["sim"] for st in chain) / matched,
        "avg_cov": sum(st["cov"] for st in chain) / matched,
        "retakes": [],
        "gaps_removed": 0,
        "gap_seconds": 0.0,
        "gap_cut": cfg.gap_cut,
        "keep_sentence": cfg.keep_sentence,
        "keep_inner": cfg.keep_inner,
        "zoom_cycle": cfg.zoom_cycle,
        "speed": cfg.speed,
        "weak_sentences": weak_sentences,
    }
    # 重录统计
    for i, s in enumerate(sentences):
        occ = [c for c in cands[i] if c["cov"] >= cfg.min_cov]
        kept = [st for st in chain if st["i"] == i]
        if len(occ) >= 2 and kept:
            kt = kept[0]["s"]
            idx = sum(1 for c in occ if c["s"] <= kt)
            stats["retakes"].append({
                "text": s["text"], "total": len(occ), "kept": idx, "kept_time": kt,
                "dropped": "、".join(f"{c['s']:.2f}s" for c in occ if c["s"] > kt),
            })
    # 卡壳重来片段：有候选、但没达到匹配线，也没被选用段覆盖
    dropped = []
    for i, s in enumerate(sentences):
        for c in cands[i]:
            if c.get("manual") or 0.45 > c["cov"]:
                continue
            if c["cov"] >= cfg.min_cov:
                continue
            covered = any(c["s"] < st["e"] - 0.1 and c["e"] > st["s"] + 0.1 for st in chain)
            if not covered:
                dropped.append({"text": s["text"], "s": c["s"], "e": c["e"], "cov": c["cov"]})
    stats["dropped_takes"] = dropped[:20]
    # 气口统计：按实际裁剪记录汇总
    trims = built.get("trims", [])
    gaps = [t for t in trims if t["kind"] == "gap"]
    stats["gaps_removed"] = len(gaps)
    stats["gap_seconds"] = sum(t["orig"] - t["kept"] for t in gaps)
    stam_cuts = [t["cut"] for t in trims if t["kind"] == "stammer"]
    stats["dropped_gaps"] = [t for t in gaps if len(t.get("text", "")) >= 3
                             and not any(not (t["cut"][1] <= x[0] or t["cut"][0] >= x[1]) for x in stam_cuts)]
    stats["stammer"] = [t for t in trims if t["kind"] == "stammer"]
    stats["trims"] = trims

    # 方案落盘
    plan = {
        "video": video,
        "script": script,
        "out_path": out_path,
        "speed": cfg.speed,
        "media": media.__dict__,
        "segments": segs,
        "sentences": [s["text"] for s in sentences],
        "chain": [{k: v for k, v in st.items() if k != "prev"} for st in chain],
        "stats": stats,
            }
    (work / "plan.json").write_text(json.dumps(plan, ensure_ascii=False, indent=1), encoding="utf-8")

    write_reports(work, media, chain, sentences, cands, segs, built["exclusions"], stats, out_path)

    print(f"方案：{len(segs)} 段，{out_duration:.2f}s（原 {media.duration:.2f}s，删 {media.duration - out_duration:.2f}s）")
    print(f"匹配 {matched}/{len(sentences)} 句，平均相似度 {stats['avg_sim']:.2f}")
    print("剪点表: " + str(work / "剪点表.csv"))

    if cfg.dry_run:
        print("dry-run：未渲染。确认方案后运行：python roughcut.py --from-plan \"" + str(work / "plan.json") + "\"")
        return

    render(env["ffmpeg"], media, segs, out_path, cfg.preset, work, not cfg.single_pass, cfg.speed)
    write_reports(work, media, chain, sentences, cands, segs, built["exclusions"], stats, out_path)
    print("完成：" + out_path)
    try:
        out_media = probe_media(env["ffprobe"], out_path)
        print(f"校验：{out_media.width}x{out_media.height} {out_media.fps:g}fps "
              f"{out_media.duration:.2f}s {out_media.v_bitrate / 1e6:.2f}Mbps")
    except Exception:
        pass


if __name__ == "__main__":
    main()
