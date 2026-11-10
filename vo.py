#!/usr/bin/env python3
"""
vo.py - ElevenLabs voiceover inside PowerPoint decks, slide by slide.

On each slide, add a comment (or a speaker-notes line) that starts with "VO:" followed by
the text to read. The script reads those lines, shows how many credits are needed, asks for
confirmation, generates the audio and embeds it in the slide: during the slideshow it plays
on its own and the slide advances when the audio ends (or on click, with --manual).
Saves <name>_vo.pptx next to the original.

Re-running is safe: audio is replaced, not duplicated, and only slides whose text changed
are regenerated (and paid for). All options: python vo.py --help
"""
import argparse, glob, hashlib, io, json, os, re, shutil, subprocess, sys, time, zipfile
from datetime import datetime
from pathlib import Path

VERSION = "2.0"

try:
    sys.stdout.reconfigure(errors="replace")
    sys.stderr.reconfigure(errors="replace")
except Exception:
    pass


class VOError(Exception):
    """An error with a readable message and, when possible, what to do about it."""
    def __init__(self, msg, hint=None, fatal=True, kind=None):
        super().__init__(msg)
        self.hint, self.fatal, self.kind = hint, fatal, kind


try:
    from lxml import etree
    from pptx import Presentation
    from pptx.exc import PackageNotFoundError
    from pptx.opc.package import Part
    from pptx.oxml import parse_xml
except ImportError as e:
    sys.exit(f"ERROR: missing Python library ({e.name}).\n"
             f"Fix: pip install python-pptx lxml requests")

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
    "p188": "http://schemas.microsoft.com/office/powerpoint/2018/8/main",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}
RT_AUDIO = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio"
RT_MEDIA = "http://schemas.microsoft.com/office/2007/relationships/media"
SHAPE_PREFIX = "VO-auto"          # name of the shapes this script creates (used to replace them)
API = os.environ.get("ELEVENLABS_API_BASE", "https://api.elevenlabs.io")   # override for tests only
MAX_CHARS_PER_SLIDE = 9500        # below ElevenLabs' per-request limit
TAG_RE = re.compile(r"^\s*VO\s*:\s*", re.I)
YES = ("y", "yes", "s", "si", "sì")

CFG = {
    "model_id": "eleven_multilingual_v2",
    "output_format": "mp3_44100_128",
    "voice_settings": {"stability": 0.5, "similarity_boost": 0.75, "style": 0.0,
                       "use_speaker_boost": True, "speed": 0.9},
    "lead_in": 0.8,          # seconds on the slide before the voice starts
    "tail": 2.0,             # seconds of silence after the voice (time to read/look)
    "silent_slide": 4.0,     # duration of slides without "VO:"
    "pause_segment": 1.0,    # pause between two VO: lines on the same slide (max 3)
    "loudness": -16,         # integrated LUFS per slide: even volume across the deck
}


def q(tag):
    pfx, name = tag.split(":")
    return "{%s}%s" % (NS[pfx], name)


def run(cmd, timeout=300, **kw):
    # stdin closed and a timeout: ffmpeg must never hang
    return subprocess.run(cmd, check=True, capture_output=True, text=True, stdin=subprocess.DEVNULL,
                          timeout=timeout, **kw)


def probe_duration(path):
    out = run(["ffprobe", "-v", "error", "-show_entries", "format=duration", "-of", "csv=p=0", str(path)], timeout=60)
    return float(out.stdout.strip())


def audio_ok(path):
    """The file exists, is not empty and ffprobe can read a duration."""
    p = Path(path)
    if not p.exists() or p.stat().st_size < 1000:
        return False
    try:
        return probe_duration(p) > 0.2
    except Exception:
        return False


def load_dotenv():
    """Reads ELEVENLABS_* variables from a .env file next to the script or in the current folder."""
    for f in (Path(__file__).resolve().with_name(".env"), Path.cwd() / ".env"):
        if f.exists():
            for line in f.read_text(encoding="utf-8-sig").splitlines():
                m = re.match(r"\s*([A-Z_][A-Z0-9_]*)\s*=\s*(.*?)\s*$", line)
                if m and m.group(1).startswith("ELEVENLABS_") and not line.lstrip().startswith("#"):
                    os.environ.setdefault(m.group(1), m.group(2).strip("'\""))


# --------------------------------------------------------------------------- reading the text

def _text_of(el):
    paras = ["".join(t.text or "" for t in p.iter(q("a:t"))) for p in el.iter(q("a:p"))]
    return "\n".join(paras).strip()


def slide_comments(slide):
    """Top-level comments of the slide (modern and legacy); replies are ignored."""
    out = []
    for rel in slide.part.rels.values():
        if rel.is_external or not rel.reltype.endswith("/comments"):
            continue
        try:
            root = etree.fromstring(rel.target_part.blob)
        except Exception:
            continue
        for cm in root.iter(q("p188:cm")):
            body = cm.find(q("p188:txBody"))
            if body is not None:
                out.append(_text_of(body))
        for cm in root.iter(q("p:cm")):
            t = cm.find(q("p:text"))
            if t is not None and t.text:
                out.append(t.text.strip())
    return [c for c in out if c]


def slide_notes_lines(slide):
    """Lines of the speaker notes (empty if the slide has no notes)."""
    if not slide.has_notes_slide:
        return []
    tf = slide.notes_slide.notes_text_frame
    if tf is None:
        return []
    return [ln for p in tf.paragraphs for ln in p.text.replace("\v", "\n").split("\n")]


_PAREN_RE = re.compile(r"^\s*\([^)]*\)\s*")          # stage directions: (minute 0), (1)...
# quoted text, optionally followed by a stage direction: "..." (serious tone)
_QUOTED_RE = re.compile(r'^["“„«](.*)["”“»]\s*(\([^)]*\))?\s*$', re.S)
_CLICK_RE = re.compile(r"^\(\s*(?:click|clic)\s*(\d+)\s*\)", re.I)
_PAUSE_RE = re.compile(r"\[(?:pause|pausa):\s*([\d.,]+)\s*s?\]", re.I)


def _clean_segment(s):
    """'VO: (click 2) "Text"' -> ('Text', 2). Text in parentheses is not read;
    '(click N)' says the line goes with the N-th animation click (None if absent)."""
    s = TAG_RE.sub("", s, count=1).strip()
    click = None
    while _PAREN_RE.match(s):
        m = _CLICK_RE.match(s.strip())
        if m and click is None:
            click = int(m.group(1))
        s = _PAREN_RE.sub("", s, count=1)
    m = _QUOTED_RE.match(s.strip())
    if m:
        s = m.group(1)
    return re.sub(r"\s+", " ", s).strip(), click


def slide_vo_segments(slide):
    """([(text, click)], source). 'VO:' comments first; if there are none, 'VO:' lines in the notes."""
    segs = [_clean_segment(c) for c in slide_comments(slide) if TAG_RE.match(c)]
    segs = [s for s in segs if s[0]]
    if segs:
        return segs, "comments"
    segs = [_clean_segment(ln) for ln in slide_notes_lines(slide) if TAG_RE.match(ln)]
    segs = [s for s in segs if s[0]]
    return segs, ("notes" if segs else "")


def slide_vo_text(slide, cfg):
    """(text for TTS, VO: lines [{"chars", "pause_before", "click"}], source).
    The whole slide is generated in a single request, so the intonation stays natural."""
    segs, source = slide_vo_segments(slide)
    v3 = cfg["model_id"] == "eleven_v3"

    def brk(m):
        s = min(float(m.group(1).replace(",", ".")), 3.0)
        return "" if v3 else f' <break time="{s:.1f}s" /> '

    parts = [_PAUSE_RE.sub(brk, s) for s, _ in segs]
    gap = 0.0 if v3 else min(cfg["pause_segment"], 3.0)
    sep = " " if v3 else f' <break time="{gap:.1f}s" /> '
    text = re.sub(r"\s+", " ", sep.join(parts)).strip()
    # to estimate when each line starts: spoken characters and pauses (<break>) before it
    meta, chars, pauses = [], 0, 0.0
    for i, (p, (_, click)) in enumerate(zip(parts, segs)):
        if i:
            pauses += gap
        meta.append({"chars": chars, "pause_before": pauses, "click": click})
        chars += len(re.sub(r"<[^>]+>", "", p).strip())
        pauses += sum(float(x) for x in re.findall(r'<break time="([\d.]+)s"', p))
    for m in meta:
        m["tot_chars"], m["tot_pause"] = chars, pauses
    return text, meta, source


def segment_starts_ms(meta, dur_ms, resumes_ms=()):
    """Start (ms from audio start) of each VO: line. Estimate: pauses are known, the rest of the
    duration is split in proportion to spoken characters. The estimate is then snapped to the
    closest point where the voice resumes after a real pause (resumes_ms), if any."""
    out, free = [], list(resumes_ms)
    for i, m in enumerate(meta):
        speech = max(dur_ms - m["tot_pause"] * 1000, dur_ms * 0.5)
        t = int(m["chars"] / max(m["tot_chars"], 1) * speech + m["pause_before"] * 1000)
        if i and free:
            r = min(free, key=lambda x: abs(x - t))
            if abs(r - t) <= 2000 and (not out or r > out[-1]):
                t = max(r - 150, 0)                   # just before the voice resumes
                free.remove(r)
        out.append(t)
    return out


def voice_resumes_ms(path, min_pause):
    """Moments (ms) where the voice resumes after at least min_pause seconds of silence."""
    try:
        err = subprocess.run(["ffmpeg", "-v", "info", "-i", str(path), "-af",
                              f"silencedetect=noise=-40dB:d={min_pause:.2f}", "-f", "null", "-"],
                             capture_output=True, text=True, stdin=subprocess.DEVNULL, timeout=120).stderr
        return [int(float(x) * 1000) for x in re.findall(r"silence_end: ([\d.]+)", err)]
    except (OSError, subprocess.SubprocessError):
        return []


# --------------------------------------------------------------------------- ElevenLabs

def explain(status, body):
    """Turns an ElevenLabs error into a message that says what to do."""
    try:
        d = json.loads(body).get("detail", {})
    except Exception:
        d = {}
    d = d if isinstance(d, dict) else {"message": str(d)}
    code = f"{d.get('code', '')} {d.get('status', '')}".lower()
    msg = d.get("message") or body[:300]
    if "quota" in code:
        return VOError(f"Out of ElevenLabs credits. {msg}",
                       "Wait for the monthly reset or upgrade your plan. Audio already generated stays "
                       "in the cache: re-running will not charge for it again.", kind="quota")
    if "missing_permissions" in code:
        return VOError(f"The API key is missing a required permission. {msg}",
                       "At elevenlabs.io/app/api/api-keys, edit the key and enable Text to Speech.",
                       kind="permission")
    if "voice_not_found" in code or (status == 404 and "voice" in msg.lower()):
        return VOError("The voice does not exist or is not available to this account.",
                       "Check ELEVENLABS_VOICE_ID, or run: python vo.py --voices", kind="voice")
    if "free" in msg.lower() and "voice" in msg.lower():
        return VOError(f"This voice is not available on the free plan. {msg}",
                       "Pick a default voice (python vo.py --voices) or upgrade your plan.", kind="voice")
    if status == 401:
        return VOError(f"Invalid or revoked API key. {msg}",
                       "Copy the key again from elevenlabs.io/app/api/api-keys into ELEVENLABS_API_KEY.",
                       kind="auth")
    if status == 422:
        return VOError(f"ElevenLabs rejected the parameters. {msg}",
                       "Check --model and --speed (0.7-1.2).")
    return VOError(f"ElevenLabs returned {status}. {msg}", fatal=False)


def api(method, path, key, retries=4, **kw):
    """API call with retries on network errors, 429 and 5xx."""
    try:
        import requests
    except ImportError:
        raise VOError("The requests library is missing.", "pip install requests")
    last = None
    for attempt in range(retries + 1):
        try:
            r = requests.request(method, API + path, headers={"xi-api-key": key}, timeout=(15, 300), **kw)
        except requests.exceptions.RequestException as e:
            last = VOError(f"Network problem reaching ElevenLabs ({type(e).__name__}).",
                           "Check your connection, VPN or corporate proxy, then run again.", fatal=False)
        else:
            if r.status_code < 400:
                return r
            if r.status_code in (429, 500, 502, 503, 504):
                last = explain(r.status_code, r.text)
                last.fatal = False
            else:
                raise explain(r.status_code, r.text)
        if attempt < retries:
            wait = min(60, float(os.environ.get("VO_RETRY_WAIT", 3)) * 2 ** attempt)
            print(f"\n    retrying in {wait:.0f} s ({attempt + 1}/{retries})...", end="", flush=True)
            time.sleep(wait)
    raise last


def get_credits(key):
    """(remaining, limit, reset date) or None if the key cannot read them."""
    try:
        d = api("GET", "/v1/user/subscription", key, retries=1).json()
    except VOError as e:
        if e.kind == "permission":
            return None
        raise
    left = d.get("character_limit", 0) - d.get("character_count", 0)
    reset = d.get("next_character_count_reset_unix")
    return left, d.get("character_limit"), datetime.fromtimestamp(reset).strftime("%Y-%m-%d") if reset else None


def get_voice_name(key, voice_id):
    try:
        return api("GET", f"/v1/voices/{voice_id}", key, retries=1).json().get("name", voice_id)
    except VOError as e:
        if e.kind == "permission":
            return None            # TTS-only key: cannot check, carry on
        raise


def tts(text, key, voice_id, cfg, prev_text, next_text):
    body = {"text": text, "model_id": cfg["model_id"], "voice_settings": cfg["voice_settings"]}
    if cfg["model_id"] != "eleven_v3":
        if prev_text:
            body["previous_text"] = prev_text[-1000:]
        if next_text:
            body["next_text"] = next_text[:1000]
    r = api("POST", f"/v1/text-to-speech/{voice_id}", key, json=body,
            params={"output_format": cfg["output_format"]})
    if len(r.content) < 1000 or "audio" not in r.headers.get("Content-Type", "audio"):
        raise VOError("ElevenLabs returned an empty or non-audio response.", fatal=False)
    return r.content


def tts_fake(text, out_path):
    dur = max(1.0, len(text) / 14)
    run(["ffmpeg", "-v", "error", "-y", "-f", "lavfi", "-i", f"sine=frequency=440:duration={dur:.2f}",
         "-af", "volume=0.2,apad=pad_dur=0.3", "-c:a", "libmp3lame", "-b:a", "128k", str(out_path)])


def postprocess(raw, out, lufs):
    """Trims silence and evens out the volume. Falls back to the original audio if ffmpeg fails.
    Two separate passes: in a single filtergraph (silenceremove+areverse+loudnorm+resampling)
    some ffmpeg versions hang forever."""
    trim = ("silenceremove=start_periods=1:start_threshold=-50dB:start_silence=0.05,"
            "areverse,silenceremove=start_periods=1:start_threshold=-50dB:start_silence=0.12,areverse")
    tmp = Path(out).with_suffix(".trim.wav")
    try:
        run(["ffmpeg", "-v", "error", "-y", "-i", str(raw), "-af", trim, "-c:a", "pcm_s16le", str(tmp)])
        run(["ffmpeg", "-v", "error", "-y", "-i", str(tmp), "-af", f"loudnorm=I={lufs}:TP=-1.5:LRA=11",
             "-ar", "44100", "-ac", "1", "-c:a", "libmp3lame", "-b:a", "160k", str(out)])
        if audio_ok(out):
            return True
    except (subprocess.SubprocessError, OSError):
        pass
    finally:
        tmp.unlink(missing_ok=True)
    shutil.copyfile(raw, out)
    return False


# --------------------------------------------------------------------------- PPTX: removal

def remove_previous_vo(slide):
    """Removes shapes, timing and relationships created by a previous run."""
    spTree = slide.shapes._spTree
    removed_ids, rids = [], set()
    for pic in spTree.findall(q("p:pic")):
        cnv = pic.find(".//" + q("p:cNvPr"))
        if cnv is not None and cnv.get("name", "").startswith(SHAPE_PREFIX):
            removed_ids.append(cnv.get("id"))
            for el in pic.iter():
                for attr, val in el.attrib.items():
                    if attr.startswith("{%s}" % NS["r"]) and val:
                        rids.add(val)
            spTree.remove(pic)
    sld = slide._element
    timing = sld.find(q("p:timing"))
    if timing is not None and removed_ids:
        for tgt in list(timing.iter(q("p:spTgt"))):
            if tgt.get("spid") not in removed_ids:
                continue
            # walk up to the <p:audio> node or to the <p:par> of the mediacall effect (only that:
            # slide animations merged into the same group stay)
            node = tgt
            while node is not None:
                if node.tag == q("p:audio"):
                    break
                if node.tag == q("p:par"):
                    ctn = node.find(q("p:cTn"))
                    if ctn is not None and ctn.get("presetClass") == "mediacall":
                        break
                node = node.getparent()
            if node is None or node.getparent() is None:
                continue
            parent = node.getparent()
            parent.remove(node)
            # prune containers left empty, up to the main sequence
            while parent is not None and parent.tag == q("p:childTnLst") and len(parent) == 0:
                ctn = parent.getparent()
                if ctn is None or ctn.get("nodeType") in ("mainSeq", "tmRoot"):
                    break
                par = ctn.getparent()
                if par is None or par.getparent() is None:
                    break
                parent = par.getparent()
                parent.remove(par)
        ms = timing.find(".//" + q("p:cTn") + "[@nodeType='mainSeq']")
        if ms is not None:
            lst = ms.find(q("p:childTnLst"))
            if lst is None or len(lst) == 0:                # empty main sequence: remove it
                seq = ms.getparent()
                seq.getparent().remove(seq)
        if not list(timing.iter(q("p:spTgt"))) and not list(timing.iter(q("p:sldTgt"))) \
                and not list(timing.iter(q("p:inkTgt"))):
            sld.remove(timing)
        else:
            bld = timing.find(q("p:bldLst"))
            if bld is not None and len(bld) == 0:
                timing.remove(bld)
    still = {v for el in sld.iter() for k, v in el.attrib.items() if k.startswith("{%s}" % NS["r"])}
    for rid in rids - still:
        if rid in slide.part.rels:
            slide.part._rels.pop(rid)
    return bool(removed_ids)


# --------------------------------------------------------------------------- PPTX: insertion

ICON_PNG = "iVBORw0KGgoAAAANSUhEUgAAAIAAAACACAYAAADDPmHLAAACm0lEQVR42u3c0ZGjMBCEYaFyMJB/MJCN7+UerrZqXfYZhKT+OoGV6X96RrPYpRAREVGilqv/wLquT4/5Ox3HsQwDAMPHAmJhfDYIC+OzQajMH1/f+LAwPjsNKvOz06AyPxuCyvxsCCrzsyGoHlO2qurPToHK/GwItAAtQPUnp4AEkACqPzkFJIAEIAAQAPT/zDlAAkgAAgABgABAAKBX2vcdAMwHAPMn0oO9mcZLAOYDIN18AISbbwYINl4CMB8A6ebHA9Cb+fu+Nz9TTTW+R/PvALOq+j7P1Oqclfn9nqnFeSvz79W2bbeeu+r3/UMAgIAr3isIrvwclfnZEFTmF4sg/T43BR6qtM1n/WTQ27bNHmC2dtQr8ABomHJnQHA2SABo3OLeNbDVbgAA4bccAHQ8zbdIAQBMuuIFgGUVAKQAAKYF6KwEAUB4GwBAeBsAQPHfQAIAAYAAQAAgAFDS208ACN8PAGDQ9DgLHgCEv/wKgPD1MADCX30HwA3mv1v9vh5Ol7cPADQ2qafqB0BjCM6o3LOHRwA0gqDXl0IeM12DRv0ZmDs/R52t2kZ4F//q20N8CxgdAnuAUAhaV//0Q+BIENxhfsQtYIS54M7htdqe9W3+1eeutnP9nsnXw4Mh8AshYXPBv+dpebZqTZsNZrWr3yyCQLABAARbJAgACE8DAIRDAIBwCAAQPhcAIDwNABAOAQDCIQBA8UoYAYAAQAAgAJTjOBaPZV799FcCSAACAAHAHJDX/yUA/Q6AFJi/+iUAvQZACsxd/W8lAAjmNV8LoPcAkAJzVv9HCQCC+cz/uAWAYC7z/2sGAME85pdSyldmruv69MjHNP6UW4A0GNv8rxNAGoxr/OkAAGEs4y8DABB9G05ERPRXfwC+cTV8DcMMyAAAAABJRU5ErkJggg=="


def speaker_icon_png():
    import base64
    return io.BytesIO(base64.b64decode(ICON_PNG))


def next_shape_id(slide):
    ids = [int(el.get("id")) for el in slide._element.iter(q("p:cNvPr")) if (el.get("id") or "").isdigit()]
    return max(ids + [1]) + 1


def add_audio_shape(prs, slide, mp3_path, offslide):
    pkg = prs.part.package
    media = Part(pkg.next_media_partname("mp3"), "audio/mpeg", pkg, Path(mp3_path).read_bytes())
    rid_audio = slide.part.relate_to(media, RT_AUDIO)
    rid_media = slide.part.relate_to(media, RT_MEDIA)
    _, rid_img = slide.part.get_or_add_image_part(speaker_icon_png())

    size = 457200  # 0.5"
    W, H = prs.slide_width, prs.slide_height
    x, y = (W + 228600, H - size) if offslide else (W - size - 228600, H - size - 228600)
    spid = next_shape_id(slide)
    xml = f"""
<p:pic xmlns:p="{NS['p']}" xmlns:a="{NS['a']}" xmlns:r="{NS['r']}">
  <p:nvPicPr>
    <p:cNvPr id="{spid}" name="{SHAPE_PREFIX} {spid}" descr="Voiceover">
      <a:hlinkClick r:id="" action="ppaction://media"/>
    </p:cNvPr>
    <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
    <p:nvPr>
      <a:audioFile r:link="{rid_audio}"/>
      <p:extLst><p:ext uri="{{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}}">
        <p14:media xmlns:p14="{NS['p14']}" r:embed="{rid_media}"/>
      </p:ext></p:extLst>
    </p:nvPr>
  </p:nvPicPr>
  <p:blipFill><a:blip r:embed="{rid_img}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
  <p:spPr>
    <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{size}" cy="{size}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>"""
    slide.shapes._spTree.append(parse_xml(xml))
    return spid


def _max_ctn_id(timing):
    ids = [int(c.get("id")) for c in timing.iter(q("p:cTn")) if (c.get("id") or "").isdigit()]
    return max(ids + [0])


def _main_seq(slide):
    """(timing, root node list, mainSeq cTn, its childTnLst), creating them if missing."""
    sld = slide._element
    timing = sld.find(q("p:timing"))
    if timing is None:
        timing = parse_xml(f"""
<p:timing xmlns:p="{NS['p']}"><p:tnLst><p:par>
  <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst/></p:cTn>
</p:par></p:tnLst></p:timing>""")
        ext = sld.find(q("p:extLst"))
        if ext is not None:
            ext.addprevious(timing)
        else:
            sld.append(timing)
    root_ctn = timing.find(".//" + q("p:cTn") + "[@nodeType='tmRoot']")
    root_list = root_ctn.find(q("p:childTnLst"))
    n = _max_ctn_id(timing)

    ms = timing.find(".//" + q("p:cTn") + "[@nodeType='mainSeq']")
    if ms is None:
        n += 1
        seq = parse_xml(f"""
<p:seq xmlns:p="{NS['p']}" concurrent="1" nextAc="seek">
  <p:cTn id="{n}" dur="indefinite" nodeType="mainSeq"><p:childTnLst/></p:cTn>
  <p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>
  <p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>
</p:seq>""")
        root_list.insert(0, seq)
        ms = seq.find(q("p:cTn"))
    ms_list = ms.find(q("p:childTnLst"))
    if ms_list is None:
        ms_list = etree.SubElement(ms, q("p:childTnLst"))
    return timing, root_list, ms, ms_list


def _auto_group(timing, ms, ms_list):
    """Group at the head of the main sequence that starts by itself when the slide opens."""
    a = _max_ctn_id(timing) + 1
    group = parse_xml(f"""
<p:par xmlns:p="{NS['p']}"><p:cTn id="{a}" fill="hold">
  <p:stCondLst><p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="{ms.get('id')}"/></p:cond></p:stCondLst>
  <p:childTnLst/>
</p:cTn></p:par>""")
    ms_list.insert(0, group)
    return group


def add_autoplay(slide, spid, dur_ms, lead_ms):
    """Plays automatically when the slide opens; icon hidden during the slideshow.
    Returns the auto-start group (click animations get merged into it)."""
    timing, root_list, ms, ms_list = _main_seq(slide)
    group = _auto_group(timing, ms, ms_list)
    n = _max_ctn_id(timing)
    b, c, d, e = range(n + 1, n + 5)
    inner = parse_xml(f"""
<p:par xmlns:p="{NS['p']}"><p:cTn id="{b}" fill="hold">
    <p:stCondLst><p:cond delay="0"/></p:stCondLst>
    <p:childTnLst><p:par>
      <p:cTn id="{c}" presetID="1" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="afterEffect">
        <p:stCondLst><p:cond delay="{lead_ms}"/></p:stCondLst>
        <p:childTnLst><p:cmd type="call" cmd="playFrom(0.0)">
          <p:cBhvr><p:cTn id="{d}" dur="{dur_ms}" fill="hold"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr>
        </p:cmd></p:childTnLst>
      </p:cTn>
    </p:par></p:childTnLst>
</p:cTn></p:par>""")
    group.find(q("p:cTn")).find(q("p:childTnLst")).append(inner)

    audio = parse_xml(f"""
<p:audio xmlns:p="{NS['p']}"><p:cMediaNode vol="100000">
  <p:cTn id="{e}" fill="hold" display="0">
    <p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>
    <p:endCondLst><p:cond evt="onStopAudio" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:endCondLst>
  </p:cTn>
  <p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>
</p:cMediaNode></p:audio>""")
    root_list.append(audio)
    return group


def _is_click_group(par):
    """Main-sequence group that waits for a click (no automatic start)."""
    ctn = par.find(q("p:cTn"))
    if par.tag != q("p:par") or ctn is None:
        return False
    conds = ctn.findall(q("p:stCondLst") + "/" + q("p:cond"))
    return bool(conds) and all(c.get("delay") == "indefinite" and c.get("evt") is None and
                               len(c) == 0 for c in conds)


def count_click_groups(slide):
    timing = slide._element.find(q("p:timing"))
    ms = None if timing is None else timing.find(".//" + q("p:cTn") + "[@nodeType='mainSeq']")
    lst = None if ms is None else ms.find(q("p:childTnLst"))
    return 0 if lst is None else sum(1 for p in lst if _is_click_group(p))


def retime_click_animations(slide, start_times_ms, group=None):
    """On-click animations of the main sequence become automatic: the k-th click starts at
    start_times_ms[k] from when the slide opens. Without this, with auto-advance PowerPoint
    would fire them one per advTm interval (very long slides)."""
    # PowerPoint's native layout: all effects go into the first sub-group of the auto-start group
    # (the audio's) as "With Previous", with delay = the target time.
    # ("After Previous" would not work: it would wait for the audio to end.)
    def delay_of(ctn):
        cond = ctn.find(q("p:stCondLst") + "/" + q("p:cond")) if ctn is not None else None
        try:
            return cond, int(cond.get("delay", "0")) if cond is not None else 0
        except ValueError:
            return cond, 0

    timing, _, ms, ms_list = _main_seq(slide)
    clicks = [p for p in ms_list if _is_click_group(p)]
    if not clicks:
        return 0
    if group is None:
        group = _auto_group(timing, ms, ms_list)
    glist = group.find(q("p:cTn")).find(q("p:childTnLst"))
    if len(glist) == 0:
        glist.append(parse_xml(f'<p:par xmlns:p="{NS["p"]}"><p:cTn id="{_max_ctn_id(timing) + 1}" '
                               f'fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst>'
                               f'<p:childTnLst/></p:cTn></p:par>'))
    target = glist[0].find(q("p:cTn")).find(q("p:childTnLst"))
    for k, par in enumerate(clicks):
        t0 = int(start_times_ms[min(k, len(start_times_ms) - 1)])
        inner_list = par.find(q("p:cTn")).find(q("p:childTnLst"))
        for inner in list(inner_list if inner_list is not None else []):
            _, d_inner = delay_of(inner.find(q("p:cTn")))
            eff_list = inner.find(q("p:cTn")).find(q("p:childTnLst"))
            for eff_par in list(eff_list if eff_list is not None else []):
                ectn = eff_par.find(q("p:cTn"))
                cond, d_eff = delay_of(ectn)
                if ectn is None or cond is None:
                    continue
                cond.set("delay", str(t0 + d_inner + d_eff))
                if ectn.get("nodeType") in ("clickEffect", "afterEffect", "withEffect"):
                    ectn.set("nodeType", "withEffect")
                target.append(eff_par)
        ms_list.remove(par)
    return len(clicks)


def click_start_times(n_clicks, meta, lead_ms, dur_ms, resumes_ms=()):
    """When to start each click -> (times in ms from slide open, warning or None).
    1. lines tagged '(click N)': click N starts when that line starts;
    2. no tags: the last n clicks on the last n lines, if there are enough lines;
    3. otherwise spread evenly over the voice duration."""
    if n_clicks <= 0:
        return [], None
    starts = segment_starts_ms(meta, dur_ms, resumes_ms) if meta else []
    labeled = {m["click"]: t for m, t in zip(meta or [], starts) if m["click"]}
    if labeled:
        times, missing, prev = [], [], 0
        for k in range(1, n_clicks + 1):
            if k in labeled:
                prev = labeled[k]
            else:
                missing.append(k)
            times.append(lead_ms + prev)
        extra = sorted(k for k in labeled if k > n_clicks)
        note = None
        if missing or extra:
            note = (f"the slide has {n_clicks} animation clicks"
                    + (f"; no VO: line for click {', '.join(map(str, missing))}" if missing else "")
                    + (f"; lines for non-existent clicks: {', '.join(map(str, extra))}" if extra else ""))
        return times, note
    if len(starts) > 1 and len(starts) >= n_clicks:
        return [lead_ms + t for t in starts[len(starts) - n_clicks:]], None
    return [lead_ms + int(k * dur_ms / n_clicks) for k in range(n_clicks)], None


def set_advance(slide, adv_ms):
    """Auto-advance after adv_ms (clicking still works too)."""
    sld = slide._element
    trans = list(sld.iter(q("p:transition")))          # including those inside mc:AlternateContent
    if not trans:
        t = etree.Element(q("p:transition"))
        anchor = sld.find(q("p:timing"))
        if anchor is None:
            anchor = sld.find(q("p:extLst"))
        if anchor is not None:
            anchor.addprevious(t)
        else:
            clr = sld.find(q("p:clrMapOvr"))
            (clr if clr is not None else sld.find(q("p:cSld"))).addnext(t)
        trans = [t]
    for t in trans:
        t.set("advTm", str(int(adv_ms)))


# --------------------------------------------------------------------------- checks

def check_ffmpeg():
    missing = [t for t in ("ffmpeg", "ffprobe") if not shutil.which(t)]
    if missing:
        hint = ("Windows: winget install ffmpeg, then close and reopen the terminal. "
                "macOS: brew install ffmpeg. Linux: sudo apt install ffmpeg.")
        raise VOError(f"{' and '.join(missing)} not found in PATH.", hint)


def expand_decks(items):
    out = []
    for it in items:
        found = sorted(glob.glob(it)) if any(c in it for c in "*?[") else [it]
        if not found:
            raise VOError(f"No files match {it}.", fatal=False)
        for f in found:
            p = Path(f)
            if p.name.startswith("~$") or (len(found) > 1 and re.search(r"_vo( \(\d+\))?$", p.stem)):
                continue                      # Office lock files and outputs of previous runs
            out.append(p)
    return list(dict.fromkeys(out))


def open_deck(path):
    if not path.exists():
        raise VOError(f"File not found: {path}", "Check the name; if it contains spaces, put it in quotes.",
                      fatal=False)
    if path.suffix.lower() == ".ppt":
        raise VOError(f"{path.name} is in the old .ppt format.",
                      "Open it in PowerPoint and save it as .pptx (File > Save As).", fatal=False)
    if path.suffix.lower() != ".pptx":
        raise VOError(f"{path.name} is not a .pptx file.", fatal=False)
    try:
        return Presentation(str(path))
    except PermissionError:
        raise VOError(f"Cannot read {path.name}: it is locked by another program.",
                      "Close it in PowerPoint and try again.", fatal=False)
    except (PackageNotFoundError, zipfile.BadZipFile, KeyError):
        raise VOError(f"{path.name} does not open as a valid presentation.",
                      "If it is on OneDrive, make sure it is downloaded (right click > Always keep on "
                      "this device). Otherwise open it and save it again in PowerPoint.", fatal=False)


def writable(path):
    """Can the output file be written? (not open in PowerPoint, folder writable)."""
    try:
        if path.exists():
            with open(path, "a+b"):
                pass
        else:
            path.touch()
            path.unlink()
        return True
    except OSError:
        return False


def free_name(path):
    for i in range(2, 100):
        alt = path.with_name(f"{path.stem} ({i}){path.suffix}")
        if not alt.exists() or writable(alt):
            return alt
    raise VOError(f"Cannot find a free name to save next to {path.name}.")


def parse_slides(spec):
    if not spec:
        return None
    out = set()
    try:
        for part in spec.replace(" ", "").split(","):
            lo, _, hi = part.partition("-")
            out.update(range(int(lo), int(hi or lo) + 1))
    except ValueError:
        raise VOError(f"--slides '{spec}' is not valid.", "Examples: --slides 3   --slides 1,4-6")
    return out


def open_file(path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(path))
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except Exception as e:
        print(f"  (cannot open {path.name}: {e})")


# --------------------------------------------------------------------------- plan and generation

def plan_deck(path, args, cfg, voice_id):
    prs = open_deck(path)
    wanted = parse_slides(args.slides)
    cache = Path(args.cache) if args.cache else path.with_name("vo_cache")
    out = Path(args.output) if args.output else path.with_name(path.stem + "_vo.pptx")
    slides = []
    for i, slide in enumerate(prs.slides, 1):
        text, starts, source = slide_vo_text(slide, cfg)
        sel = wanted is None or i in wanted
        k = hashlib.sha1(json.dumps([text, voice_id, cfg["model_id"], cfg["voice_settings"],
                                     "fake" if args.fake else "el"], sort_keys=True).encode()).hexdigest()[:16]
        cached = bool(text) and not args.regenerate and audio_ok(cache / f"{k}.raw.mp3")
        if len(text) > MAX_CHARS_PER_SLIDE:
            raise VOError(f"{path.name}, slide {i}: {len(text)} characters, too long for one request.",
                          f"Split it across slides (max {MAX_CHARS_PER_SLIDE} characters).", fatal=False)
        slides.append({"n": i, "text": text if sel else "", "key": k, "cached": cached, "sel": sel,
                       "starts": starts, "source": source,
                       "hidden": slide._element.get("show") == "0"})
    todo = [s for s in slides if s["text"] and not s["cached"]]
    return {"path": path, "prs": prs, "out": out, "cache": cache, "slides": slides, "todo": todo,
            "chars": sum(len(s["text"]) for s in todo)}


def print_plan(p):
    print(f"\n{p['path'].name}  ({len(p['slides'])} slides)")
    shown = [s for s in p["slides"] if s["text"]]
    for s in shown:
        state = "cached" if s["cached"] else f"{len(s['text'])} chars"
        txt = re.sub(r'<break time="([\d.]+)s" />', r"[\1s]", s["text"])
        txt = txt if len(txt) <= 70 else txt[:69] + "…"
        print(f"  slide {s['n']:>2} [{state:>9}]  {txt}")
    if not shown:
        print("  no 'VO:' found (neither in comments nor in notes)")
    else:
        src = sorted({s["source"] for s in shown})
        print(f"  text read from: {', '.join(src)}")
    print(f"  to generate: {len(p['todo'])} audio clips, ~{p['chars']} credits")


def build_deck(p, args, cfg, key, voice_id):
    prs, cache, slides = p["prs"], p["cache"], p["slides"]
    cache.mkdir(parents=True, exist_ok=True)
    texts = [s["text"] for s in slides]
    total = sum(1 for s in slides if s["text"])
    done, gen, failed, warn, streak, retimed = 0, 0, [], [], 0, []
    for idx, (slide, s) in enumerate(zip(prs.slides, slides)):
        if not s["sel"]:
            continue
        remove_previous_vo(slide)
        if not s["text"]:
            if not args.manual:
                silent_ms = int(cfg["silent_slide"] * 1000)
                nclk = count_click_groups(slide)
                if nclk:
                    retime_click_animations(slide, click_start_times(nclk, [], 0, silent_ms)[0])
                    retimed.append(s["n"])
                set_advance(slide, silent_ms)
            continue
        busy = "from cache..." if s["cached"] else "generating audio..."
        print(f"  [{int(100 * done / total):3d}%] slide {s['n']}: {busy}", end="", flush=True)
        raw, clean = cache / f"{s['key']}.raw.mp3", cache / f"{s['key']}.mp3"
        try:
            if not s["cached"]:
                for f in (raw, clean):
                    f.unlink(missing_ok=True)
                if args.fake:
                    tts_fake(s["text"], raw)
                else:
                    prev = next((t for t in reversed(texts[:idx]) if t), None)
                    nxt = next((t for t in texts[idx + 1:] if t), None)
                    tmp = raw.with_suffix(".part")
                    tmp.write_bytes(tts(s["text"], key, voice_id, cfg, prev, nxt))
                    tmp.replace(raw)               # only complete files go into the cache
                gen += 1
            if not audio_ok(clean):
                if not postprocess(raw, clean, cfg["loudness"]):
                    warn.append(f"slide {s['n']}: audio cleanup failed, used the original audio")
            dur = probe_duration(clean)
        except VOError as e:
            if e.fatal:
                print()
                raise
            print(f"\r  [ERR ] slide {s['n']}: {e}" + " " * 10)
            failed.append(s["n"])
            streak += 1
            if streak >= 3:
                warn.append("3 errors in a row: generation stopped, saving what is done")
                break
            continue
        except (subprocess.SubprocessError, OSError, ValueError) as e:
            print(f"\r  [ERR ] slide {s['n']}: audio could not be processed ({type(e).__name__})" + " " * 10)
            failed.append(s["n"])
            continue
        streak = 0
        lead = int(cfg["lead_in"] * 1000)
        spid = add_audio_shape(prs, slide, clean, not args.icon_inside)
        dur_ms = int(dur * 1000)
        group = add_autoplay(slide, spid, dur_ms, lead)
        s["adv"] = lead + dur_ms + int(cfg["tail"] * 1000)
        if not args.manual:
            nclk = count_click_groups(slide)
            if nclk:
                resumes = voice_resumes_ms(clean, max(0.5, cfg["pause_segment"] * 0.7))
                times, note = click_start_times(nclk, s["starts"], lead, dur_ms, resumes)
                if note:
                    warn.append(f"slide {s['n']}: {note}")
                retime_click_animations(slide, times, group)
                retimed.append(s["n"])
                # the last element to appear must also stay on screen for at least 'tail' seconds
                s["adv"] = max(s["adv"], times[-1] + int(cfg["tail"] * 1000))
            set_advance(slide, s["adv"])
        done += 1
        print(f"\r  [{int(100 * done / total):3d}%] slide {s['n']}: done ({dur:.1f} s)" + " " * 20, flush=True)

    out = p["out"]
    if not writable(out):
        alt = free_name(out)
        warn.append(f"{out.name} is open or locked: saved as {alt.name}")
        out = alt
    print("  saving...", end="", flush=True)
    try:
        prs.save(str(out))
    except PermissionError:
        alt = free_name(out)
        prs.save(str(alt))
        warn.append(f"{out.name} locked while saving: saved as {alt.name}")
        out = alt
    if args.manual:
        print(f"\r  saved {out.name}: {gen} new clips, advance on click" + " " * 10)
    else:
        tot = sum(s.get("adv", cfg["silent_slide"] * 1000) for s in slides if not s["hidden"]) / 1000
        print(f"\r  saved {out.name}: {gen} new clips, duration ~{tot / 60:.1f} min" + " " * 10)
    mute = [str(s["n"]) for s in slides if s["sel"] and not s["text"]]
    if mute:
        extra = "" if args.manual else f" (shown for {cfg['silent_slide']:.0f} s)"
        warn.insert(0, f"no 'VO:'{extra}: slide {', '.join(mute)}")
    if retimed:
        warn.append(f"on-click animations made automatic, synced with the voice: slide "
                    f"{', '.join(map(str, retimed))} (with --manual they stay on click)")
    if failed:
        warn.append(f"NOT generated: slide {', '.join(map(str, failed))}. Run the same command again to "
                    f"retry: the others are cached and will not be charged again.")
    for w in warn:
        print("   - " + w)
    if args.open:
        open_file(out)
    return not failed


# --------------------------------------------------------------------------- service commands

def cmd_voices(key):
    data = api("GET", "/v1/voices", key, retries=1).json().get("voices", [])
    print(f"{'NAME':<28} {'CATEGORY':<12} VOICE ID")
    for v in sorted(data, key=lambda v: (v.get("category", ""), v.get("name", ""))):
        print(f"{v.get('name', '')[:27]:<28} {v.get('category', ''):<12} {v.get('voice_id')}")
    print(f"\n{len(data)} voices. Copy the VOICE ID into ELEVENLABS_VOICE_ID (or the .env file).")


def cmd_credits(key):
    c = get_credits(key)
    if c is None:
        print("This key is not allowed to read the subscription (User: Read).")
    else:
        print(f"Credits left: {c[0]} of {c[1]}" + (f" · resets on {c[2]}" if c[2] else ""))


def cmd_check(key, voice_id):
    ok = True

    def line(good, label, detail="", info=False):
        nonlocal ok
        ok &= good
        print(f"  [{'OK' if good else 'NO'}] {label}{(' - ' + detail) if detail and (not good or info) else ''}")
    print(f"vo.py {VERSION} · Python {sys.version.split()[0]}")
    try:
        import requests
        line(bool(requests.__version__), "Python libraries")
    except ImportError:
        line(False, "Python libraries", "pip install requests")
    line(all(shutil.which(t) for t in ("ffmpeg", "ffprobe")), "ffmpeg and ffprobe", "winget install ffmpeg")
    line(bool(key), "ELEVENLABS_API_KEY", "" if key else "not set")
    line(bool(voice_id), "ELEVENLABS_VOICE_ID", "" if voice_id else "not set (python vo.py --voices)")
    if key:
        try:
            c = get_credits(key)
            line(True, "API key valid", "credits not readable with this key" if c is None
                 else f"{c[0]} credits left" + (f", resets on {c[2]}" if c[2] else ""), info=True)
        except VOError as e:
            line(False, "API key", str(e))
        if voice_id:
            try:
                name = get_voice_name(key, voice_id)
                line(True, "voice", name or "cannot be checked with this key", info=True)
            except VOError as e:
                line(False, "voice", str(e))
    print("All set." if ok else "Fix the [NO] lines and run --check again.")
    return ok


# --------------------------------------------------------------------------- main

HELP_EPILOG = """
examples:
  python vo.py course.pptx                    generate and embed the audio
  python vo.py course.pptx --preview          show texts and credits, spends nothing
  python vo.py *.pptx -y                      every deck in the folder, no confirmation
  python vo.py course.pptx --slides 2,5-7     only some slides (the others stay silent)
  python vo.py course.pptx --manual --open    advance on click, then open the result
  python vo.py --check                        check key, voice, ffmpeg and credits
  python vo.py --voices                       list voices with their IDs

in the slides:
  comment "VO: text to read"        one comment per slide, or several VO: comments (read in order)
  in the notes: VO: "text"          one or more VO: lines (used if the slide has no VO: comments);
                                    (directions) in parentheses and outer quotes are not read;
                                    VO: (click 2) "..." = the line starts with the 2nd animation click
  [pause: 0.8]                      pause inside the text, in seconds (max 3)

configuration (environment variables or a .env file next to the script):
  ELEVENLABS_API_KEY=...            required
  ELEVENLABS_VOICE_ID=...           required (or --voice)
"""


def build_parser():
    ap = argparse.ArgumentParser(
        prog="vo.py", formatter_class=argparse.RawDescriptionHelpFormatter, epilog=HELP_EPILOG,
        description="ElevenLabs voiceover inside PowerPoint decks: reads the 'VO:' lines of each slide, "
                    "generates the audio and embeds it with autoplay.")
    ap.add_argument("deck", nargs="*", help="one or more .pptx files (accepts *.pptx)")

    def opt(group, names, legacy=(), **kw):
        """Adds an option; 'legacy' are the old Italian names, still accepted but hidden."""
        group.add_argument(*names, **kw)
        if legacy:
            hidden = dict(kw, help=argparse.SUPPRESS, default=argparse.SUPPRESS)
            hidden.pop("metavar", None)
            group.add_argument(*legacy, dest=kw.get("dest") or names[-1].lstrip("-").replace("-", "_"),
                               **hidden)

    g = ap.add_argument_group("generation")
    opt(g, ["--preview"], ["--anteprima"], action="store_true", help="show texts and credits needed, then exit")
    opt(g, ["-y", "--yes"], ["--si"], action="store_true", help="do not ask for confirmation")
    opt(g, ["--slides"], ["--slide"], metavar="N", help="only these slides, e.g. 3 or 1,4-6")
    opt(g, ["--regenerate"], ["--rigenera"], action="store_true", help="ignore the cache and regenerate (costs credits)")
    opt(g, ["--fake"], action="store_true", help="fake audio (tones), no credits: for testing")
    g = ap.add_argument_group("voice")
    opt(g, ["--voice"], ["--voce"], metavar="ID", help="voice ID (default: ELEVENLABS_VOICE_ID)")
    opt(g, ["--model"], ["--modello"], metavar="ID", help=f"model (default: {CFG['model_id']})")
    opt(g, ["--speed"], ["--velocita"], type=float, metavar="X",
        help=f"voice speed, 0.7-1.2 (default {CFG['voice_settings']['speed']})")
    opt(g, ["--stability"], ["--stabilita"], type=float, metavar="X",
        help=f"0-1: lower = more expressive (default {CFG['voice_settings']['stability']})")
    g = ap.add_argument_group("slideshow")
    opt(g, ["--manual"], ["--manuale"], action="store_true", help="slides advance on click only (audio still autoplays)")
    opt(g, ["--lead-in"], ["--pausa-iniziale"], type=float, metavar="S",
        help=f"seconds before the voice starts (default {CFG['lead_in']})")
    opt(g, ["--tail"], ["--pausa-finale"], type=float, metavar="S",
        help=f"seconds after the voice ends (default {CFG['tail']})")
    opt(g, ["--line-pause"], ["--pausa-righe"], type=float, metavar="S",
        help=f"pause between two VO: lines on the same slide, max 3 (default {CFG['pause_segment']})")
    opt(g, ["--silent-slide"], ["--slide-muta"], type=float, metavar="S",
        help=f"duration of slides without VO: (default {CFG['silent_slide']})")
    opt(g, ["--icon-inside"], ["--icona-dentro"], action="store_true",
        help="put the audio icon in the bottom-right corner of the slide (default: off-slide, invisible)")
    g = ap.add_argument_group("files")
    opt(g, ["-o", "--output"], metavar="FILE", help="output file (single deck only)")
    opt(g, ["--cache"], metavar="DIR", help="audio cache folder (default: vo_cache next to the deck)")
    opt(g, ["--open"], ["--apri"], action="store_true", help="open the result when done")
    g = ap.add_argument_group("account and diagnostics")
    opt(g, ["--check"], ["--verifica"], action="store_true", help="check ffmpeg, key, voice and credits, then exit")
    opt(g, ["--credits"], ["--crediti"], action="store_true", help="show remaining credits, then exit")
    opt(g, ["--voices"], ["--voci"], action="store_true", help="list available voices with their IDs, then exit")
    g.add_argument("--version", action="version", version=f"vo.py {VERSION}")

    def error(msg):
        ap.print_usage(sys.stderr)
        sys.exit(f"ERROR: {msg}\nRun python vo.py --help for all options.")
    ap.error = error
    return ap


def main():
    load_dotenv()
    ap = build_parser()
    args = ap.parse_args()
    key = os.environ.get("ELEVENLABS_API_KEY", "").strip()
    voice_id = (args.voice or os.environ.get("ELEVENLABS_VOICE_ID", "")).strip()

    if args.check:
        return 0 if cmd_check(key, voice_id) else 1
    if args.voices or args.credits:
        if not key:
            raise VOError("ELEVENLABS_API_KEY is not set.", "Set it in the terminal or in the .env file.")
        cmd_voices(key) if args.voices else cmd_credits(key)
        return 0
    if not args.deck:
        ap.print_help()
        return 1

    cfg = json.loads(json.dumps(CFG))
    if args.model:
        cfg["model_id"] = args.model
    if args.speed is not None:
        if not 0.7 <= args.speed <= 1.2:
            raise VOError("--speed must be between 0.7 and 1.2.")
        cfg["voice_settings"]["speed"] = args.speed
    if args.stability is not None:
        if not 0 <= args.stability <= 1:
            raise VOError("--stability must be between 0 and 1.")
        cfg["voice_settings"]["stability"] = args.stability
    for opt_name, k in (("lead_in", "lead_in"), ("tail", "tail"), ("silent_slide", "silent_slide"),
                        ("line_pause", "pause_segment")):
        v = getattr(args, opt_name)
        if v is not None:
            if v < 0:
                raise VOError(f"--{opt_name.replace('_', '-')} cannot be negative.")
            cfg[k] = v

    check_ffmpeg()
    decks = expand_decks(args.deck)
    if args.output and len(decks) > 1:
        raise VOError("-o/--output works with a single deck only.", "Without -o each deck becomes <name>_vo.pptx.")
    if not args.fake and not args.preview:
        if not key:
            raise VOError("ELEVENLABS_API_KEY is not set.",
                          'PowerShell: $env:ELEVENLABS_API_KEY="..."  or put it in the .env file')
        if not voice_id:
            raise VOError("No voice set.", "Set ELEVENLABS_VOICE_ID or use --voice ID (list: python vo.py --voices)")

    # 1. plan every deck before spending any credit
    plans, errors = [], []
    for d in decks:
        try:
            plans.append(plan_deck(d, args, cfg, voice_id))
            print_plan(plans[-1])
        except VOError as e:
            print(f"\n{d.name}\n  ERROR: {e}" + (f"\n  Fix: {e.hint}" if e.hint else ""))
            errors.append(d.name)
    plans = [p for p in plans if any(s["text"] for s in p["slides"])]
    need = sum(p["chars"] for p in plans)
    if len(plans) > 1:
        print(f"\nTotal: {sum(len(p['todo']) for p in plans)} audio clips to generate, ~{need} credits")
    if args.preview or not plans:
        return 1 if errors else 0

    # 2. account checks
    if need and not args.fake:
        name = get_voice_name(key, voice_id)
        if name:
            print(f"Voice: {name}")
        c = get_credits(key)
        if c is None:
            print("Credits: not readable with this key, check skipped")
        else:
            left, _, reset = c
            print(f"Credits left: {left}")
            if need > left:
                raise VOError(f"About {need} credits needed, you have {left}.",
                              (f"Credits reset on {reset}. " if reset else "") +
                              "You can generate part of it with --slides, or upgrade your plan.")

    # 3. confirmation
    if need and not args.fake and not args.yes:
        try:
            ans = input(f"Generate {sum(len(p['todo']) for p in plans)} audio clips (~{need} credits)? [y/N] ")
        except EOFError:
            ans = ""
        if ans.strip().lower() not in YES:
            print("Cancelled, no credits used.")
            return 0

    # 4. generation
    ok = True
    for p in plans:
        print(f"\n{p['path'].name}")
        try:
            ok &= build_deck(p, args, cfg, key, voice_id)
        except VOError as e:
            if e.fatal:
                raise
            print(f"  ERROR: {e}" + (f"\n  Fix: {e.hint}" if e.hint else ""))
            ok = False
    return 0 if ok and not errors else 1


if __name__ == "__main__":
    try:
        sys.exit(main())
    except VOError as e:
        print(f"\nERROR: {e}" + (f"\nFix: {e.hint}" if e.hint else ""), file=sys.stderr)
        sys.exit(1)
    except KeyboardInterrupt:
        print("\nInterrupted. Audio already generated is in vo_cache: re-running will not charge for it again.",
              file=sys.stderr)
        sys.exit(130)
