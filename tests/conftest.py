"""Shared fixtures: a fake ElevenLabs server and a few test decks built on the fly.

No test calls the real API or spends credits.
"""
import json
import os
import re
import subprocess
import sys
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

import pytest
from pptx import Presentation
from pptx.opc.package import Part
from pptx.opc.packuri import PackURI
from pptx.oxml import parse_xml
from pptx.util import Inches

ROOT = Path(__file__).resolve().parent.parent
VO = ROOT / "vo.py"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P188 = "http://schemas.microsoft.com/office/powerpoint/2018/8/main"
GOOD_KEY = "test-key"


# --------------------------------------------------------------------------- fake ElevenLabs

class Mock:
    """State of the fake server: what to answer and what it received."""

    def __init__(self, mp3):
        self.mp3 = mp3
        self.reset()

    def reset(self):
        self.tts = "ok"              # ok | quota | 500 | flaky | empty | fail:<text>
        self.credits_left = 5000
        self.sub_perm = True
        self.voice = "ok"            # ok | 404
        self.tts_calls = []          # JSON bodies of the synthesis requests
        self._n = 0


def make_handler(mock):
    class H(BaseHTTPRequestHandler):
        def log_message(self, *a):
            pass

        def _send(self, code, body, ctype="application/json"):
            b = body if isinstance(body, bytes) else json.dumps(body).encode()
            self.send_response(code)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(b)))
            self.end_headers()
            self.wfile.write(b)

        def _err(self, code, status, msg):
            self._send(code, {"detail": {"status": status, "code": status, "message": msg}})

        def do_GET(self):
            if self.headers.get("xi-api-key") != GOOD_KEY:
                return self._err(401, "invalid_api_key", "Invalid API key")
            if self.path.startswith("/v1/user/subscription"):
                if not mock.sub_perm:
                    return self._err(401, "missing_permissions", "missing user_read")
                return self._send(200, {"character_count": 10000 - mock.credits_left, "character_limit": 10000,
                                        "next_character_count_reset_unix": 1767225600})
            if re.match(r"/v1/voices/\w+", self.path):
                if mock.voice == "404":
                    return self._err(404, "voice_not_found", "A voice with the voice_id was not found.")
                return self._send(200, {"name": "Test voice"})
            if self.path.startswith("/v1/voices"):
                return self._send(200, {"voices": [{"name": "Test voice", "category": "premade",
                                                    "voice_id": "voice123"}]})
            self._send(404, {"detail": "not found"})

        def do_POST(self):
            body = json.loads(self.rfile.read(int(self.headers.get("Content-Length", 0))))
            if self.headers.get("xi-api-key") != GOOD_KEY:
                return self._err(401, "invalid_api_key", "Invalid API key")
            mock.tts_calls.append(body)
            mock._n += 1
            mode = mock.tts
            if mode == "quota":
                return self._err(401, "quota_exceeded",
                                 "This request exceeds your quota of 10000. You have 27 credits remaining.")
            if mode == "500":
                return self._err(500, "internal", "boom")
            if mode == "flaky" and mock._n % 3:
                return self._err(429, "too_many_concurrent_requests", "busy")
            if mode == "empty":
                return self._send(200, b"", "audio/mpeg")
            if mode.startswith("fail:") and mode[5:] in body["text"]:
                return self._err(500, "internal", "boom")
            self._send(200, mock.mp3, "audio/mpeg")
    return H


@pytest.fixture(scope="session")
def _server(tmp_path_factory):
    d = tmp_path_factory.mktemp("mock")
    # two phrases separated by 1 s of silence: used by the animation sync test
    subprocess.run(["ffmpeg", "-v", "error", "-y", "-f", "lavfi", "-i", "sine=frequency=300:duration=1.5",
                    "-f", "lavfi", "-i", "anullsrc=r=44100:cl=mono", "-f", "lavfi", "-i",
                    "sine=frequency=300:duration=1.5", "-filter_complex",
                    "[1:a]atrim=0:1[s];[0:a][s][2:a]concat=n=3:v=0:a=1", "-c:a", "libmp3lame",
                    str(d / "voice.mp3")], check=True)
    mock = Mock((d / "voice.mp3").read_bytes())
    srv = ThreadingHTTPServer(("127.0.0.1", 0), make_handler(mock))
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    yield mock, f"http://127.0.0.1:{srv.server_address[1]}"
    srv.shutdown()


@pytest.fixture
def api(_server):
    mock, _ = _server
    mock.reset()
    return mock


@pytest.fixture
def run(_server, tmp_path):
    """Runs vo.py in the test folder, against the fake server, with no wait between retries."""
    _, base = _server

    def _run(*args, key=GOOD_KEY, voice="voice123", stdin=None):
        env = {k: v for k, v in os.environ.items() if not k.startswith("ELEVENLABS_")}
        env.update(ELEVENLABS_API_BASE=base, VO_RETRY_WAIT="0.01", PYTHONIOENCODING="utf-8")
        if key:
            env["ELEVENLABS_API_KEY"] = key
        if voice:
            env["ELEVENLABS_VOICE_ID"] = voice
        p = subprocess.run([sys.executable, str(VO), *map(str, args)], cwd=tmp_path, env=env,
                           input=stdin, capture_output=True, text=True, timeout=120)
        p.out = (p.stdout + p.stderr).replace("\r", "\n")
        return p
    return _run


# --------------------------------------------------------------------------- test decks

def _new_deck():
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    return prs


def _slide(prs, title):
    s = prs.slides.add_slide(prs.slide_layouts[1])
    s.shapes.title.text = title
    s.placeholders[1].text = "Slide body"
    return s


def add_legacy_comments(prs, slide, texts, author_id=0):
    pkg = prs.part.package
    if not getattr(prs, "_authors", None):
        authors = "".join(f'<p:cmAuthor id="{i}" name="Author {i}" initials="A{i}" lastIdx="9" clrIdx="{i}"/>'
                          for i in range(3))
        part = Part(PackURI("/ppt/commentAuthors.xml"),
                    "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml", pkg,
                    f'<p:cmAuthorLst xmlns:p="{P}">{authors}</p:cmAuthorLst>'.encode())
        prs.part.relate_to(part, f"{R}/commentAuthors")
        prs._authors = True
    n = len([p for p in pkg.iter_parts() if str(p.partname).startswith("/ppt/comments/")]) + 1
    cms = "".join(f'<p:cm authorId="{i}" dt="2026-11-09T10:0{i}:00.000" idx="{n}"><p:pos x="0" y="0"/>'
                  f'<p:text>{t}</p:text></p:cm>' for i, t in enumerate(texts))
    part = Part(PackURI(f"/ppt/comments/comment{n}.xml"),
                "application/vnd.openxmlformats-officedocument.presentationml.comments+xml", pkg,
                f'<p:cmLst xmlns:p="{P}">{cms}</p:cmLst>'.encode())
    slide.part.relate_to(part, f"{R}/comments")


def add_modern_comment(prs, slide, paragraphs, reply):
    pkg = prs.part.package
    au = "{11111111-1111-1111-1111-111111111111}"
    authors = Part(PackURI("/ppt/authors.xml"), "application/vnd.ms-powerpoint.authors+xml", pkg,
                   f'<p188:authorLst xmlns:a="{A}" xmlns:r="{R}" xmlns:p188="{P188}"><p188:author id="{au}" '
                   f'name="Test" initials="T" userId="Test" providerId="None"/></p188:authorLst>'.encode())
    prs.part.relate_to(authors, "http://schemas.microsoft.com/office/2018/10/relationships/authors")
    ps = "".join(f"<a:p><a:r><a:t>{t}</a:t></a:r></a:p>" for t in paragraphs)
    xml = (f'<p188:cmLst xmlns:a="{A}" xmlns:r="{R}" xmlns:p188="{P188}">'
           f'<p188:cm id="{{22222222-2222-2222-2222-222222222222}}" authorId="{au}" created="2026-11-09T10:00:00.000">'
           f'<p188:replyLst><p188:reply id="{{33333333-3333-3333-3333-333333333333}}" authorId="{au}" '
           f'created="2026-11-09T10:05:00.000"><p188:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{reply}'
           f'</a:t></a:r></a:p></p188:txBody></p188:reply></p188:replyLst>'
           f'<p188:txBody><a:bodyPr/><a:lstStyle/>{ps}</p188:txBody></p188:cm></p188:cmLst>')
    part = Part(PackURI("/ppt/comments/modernComment_100_0.xml"), "application/vnd.ms-powerpoint.comments+xml",
                pkg, xml.encode())
    slide.part.relate_to(part, "http://schemas.microsoft.com/office/2018/10/relationships/comments")


def add_click_animations(slide, n):
    """n on-click animations (fade in) on the slide body."""
    sp = slide.placeholders[1].shape_id
    groups, cid = [], 3
    for _ in range(n):
        groups.append(f'''<p:par><p:cTn id="{cid}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>
<p:par><p:cTn id="{cid + 1}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
<p:par><p:cTn id="{cid + 2}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect">
<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
<p:set><p:cBhvr><p:cTn id="{cid + 3}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="{sp}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>
</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>''')
        cid += 4
    slide._element.append(parse_xml(f'''<p:timing xmlns:p="{P}"><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>
<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>{"".join(groups)}</p:childTnLst></p:cTn>
<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>
<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>
</p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst><p:bldP spid="{sp}" grpId="0" build="p"/></p:bldLst></p:timing>'''))


@pytest.fixture
def deck_comments(tmp_path):
    """4 slides: legacy VO: comment + review comment, modern comment with a reply,
    slide without text, slide with two VO: comments."""
    prs = _new_deck()
    s1, s2, _, s4 = (_slide(prs, f"Slide {i}") for i in range(1, 5))
    add_legacy_comments(prs, s1, ["Fix the title", "VO: Welcome to the course."])
    add_modern_comment(prs, s2, ["VO: Second slide,", "on two paragraphs."], "REPLY TO IGNORE")
    add_legacy_comments(prs, s4, ["VO: First part.", "VO: Second part."])
    path = tmp_path / "comments.pptx"
    prs.save(path)
    return path


@pytest.fixture
def deck_notes(tmp_path):
    """Text in the notes, with stage directions, quotes and (click N); 2 on-click animations."""
    prs = _new_deck()
    s1 = _slide(prs, "Notes")
    s1.notes_slide.notes_text_frame.text = (
        "Notes for the presenter, not to be read\n"
        'VO: (minute 0) "Hello everyone." (firm tone)\n'
        'VO: (click 1) "Here is the first point."\n'
        'VO: (clic 2) "And here is the second."')     # the Italian "clic" is accepted too
    add_click_animations(s1, 2)
    s2 = _slide(prs, "Animations only")
    add_click_animations(s2, 1)
    path = tmp_path / "notes.pptx"
    prs.save(path)
    return path
