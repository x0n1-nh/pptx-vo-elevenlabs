"""Records docs/demo.gif: a real run of vo.py on a demo deck.

The script runs in a virtual terminal (pty) and its output is rendered frame by frame.
ElevenLabs responses come from a local fake server: no credits, no real voice. Fake clips last
as long as speech at ~14 characters per second, and each request waits ~1.5 s like a short
real synthesis.

Usage (Linux/macOS, needs ffmpeg):  pip install pyte Pillow && python tools/record_demo.py
"""
import json, os, pty, select, subprocess, sys, tempfile, threading, time
from http.server import ThreadingHTTPServer
from pathlib import Path

import pyte
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests"))
import conftest  # noqa: E402  (test decks and fake server)

COLS, ROWS = 112, 26
FONT = "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf"
BG, FG, DIM, ACCENT = (12, 12, 12), (204, 204, 204), (110, 110, 110), (97, 214, 214)
PROMPT = "PS C:\\course> "
COMMAND = "python vo.py budgeting.pptx"


def build_deck(path):
    prs = conftest._new_deck()
    s = [conftest._slide(prs, f"Slide {i}") for i in range(1, 6)]
    conftest.add_legacy_comments(prs, s[0], [
        "VO: Welcome. In this lesson we will set up a monthly budget in three simple steps."])
    conftest.add_legacy_comments(prs, s[1], [
        "Check the chart",
        "VO: Step one is knowing your income: start from your take-home pay, not your gross salary."])
    s[3].notes_slide.notes_text_frame.text = (
        'VO: (click 1) "Next, split your expenses into fixed and variable."\n'
        'VO: (click 2) "Fixed costs are paid anyway: savings come from the variable ones."')
    conftest.add_click_animations(s[3], 2)
    conftest.add_legacy_comments(prs, s[4], [
        "VO: Last step: set a savings goal and review it every month. See you in the next lesson."])
    prs.save(path)


def start_server(tmp):
    mock = conftest.Mock(b"")
    mock.credits_left = 28450
    base = conftest.make_handler(mock)

    class H(base):
        def do_GET(self):
            if self.path.startswith("/v1/voices/"):
                return self._send(200, {"name": "Demo voice"})
            return super().do_GET()

        def do_POST(self):
            n = int(self.headers.get("Content-Length", 0))
            body = json.loads(self.rfile.read(n))
            text = body["text"]
            dur = max(1.5, len(text) / 14)
            out = Path(tmp, f"tts_{abs(hash(text))}.mp3")
            subprocess.run(["ffmpeg", "-v", "error", "-y", "-f", "lavfi", "-i",
                            f"sine=frequency=300:duration={dur:.2f}", "-c:a", "libmp3lame", str(out)],
                           check=True)
            time.sleep(1.5)
            self._send(200, out.read_bytes(), "audio/mpeg")

    srv = ThreadingHTTPServer(("127.0.0.1", 0), H)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    return f"http://127.0.0.1:{srv.server_address[1]}"


def record(workdir, base):
    """Runs vo.py in a pty and returns [(seconds, bytes)]. Answers 'y' to the confirmation."""
    env = {k: v for k, v in os.environ.items() if not k.startswith("ELEVENLABS_")}
    env.update(ELEVENLABS_API_BASE=base, ELEVENLABS_API_KEY=conftest.GOOD_KEY,
               ELEVENLABS_VOICE_ID="voice123", PYTHONIOENCODING="utf-8", COLUMNS=str(COLS), TERM="xterm")
    pid, fd = pty.fork()
    if pid == 0:
        os.chdir(workdir)
        os.execve(sys.executable, [sys.executable, str(ROOT / "vo.py"), "budgeting.pptx"], env)
    events, t0, buf, answered = [], time.time(), b"", False
    while True:
        r, _, _ = select.select([fd], [], [], 0.05)
        if r:
            try:
                data = os.read(fd, 4096)
            except OSError:
                break
            if not data:
                break
            events.append((time.time() - t0, data))
            buf += data
        if not answered and b"[y/N]" in buf:
            time.sleep(1.2)                       # the user reads and answers
            os.write(fd, b"y\r")
            answered = True
        if os.waitpid(pid, os.WNOHANG)[0]:
            try:
                while True:
                    data = os.read(fd, 4096)
                    if not data:
                        break
                    events.append((time.time() - t0, data))
            except OSError:
                pass
            break
    return events


def render(events, out_gif):
    font = ImageFont.truetype(FONT, 16)
    cw, ch = font.getlength("M"), 22
    pad, bar = 18, 30
    W, H = int(COLS * cw + 2 * pad), ROWS * ch + 2 * pad + bar
    screen = pyte.Screen(COLS, ROWS)
    stream = pyte.ByteStream(screen)

    # first the command typed by hand, then the real output
    timeline, t = [(0.0, PROMPT.encode())], 0.8
    for c in COMMAND:
        timeline.append((t, c.encode()))
        t += 0.055
    t += 0.5
    timeline.append((t, b"\r\n"))
    timeline += [(t + 0.3 + et, data) for et, data in events]
    end = timeline[-1][0]
    timeline.append((end + 0.01, b"\r\n" + PROMPT.encode()))

    def frame(cursor_on):
        im = Image.new("RGB", (W, H), BG)
        d = ImageDraw.Draw(im)
        d.rectangle([0, 0, W, bar], fill=(32, 32, 32))
        d.text((pad, bar / 2), "Windows PowerShell", font=font, fill=DIM, anchor="lm")
        for i, sym in enumerate(["\u2715", "\u25a1", "\u2014"]):        # — □ ✕
            d.text((W - pad - 8 - i * 34, bar / 2), sym, font=font, fill=DIM, anchor="mm")
        for y, line in enumerate(screen.display):
            yy = bar + pad + y * ch
            if line.startswith("PS C:"):
                d.text((pad, yy), PROMPT, font=font, fill=ACCENT)
                d.text((pad + len(PROMPT) * cw, yy), line[len(PROMPT):], font=font, fill=FG)
            else:
                d.text((pad, yy), line, font=font, fill=FG)
        if cursor_on and not screen.cursor.hidden:
            x, y = pad + screen.cursor.x * cw, bar + pad + screen.cursor.y * ch
            d.rectangle([x, y + 2, x + cw - 1, y + ch - 3], fill=FG)
        return im

    frames, durations = [], []
    for i, (t, data) in enumerate(timeline):
        stream.feed(data)
        nxt = timeline[i + 1][0] if i + 1 < len(timeline) else t + 5.0
        dt = max(nxt - t, 0.02)
        if frames and dt < 0.04 and i + 1 < len(timeline):
            continue                                   # bursts of output: a single frame
        frames.append(frame(True))
        durations.append(int(dt * 1000))
    durations[-1] = 5000                                # final pause before looping

    pal = [f.convert("P", palette=Image.ADAPTIVE, colors=32) for f in frames]
    pal[0].save(out_gif, save_all=True, append_images=pal[1:], duration=durations, loop=0, optimize=True)
    return len(frames), sum(durations) / 1000, W, H


def main():
    out = ROOT / "docs" / "demo.gif"
    out.parent.mkdir(exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        work = Path(tmp, "course")
        work.mkdir()
        build_deck(work / "budgeting.pptx")
        events = record(work, start_server(tmp))
        n, secs, W, H = render(events, out)
    print(f"{out} · {n} frames · {secs:.1f} s · {W}x{H} · {out.stat().st_size / 1024:.0f} KB")


if __name__ == "__main__":
    main()
