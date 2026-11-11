"""End-to-end tests of vo.py against a fake ElevenLabs server (no credits spent)."""
import importlib.util
import re
import zipfile

import pytest

from conftest import VO


def load_module():
    spec = importlib.util.spec_from_file_location("vo", VO)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def slide_xml(pptx, n):
    return zipfile.ZipFile(pptx).read(f"ppt/slides/slide{n}.xml").decode()


def adv_ms(xml):
    m = re.search(r'advTm="(\d+)"', xml)
    return int(m.group(1)) if m else None


# --------------------------------------------------------------------------- reading the text

@pytest.mark.parametrize("raw, text, click", [
    ('VO: "Hello everyone."', "Hello everyone.", None),
    ('VO: (minute 0) "Hello." (serious tone)', "Hello.", None),
    ("VO: (click 2) Second point", "Second point", 2),
    ("VO: (clic 3) Italian marker still works", "Italian marker still works", 3),
    ("vo:   text   with   spaces", "text with spaces", None),
    ("VO: «Guillemets»", "Guillemets", None),
])
def test_clean_segment(raw, text, click):
    assert load_module()._clean_segment(raw) == (text, click)


def test_pause_marker():
    mod = load_module()
    for marker in ("[pause: 0.8]", "[pausa: 0,8]"):
        assert mod._PAUSE_RE.sub("X", f"a {marker} b") == "a X b"


def test_preview_makes_no_api_calls(run, api, deck_comments):
    p = run(deck_comments.name, "--preview")
    assert p.returncode == 0, p.out
    assert "Welcome to the course." in p.out
    assert "Fix the title" not in p.out              # comment without VO: is ignored
    assert "REPLY" not in p.out                      # replies are ignored
    assert api.tts_calls == []


def test_notes_directions_and_quotes_not_read(run, api, deck_notes):
    p = run(deck_notes.name, "-y")
    assert p.returncode == 0, p.out
    sent = api.tts_calls[0]["text"]
    assert sent.count('<break time="1.0s" />') == 2      # pause between VO: lines
    sent = re.sub(r"<break[^>]*>", "", sent)
    assert "Hello everyone." in sent and "And here is the second." in sent
    assert "minute" not in sent and "tone" not in sent and "click" not in sent and '"' not in sent
    assert "presenter" not in sent


# --------------------------------------------------------------------------- generation

def test_generates_and_embeds(run, api, deck_comments, tmp_path):
    p = run(deck_comments.name, "-y")
    assert p.returncode == 0, p.out
    out = tmp_path / "comments_vo.pptx"
    assert out.exists()
    assert len(api.tts_calls) == 3                    # slides 1, 2, 4 (one request per slide)
    for n in (1, 2, 4):
        x = slide_xml(out, n)
        assert x.count("playFrom") == 1 and 'name="VO-auto' in x
        assert adv_ms(x) > 0
    assert "playFrom" not in slide_xml(out, 3)
    assert adv_ms(slide_xml(out, 3)) == 4000          # silent slide: silent_slide
    media = [n for n in zipfile.ZipFile(out).namelist() if n.startswith("ppt/media/") and n.endswith(".mp3")]
    assert len(media) == 3


def test_context_between_slides(run, api, deck_comments):
    run(deck_comments.name, "-y")
    second = api.tts_calls[1]
    assert second["previous_text"].startswith("Welcome")
    assert "next_text" in second


def test_cache_and_idempotent_rerun(run, api, deck_comments, tmp_path):
    assert run(deck_comments.name, "-y").returncode == 0
    first = slide_xml(tmp_path / "comments_vo.pptx", 4)
    api.tts_calls.clear()
    p = run(deck_comments.name, "-y")
    assert p.returncode == 0 and api.tts_calls == [], p.out   # all from cache, zero credits
    assert slide_xml(tmp_path / "comments_vo.pptx", 4).count("playFrom") == 1
    assert adv_ms(slide_xml(tmp_path / "comments_vo.pptx", 4)) == adv_ms(first)


def test_rerun_on_processed_deck_does_not_duplicate(run, api, deck_comments, tmp_path):
    run(deck_comments.name, "-y")
    p = run("comments_vo.pptx", "-y", "-o", "again.pptx")
    assert p.returncode == 0, p.out
    out = tmp_path / "again.pptx"
    for n in (1, 2, 4):
        assert slide_xml(out, n).count("playFrom") == 1
    media = [n for n in zipfile.ZipFile(out).namelist() if n.endswith(".mp3")]
    assert len(media) == 3


def test_manual_no_auto_advance(run, api, deck_comments, tmp_path):
    assert run(deck_comments.name, "-y", "--manual").returncode == 0
    for n in range(1, 5):
        assert adv_ms(slide_xml(tmp_path / "comments_vo.pptx", n)) is None
    assert slide_xml(tmp_path / "comments_vo.pptx", 1).count("playFrom") == 1


def test_click_animations_synced(run, api, deck_notes, tmp_path):
    p = run(deck_notes.name, "-y")
    assert p.returncode == 0, p.out
    x = slide_xml(tmp_path / "notes_vo.pptx", 1)
    assert "clickEffect" not in x                     # no clicks left
    delays = [int(d) for d in re.findall(r'nodeType="withEffect">\s*<p:stCondLst><p:cond delay="(\d+)"', x)]
    assert len(delays) == 2 and delays[0] < delays[1]
    assert adv_ms(x) >= delays[-1]
    # slide without voice but with animations: made automatic within the silent-slide duration
    x2 = slide_xml(tmp_path / "notes_vo.pptx", 2)
    assert "clickEffect" not in x2 and adv_ms(x2) == 4000


def test_manual_keeps_click_animations(run, api, deck_notes, tmp_path):
    assert run(deck_notes.name, "-y", "--manual").returncode == 0
    assert slide_xml(tmp_path / "notes_vo.pptx", 1).count("clickEffect") == 2


def test_fake_needs_no_key(run, api, deck_comments, tmp_path):
    p = run(deck_comments.name, "--fake", key=None, voice=None)
    assert p.returncode == 0, p.out
    assert (tmp_path / "comments_vo.pptx").exists() and api.tts_calls == []


def test_legacy_italian_options_still_work(run, api, deck_comments, tmp_path):
    p = run(deck_comments.name, "--si", "--manuale", "--slide", "1", "--velocita", "1.0")
    assert p.returncode == 0, p.out
    assert adv_ms(slide_xml(tmp_path / "comments_vo.pptx", 1)) is None
    assert len(api.tts_calls) == 1 and api.tts_calls[0]["voice_settings"]["speed"] == 1.0


# --------------------------------------------------------------------------- credits and confirmation

def test_confirmation_declined(run, api, deck_comments):
    p = run(deck_comments.name, stdin="n\n")
    assert p.returncode == 0 and "Cancelled" in p.out
    assert api.tts_calls == []


def test_not_enough_credits_before_generating(run, api, deck_comments):
    api.credits_left = 10
    p = run(deck_comments.name, "-y")
    assert p.returncode == 1 and "you have 10" in p.out
    assert api.tts_calls == []


def test_key_without_read_permission(run, api, deck_comments):
    api.sub_perm = False
    p = run(deck_comments.name, "-y")
    assert p.returncode == 0 and "check skipped" in p.out


# --------------------------------------------------------------------------- errors

def test_quota_exceeded_midway(run, api, deck_comments):
    api.tts = "quota"
    p = run(deck_comments.name, "-y")
    assert p.returncode == 1 and "Out of ElevenLabs credits" in p.out and "Fix:" in p.out


def test_retries_on_429(run, api, deck_comments, tmp_path):
    api.tts = "flaky"
    p = run(deck_comments.name, "-y")
    assert p.returncode == 0 and "retrying" in p.out, p.out
    assert (tmp_path / "comments_vo.pptx").exists()


def test_one_slide_fails_then_recovers(run, api, deck_comments, tmp_path):
    api.tts = "fail:Second slide"
    p = run(deck_comments.name, "-y")
    assert p.returncode == 1 and "NOT generated: slide 2" in p.out, p.out
    assert slide_xml(tmp_path / "comments_vo.pptx", 1).count("playFrom") == 1
    api.tts, api.tts_calls = "ok", []
    p = run(deck_comments.name, "-y")
    assert p.returncode == 0 and len(api.tts_calls) == 1  # only the missing slide is generated


def test_empty_response(run, api, deck_comments):
    api.tts = "empty"
    p = run(deck_comments.name, "-y", "--slides", "1")
    assert p.returncode == 1 and "empty" in p.out


def test_invalid_key(run, api, deck_comments):
    p = run(deck_comments.name, "-y", key="wrong")
    assert p.returncode == 1 and "Invalid or revoked API key" in p.out


def test_unknown_voice(run, api, deck_comments):
    api.voice = "404"
    p = run(deck_comments.name, "-y")
    assert p.returncode == 1 and "voice does not exist" in p.out


@pytest.mark.parametrize("name, content, expected", [
    ("missing.pptx", None, "File not found"),
    ("old.ppt", b"x", "old .ppt format"),
    ("broken.pptx", b"not a zip file", "does not open as a valid presentation"),
])
def test_invalid_files(run, api, tmp_path, name, content, expected):
    if content is not None:
        (tmp_path / name).write_bytes(content)
    p = run(name, "--preview")
    assert p.returncode == 1 and expected in p.out


def test_locked_output_saved_under_another_name(run, api, deck_comments, tmp_path):
    (tmp_path / "comments_vo.pptx").mkdir()           # simulates a file that cannot be written
    p = run(deck_comments.name, "-y")
    assert p.returncode == 0 and (tmp_path / "comments_vo (2).pptx").exists(), p.out


@pytest.mark.parametrize("args, expected", [
    (["--speed", "2"], "between 0.7 and 1.2"),
    (["--slides", "2-x"], "not valid"),
    (["--whatever"], "unrecognized arguments"),
])
def test_invalid_options(run, api, deck_comments, args, expected):
    p = run(deck_comments.name, *args)
    assert p.returncode != 0 and expected in p.out


def test_check(run, api):
    p = run("--check")
    assert p.returncode == 0 and "All set" in p.out
    assert run("--check", key="wrong").returncode == 1


def test_help(run):
    p = run("--help", key=None, voice=None)
    assert p.returncode == 0 and "--preview" in p.out and "examples:" in p.out
    assert "--anteprima" not in p.out                 # legacy names stay hidden


@pytest.mark.parametrize("base", ["https://evil.example.com", "http://127.0.0.1@evil.example.com",
                                  "http://localhost.evil.example.com"])
def test_api_override_only_for_localhost(run, api, deck_comments, base):
    p = run(deck_comments.name, "-y", env_extra={"ELEVENLABS_API_BASE": base})
    assert p.returncode == 1 and "refused" in p.out
    assert api.tts_calls == []
