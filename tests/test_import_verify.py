"""Tests for --import (dictated transcript -> notes) and --verify (whisper check of the audio)."""
import json
import re

from pptx import Presentation

from conftest import TRANSCRIPT
from test_vo import load_module


def notes_of(path):
    return [s.notes_slide.notes_text_frame.text if s.has_notes_slide else "" for s in Presentation(path).slides]


# --------------------------------------------------------------------------- transcript parsing

def test_parse_voice2notes_transcript(tmp_path):
    f = tmp_path / "t.txt"
    f.write_text(TRANSCRIPT, encoding="utf-8")
    blocks = load_module().parse_transcript(f)
    assert [b["vo"] for b in blocks] == ["Benvenuti al corso. Oggi parliamo di budget.",
                                         "Il primo passo sono le entrate.", ""]
    assert blocks[1]["notes"] == ["NOTA: aggiungere un grafico"]
    assert blocks[2]["notes"] == ["TODO: controllare i numeri"]


def test_parse_plain_text_and_stray_separators(tmp_path):
    f = tmp_path / "t.txt"
    f.write_text("---\nFirst slide text\n\n---\nSecond\nslide\n---\n\n", encoding="utf-8")
    assert [b["vo"] for b in load_module().parse_transcript(f)] == ["First slide text", "Second slide"]


# --------------------------------------------------------------------------- --import

def test_import_writes_notes_and_backup(run, api, deck_plain, transcript, tmp_path):
    p = run(deck_plain.name, "--import", transcript.name, key=None, voice=None)
    assert p.returncode == 0, p.out
    notes = notes_of(deck_plain)
    assert notes[0] == "VO: Benvenuti al corso. Oggi parliamo di budget."
    assert "Presenter reminder" in notes[1]                        # existing notes are kept
    assert "old text" not in notes[1]                              # the old VO: line is replaced
    assert "VO: Il primo passo sono le entrate." in notes[1] and "NOTA: aggiungere un grafico" in notes[1]
    assert notes[2] == "TODO: controllare i numeri"                # author notes only, no VO:
    assert notes[3] == ""                                          # no block for slide 4
    assert "no block for slide 4" in p.out
    assert len(list(tmp_path.glob("plain.bak-*.pptx"))) == 1
    assert api.tts_calls == []


def test_import_then_preview_reads_only_vo(run, api, deck_plain, transcript):
    run(deck_plain.name, "--import", transcript.name)
    p = run(deck_plain.name, "--preview")
    assert p.returncode == 0, p.out
    assert "Benvenuti al corso." in p.out and "Il primo passo" in p.out
    assert "grafico" not in p.out and "controllare" not in p.out and "2026-11-11" not in p.out


def test_import_twice_does_not_duplicate(run, api, deck_plain, transcript):
    run(deck_plain.name, "--import", transcript.name)
    run(deck_plain.name, "--import", transcript.name)
    assert notes_of(deck_plain)[0].count("VO:") == 1


def test_import_on_chosen_slides(run, api, deck_plain, transcript):
    p = run(deck_plain.name, "--import", transcript.name, "--slides", "2-4")
    assert p.returncode == 0, p.out
    notes = notes_of(deck_plain)
    assert notes[0] == "" and notes[1].startswith("Presenter reminder") and "Benvenuti" in notes[1]


def test_import_too_many_blocks_changes_nothing(run, api, deck_plain, tmp_path):
    t = tmp_path / "long.txt"
    t.write_text("\n---\n".join(f"block {i}" for i in range(6)), encoding="utf-8")
    before = deck_plain.read_bytes()
    p = run(deck_plain.name, "--import", t.name)
    assert p.returncode == 1 and "6 blocks" in p.out
    assert deck_plain.read_bytes() == before and not list(tmp_path.glob("*.bak-*"))


def test_backups_skipped_by_wildcard(run, api, deck_plain, transcript):
    run(deck_plain.name, "--import", transcript.name)
    p = run("*.pptx", "--preview")
    assert p.out.count("plain.pptx") == 1 and ".bak-" not in p.out


def test_import_warns_when_vo_comments_win(run, api, deck_comments, tmp_path):
    t = tmp_path / "t.txt"
    t.write_text("Dictated text", encoding="utf-8")
    p = run(deck_comments.name, "--import", t.name)
    assert p.returncode == 0 and "take priority" in p.out, p.out


# --------------------------------------------------------------------------- --verify

def test_compare_text():
    mod = load_module()
    score, diffs = mod.compare_text('First part. <break time="1.0s" /> Second part.', "first part, second part")
    assert score == 1.0 and diffs == []
    score, diffs = mod.compare_text("open the link now", "open the lid")
    assert score < 0.85 and diffs == ['"link now" heard as "lid"']


def test_verify_flags_the_wrong_slide(run, api, deck_comments, tmp_path):
    heard = ["Welcome to the course.", "Second slide on two paragraph", "First part, second part."]
    p = run(deck_comments.name, "-y", "--verify", whisper_stub=True,
            env_extra={"VO_STUB_HEARD": json.dumps(heard)})
    assert p.returncode == 0, p.out
    assert "slide 1: done" in p.out and "check 100%" in p.out
    assert re.search(r"slide 2: the audio matches the text at 80%: \"paragraphs\" heard as \"paragraph\"", p.out), p.out
    assert "2 of 3 slides match the text, 1 to listen to" in p.out
    assert (tmp_path / "comments_vo.pptx").exists()


def test_verify_model_failure_stops_before_spending(run, api, deck_comments):
    p = run(deck_comments.name, "-y", "--verify", whisper_stub=True, env_extra={"VO_STUB_FAIL_LOAD": "1"})
    assert p.returncode == 1 and "Cannot load the whisper model" in p.out
    assert api.tts_calls == []                                     # no credits spent


def test_verify_works_on_cached_audio(run, api, deck_comments):
    run(deck_comments.name, "-y")
    api.tts_calls.clear()
    heard = ["Welcome to the course.", "Second slide, on two paragraphs.", "First part. Second part."]
    p = run(deck_comments.name, "--verify", whisper_stub=True, env_extra={"VO_STUB_HEARD": json.dumps(heard)})
    assert p.returncode == 0 and api.tts_calls == [], p.out
    assert "3 of 3 slides match the text" in p.out
