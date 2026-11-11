"""Test stand-in for faster-whisper, put first on PYTHONPATH by the tests.

It "hears" the texts listed in the VO_STUB_HEARD environment variable (a JSON list), one per
transcribed clip, in order. VO_STUB_FAIL_LOAD makes model loading fail like a blocked download.
"""
import json
import os


class _Segment:
    def __init__(self, text):
        self.text = text


class WhisperModel:
    def __init__(self, model_size_or_path, device="cpu", compute_type="default", **kw):
        if os.environ.get("VO_STUB_FAIL_LOAD"):
            raise RuntimeError("could not download the model")
        self.name = model_size_or_path
        self._heard = json.loads(os.environ.get("VO_STUB_HEARD", "[]"))

    def transcribe(self, audio, language=None, **kw):
        text = self._heard.pop(0) if self._heard else ""
        return iter([_Segment(text)]), None
