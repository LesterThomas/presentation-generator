"""Make a manual recording sound like the generated narration: matched tone and loudness.

The reference is the deck's generated narration (audio_NN.wav that is not someone's recording). From it
we take the average spectral shape (octave bands, speech frames only) and the integrated loudness. A
recording is then
  1. equalised: per-band gains that move its spectral shape towards the reference (at most +/-8 dB),
  2. turned up or down by one fixed gain to the reference loudness (EBU R128).
No compression, noise reduction or filtering: those were tried and the reviewer preferred the recording
with only EQ and a fixed gain (2026-10-17).
"""

import json
import logging
import re
import subprocess
import wave
from pathlib import Path

import numpy as np

from .deck import sha256_bytes, sha256_file, write_atomic
from .render import FFMPEG

log = logging.getLogger("presgen")

BANDS = [63, 125, 250, 500, 1000, 2000, 4000, 8000, 11000]   # octave-band centres, Hz
MAX_GAIN = 8.0                                               # dB, per band
DEFAULT_LUFS = -24.0


def _read_wav(path: Path) -> tuple:
    with wave.open(str(path)) as w:
        sr, n, ch, width = w.getframerate(), w.getnframes(), w.getnchannels(), w.getsampwidth()
        raw = w.readframes(n)
    a = np.frombuffer(raw, dtype={2: np.int16, 4: np.int32}[width]).astype(np.float64)
    a = a.reshape(-1, ch).mean(axis=1) / (2 ** (8 * width - 1))
    return a, sr


def band_shape(path: Path) -> np.ndarray:
    """Average level per octave band (dB, relative to the mean over bands), from speech frames only."""
    a, sr = _read_wav(path)
    n = 2048
    if len(a) < n * 4:
        raise ValueError(f"{path.name} is too short to analyse")
    frames = np.lib.stride_tricks.sliding_window_view(a, n)[:: n // 2] * np.hanning(n)
    energy = (frames ** 2).mean(axis=1)
    speech = frames[energy > energy.max() * 10 ** (-35 / 10)]           # within 35 dB of the loudest frame
    power = (np.abs(np.fft.rfft(speech, axis=1)) ** 2).mean(axis=0)
    freqs = np.fft.rfftfreq(n, 1 / sr)
    levels = []
    for c in BANDS:
        lo, hi = c / 2 ** 0.5, min(c * 2 ** 0.5, sr / 2)
        sel = (freqs >= lo) & (freqs < hi)
        levels.append(10 * np.log10(power[sel].mean() + 1e-20) if sel.any() else -200.0)
    levels = np.array(levels)
    return levels - levels[(levels > -150)].mean()


def loudness(path: Path) -> float:
    r = subprocess.run([FFMPEG, "-hide_banner", "-nostats", "-i", str(path), "-af", "ebur128=framelog=quiet", "-f", "null", "-"],
                       capture_output=True, text=True)
    m = re.findall(r"I:\s+(-?[\d.]+) LUFS", r.stderr)
    return float(m[-1]) if m else DEFAULT_LUFS


def reference(out: Path) -> dict:
    """The generated voice's profile, from the deck's generated narration (cached in voice_profile.json)."""
    files = []
    for wav in sorted(out.glob("audio_[0-9][0-9].wav")):
        meta = wav.with_suffix(".json")
        manual = False
        if meta.exists():
            try:
                m = json.loads(meta.read_text(encoding="utf-8"))
                manual = bool(m.get("manual")) and m.get("sha") == sha256_file(wav)
            except ValueError:
                pass
        if not manual:
            files.append(wav)
    if not files:
        raise ValueError("no generated narration to use as the reference voice yet")
    key = sha256_bytes("".join(f"{f.name}:{f.stat().st_size}:{f.stat().st_mtime_ns}" for f in files).encode())[:16]
    cache = out / "voice_profile.json"
    if cache.exists():
        c = json.loads(cache.read_text(encoding="utf-8"))
        if c.get("key") == key:
            return c
    shapes = [band_shape(f) for f in files]
    prof = {"key": key, "files": len(files), "bands": BANDS,
            "shape": [round(float(v), 2) for v in np.median(np.array(shapes), axis=0)],
            "lufs": round(float(np.median([loudness(f) for f in files])), 1)}
    write_atomic(cache, json.dumps(prof, indent=1))
    log.info(f"Voice profile from {len(files)} generated narration files: {prof['lufs']} LUFS")
    return prof


def match(src: Path, dst: Path, out: Path) -> dict:
    """Process the trimmed recording `src` into `dst` so it sounds like the generated voice."""
    prof = reference(out)
    ref = np.array(prof["shape"])
    rec = band_shape(src)
    gains = np.clip(ref - rec, -MAX_GAIN, MAX_GAIN)
    gains[(rec < -150) | (ref < -150)] = 0.0
    gains = np.convolve(np.pad(gains, 1, mode="edge"), [0.25, 0.5, 0.25], mode="valid")  # smooth neighbouring bands
    eq = ";".join(f"entry({f},{g:.1f})" for f, g in zip(BANDS, gains))
    eqd = dst.with_name(dst.stem + ".eq.wav")

    def ffmpeg(inp, af, outp):
        r = subprocess.run([FFMPEG, "-hide_banner", "-loglevel", "error", "-y", "-i", str(inp), "-af", af,
                            "-ac", "1", "-ar", "24000", "-c:a", "pcm_s16le", str(outp)], capture_output=True, text=True)
        if r.returncode != 0:
            raise ValueError(f"could not process the recording: {r.stderr.strip()[-200:]}")

    ffmpeg(src, f"firequalizer=gain_entry='{eq}'", eqd)          # 1. EQ towards the voice
    gain = round(prof["lufs"] - loudness(eqd), 1)                  # 2. one fixed gain to its loudness
    ffmpeg(eqd, f"volume={gain}dB", dst)
    eqd.unlink(missing_ok=True)
    return {"eq_db": [round(float(g), 1) for g in gains], "gain_db": gain, "lufs_target": prof["lufs"],
            "lufs_before": loudness(src), "lufs_after": loudness(dst)}
