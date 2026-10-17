"""Narration audio, per-slide video clips, and the final video.

Pictures come from PowerPoint itself (``Presentation.CreateVideo``), so entrance animations and
embedded video clips play exactly as in a slide show. Each slide is rendered on its own (all other
slides hidden), silent. FFmpeg then shifts it to start with the narration, ``lead`` seconds in, and
stretches it slightly so the builds keep pace with the real voice (see ``timing``). Rendering
one slide at a time lets unchanged slides be reused and lets approved slides be rendered before the
rest of the deck is ready.
"""

import json
import logging
import os
import re
import subprocess
import time
import wave
from pathlib import Path

import imageio_ffmpeg

from .deck import PowerPoint, sha256_bytes, sha256_file, write_atomic

log = logging.getLogger("presgen")

DEFAULT_TTS = r'"D:\Dev\lesterthomas\csm-lester-voice\.venv\Scripts\csm-voice.exe" -f {text} -o {audio}'
FFMPEG = imageio_ffmpeg.get_ffmpeg_exe()
STATUS_IN_PROGRESS, STATUS_QUEUED, STATUS_DONE = 1, 2, 3   # PpMediaTaskStatus


# ---------------------------------------------------------------- audio


def wav_secs(path: Path) -> float:
    with wave.open(str(path)) as w:
        return w.getnframes() / w.getframerate()


def synthesize(out: Path, slide, tts_cmd: str, retries: int = 3, force: bool = False):
    """Make audio_NN.wav from text_NN.txt. Reused while it is newer than the text. Returns the path or None."""
    text = out / f"text_{slide.stem}.txt"
    audio = out / f"audio_{slide.stem}.wav"
    if not slide.notes.strip():
        audio.unlink(missing_ok=True)
        return None
    if force:
        audio.unlink(missing_ok=True)
    if audio.exists() and audio.stat().st_mtime > text.stat().st_mtime:
        log.info(f"Slide {slide.number}: narration up to date")
        return audio
    env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
    cmd = tts_cmd.format(text=text.name, audio=audio.name)
    for attempt in range(1, retries + 1):
        log.info(f"Slide {slide.number}: generating narration (attempt {attempt}/{retries})")
        t0 = time.time()
        r = subprocess.run(cmd, cwd=str(out), capture_output=True, text=True, encoding="utf-8",
                           errors="replace", env=env)
        if r.returncode == 0 and audio.exists():
            log.info(f"Slide {slide.number}: {wav_secs(audio):.1f}s of narration in {time.time() - t0:.0f}s")
            return audio
        tail = (r.stderr or r.stdout or "").strip().splitlines()[-3:]
        log.warning(f"Slide {slide.number}: TTS failed with exit code {r.returncode}: {' | '.join(tail)}")
    log.error(f"Slide {slide.number}: no narration after {retries} attempts")
    return None


# ---------------------------------------------------------------- video


def timing(slide, audio, wps: float = 2.5, max_stretch: float = 1.25) -> dict:
    """Compare the narration the slide was animated for with the real narration.

    Builds are planned at `wps` words per second from the start of the narration (that is how the
    deck's animations and clips are timed). The real voice is a little faster or slower, so the slide's
    pictures are stretched by `factor` (voice / planned) to stay in step with it. A factor beyond
    `max_stretch` is clamped and flagged: that slide's animation should be retimed at source.
    """
    words = len(slide.notes.split())
    if not audio or not words:
        return {"planned": 0.0, "voice": 0.0, "factor": 1.0, "clamped": False}
    planned = words / wps
    voice = wav_secs(audio)
    raw = voice / planned
    factor = min(max(raw, 1 / max_stretch), max_stretch)
    return {"planned": round(planned, 2), "voice": round(voice, 2), "factor": round(factor, 3),
            "clamped": abs(raw - factor) > 1e-3}


def slide_secs(slide, audio, lead: float, tail: float, factor: float = 1.0, min_secs: float = 3.0) -> float:
    """Long enough for the narration, every (stretched) animation and any embedded clip."""
    secs = max(min_secs, lead + slide.anim_secs * factor + 0.5)
    if audio:
        secs = max(secs, lead + wav_secs(audio) + tail)
    return round(secs, 2)


def render_secs(slide_total: float, lead: float, factor: float) -> float:
    """How long PowerPoint must render so that, once shifted by `lead` and stretched, it fills the slide."""
    return round((slide_total - lead) / factor + 0.2, 2)


def _ffmpeg(*args) -> None:
    r = subprocess.run([FFMPEG, "-hide_banner", "-loglevel", "error", "-y", *args], capture_output=True, text=True)
    if r.returncode != 0:
        raise RuntimeError(f"ffmpeg failed: {r.stderr.strip()}")


def _has_audio(video: Path) -> bool:
    r = subprocess.run([FFMPEG, "-hide_banner", "-i", str(video)], capture_output=True, text=True)
    return bool(re.search(r"Stream #\S+.*Audio:", r.stderr))


def _move_when_released(src: Path, dst: Path, wait: float = 30) -> None:
    """PowerPoint reports "done" a moment before it lets go of the file."""
    deadline = time.time() + wait
    while True:
        try:
            return src.replace(dst)
        except PermissionError:
            if time.time() > deadline:
                raise
            time.sleep(0.25)


def powerpoint_videos(work_deck: Path, jobs: list, quality: int = 100, fps: int = 30) -> None:
    """jobs: [(slide, secs, out_mp4)]. Renders each slide alone, silent apart from any embedded media."""
    if not jobs:
        return
    with PowerPoint() as pp:
        pres = pp.open(work_deck)
        try:
            n = pres.Slides.Count
            originally_hidden = {i for i in range(1, n + 1) if pres.Slides(i).SlideShowTransition.Hidden}
            for slide, secs, mp4 in jobs:
                for i in range(1, n + 1):
                    pres.Slides(i).SlideShowTransition.Hidden = 0 if i == slide.number else -1
                t = pres.Slides(slide.number).SlideShowTransition
                t.AdvanceOnTime, t.AdvanceTime = -1, secs
                mp4.parent.mkdir(parents=True, exist_ok=True)
                tmp = mp4.with_suffix(".tmp.mp4")
                tmp.unlink(missing_ok=True)
                t0 = time.time()
                pres.CreateVideo(str(tmp.resolve()), True, max(1, round(secs)), 1080, fps, quality)
                deadline = t0 + 120 + secs * 4
                while pres.CreateVideoStatus in (STATUS_IN_PROGRESS, STATUS_QUEUED):
                    if time.time() > deadline:
                        raise TimeoutError(f"PowerPoint took too long to render slide {slide.number}")
                    time.sleep(0.25)
                if pres.CreateVideoStatus != STATUS_DONE or not tmp.exists():
                    raise RuntimeError(f"PowerPoint could not render slide {slide.number} "
                                       f"(status {pres.CreateVideoStatus})")
                _move_when_released(tmp, mp4)
                log.info(f"Slide {slide.number}: rendered {secs:.1f}s by PowerPoint in {time.time() - t0:.1f}s")
            for i in range(1, n + 1):
                pres.Slides(i).SlideShowTransition.Hidden = -1 if i in originally_hidden else 0
        finally:
            pres.Close()


def still_video(png: Path, secs: float, mp4: Path, fps: int = 30) -> None:
    """Fallback renderer: the slide's final state as a still, no animation."""
    mp4.parent.mkdir(parents=True, exist_ok=True)
    _ffmpeg("-loop", "1", "-framerate", str(fps), "-i", str(png), "-t", f"{secs}",
            "-c:v", "libx264", "-preset", "medium", "-crf", "18", "-tune", "stillimage",
            "-pix_fmt", "yuv420p", "-vf", "scale=1920:1080", str(mp4))


def add_narration(silent: Path, audio, secs: float, lead: float, mp4: Path, factor: float = 1.0,
                  fps: int = 30) -> None:
    """Align the pictures with the narration and mux them together.

    The rendered slide starts `lead` seconds in, together with the narration (the first frame is held
    until then), and is stretched by `factor` so its builds keep pace with the real voice. Any sound
    the slide already has is stretched to match and mixed in.
    """
    mp4.parent.mkdir(parents=True, exist_ok=True)
    narration = (["-i", str(audio)] if audio else ["-f", "lavfi", "-i", "anullsrc=r=48000:cl=stereo"])
    ms = int(lead * 1000)
    chain = (f"[0:v]setpts=PTS*{factor},tpad=start_duration={lead}:start_mode=clone,fps={fps},format=yuv420p[vid];"
             f"[1:a]aresample=48000,adelay={ms}:all=1,apad[n]")
    if _has_audio(silent):
        chain += (f";[0:a]aresample=48000,atempo={1 / factor:.5f},adelay={ms}:all=1,apad[v];"
                  "[n][v]amix=inputs=2:normalize=0[a]")
    else:
        chain += ";[n]anull[a]"
    _ffmpeg("-i", str(silent), *narration, "-filter_complex", chain, "-map", "[vid]", "-map", "[a]",
            "-c:v", "libx264", "-preset", "medium", "-crf", "18", "-c:a", "aac", "-b:a", "192k", "-ac", "2",
            "-t", f"{secs}", "-movflags", "+faststart", str(mp4))


def concat(clips: list, mp4: Path) -> None:
    listing = mp4.with_name("concat_list.txt")
    listing.write_text("".join(f"file '{c.resolve().as_posix()}'\n" for c in clips), encoding="utf-8")
    try:
        _ffmpeg("-f", "concat", "-safe", "0", "-i", str(listing), "-c", "copy", "-movflags", "+faststart", str(mp4))
    finally:
        listing.unlink(missing_ok=True)


def chapters(slides: list, secs: list) -> str:
    """YouTube-style chapter list: one line per slide title, deduplicated."""
    lines, t, last = [], 0.0, None
    for s, d in zip(slides, secs):
        if s.title != last:
            lines.append(f"{int(t // 60)}:{int(t % 60):02d} {s.title}")
            last = s.title
        t += d
    return "\n".join(lines) + "\n"


# ---------------------------------------------------------------- render manifest


class Manifest:
    """render.json: what each clip was made from, so unchanged slides are reused."""

    def __init__(self, out: Path):
        self.path = out / "render.json"
        self.data = json.loads(self.path.read_text(encoding="utf-8")) if self.path.exists() else {}

    def key(self, slide, audio, renderer: str, secs: float, lead: float, factor: float = 1.0) -> str:
        return sha256_bytes(json.dumps([slide.fingerprint, sha256_file(audio) if audio else None,
                                        renderer, secs, lead, factor, 2]).encode())[:16]

    def current(self, out: Path, slide, key: str) -> bool:
        e = self.data.get(slide.stem, {})
        return e.get("key") == key and (out / e.get("clip", "-")).exists()

    def record(self, slide, key, clip, audio, secs, renderer, timing=None) -> None:
        self.data[slide.stem] = {"key": key, "fingerprint": slide.fingerprint, "clip": clip,
                                 "audio": audio.name if audio else None, "secs": secs, "renderer": renderer,
                                 "timing": timing or {}}

    def save(self) -> None:
        write_atomic(self.path, json.dumps(self.data, indent=2) + "\n")
