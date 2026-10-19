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

# --candidates 3: csm-voice voices each sentence three times (sampled), cuts each take at its first long
# pause, and keeps the take closest to the expected length. A single greedy take often drops words or
# runs on after the sentence.
DEFAULT_TTS = r'"D:\Dev\lesterthomas\csm-lester-voice\.venv\Scripts\csm-voice.exe" -f {text} -o {audio} --candidates 3'
FFMPEG = imageio_ffmpeg.get_ffmpeg_exe()
STATUS_IN_PROGRESS, STATUS_QUEUED, STATUS_DONE = 1, 2, 3   # PpMediaTaskStatus


# ---------------------------------------------------------------- audio


def wav_secs(path: Path) -> float:
    with wave.open(str(path)) as w:
        return w.getnframes() / w.getframerate()


def _tts_env() -> dict:
    # The voice model is cached locally. Offline mode stops csm-voice contacting huggingface.co on every
    # run, which fails behind TLS-intercepting antivirus (set HF_HUB_OFFLINE=0 to allow downloads).
    return {"HF_HUB_OFFLINE": "1", "TRANSFORMERS_OFFLINE": "1", **os.environ,
            "PYTHONIOENCODING": "utf-8", "PYTHONUTF8": "1"}


class VoiceWorker:
    """A long-lived `csm-voice --worker` process, so the voice model is loaded once, not for every slide.

    Jobs go to its stdin as JSON lines; it answers with "@@READY", "@@DONE" or "@@ERROR" lines on
    stdout. Its log (stderr) is drained on a thread, so the pipe never fills, and the last lines are
    kept for error reports.
    """

    def __init__(self, args: list):
        import collections
        import threading
        self.args, self.tail = args, collections.deque(maxlen=40)
        log.info("Starting the voice worker (the model stays loaded between slides)")
        t0 = time.time()
        self.proc = subprocess.Popen(args + ["--worker"], stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                                     stderr=subprocess.PIPE, text=True, encoding="utf-8", errors="replace",
                                     bufsize=1, env=_tts_env())
        threading.Thread(target=self._drain, daemon=True, name="voice-worker-log").start()
        ready = self._answer()
        if not ready or ready[0] != "READY":
            self.close()
            raise RuntimeError("voice worker did not start: " + " | ".join(list(self.tail)[-3:]))
        log.info(f"Voice worker ready in {time.time() - t0:.0f}s (model load {ready[1].get('load_secs')}s)")

    def _drain(self):
        for line in self.proc.stderr:
            self.tail.append(line.rstrip())

    def _answer(self):
        for line in self.proc.stdout:
            if line.startswith("@@"):
                kind, _, rest = line[2:].strip().partition(" ")
                return kind, json.loads(rest or "{}")
        return None  # the worker exited

    def alive(self) -> bool:
        return self.proc.poll() is None

    def run(self, text: Path, audio: Path, seed=None, context=None):
        """Voice one text file. Returns (ok, info); info["failed"] lists sentences that failed the voice's
        speech-to-text check. `context` (deck, slide) goes into the voice's record of every sentence."""
        job = {"file": str(text.resolve()), "output": str(audio.resolve())}
        if seed is not None:
            job["seed"] = seed
        if context:
            job["context"] = context
        try:
            self.proc.stdin.write(json.dumps(job) + "\n")
            self.proc.stdin.flush()
        except OSError:
            return False, {"error": "voice worker has exited"}
        ans = self._answer()
        if ans is None:
            return False, {"error": "voice worker exited: " + " | ".join(list(self.tail)[-3:])}
        return ans[0] == "DONE", ans[1]

    def close(self):
        if self.alive():
            try:
                self.proc.stdin.write('{"quit": true}\n')
                self.proc.stdin.flush()
                self.proc.wait(timeout=10)
            except Exception:
                self.proc.kill()


_workers: dict = {}


class TTSFailed(RuntimeError):
    """The voice could not say some sentences of a slide: after every take, its speech-to-text check
    still heard words missing. `failed` holds the voice's verdicts (sentence, transcript, missing words)."""

    def __init__(self, slide, failed: list):
        self.failed = failed
        heard = " ".join(f'"{f["sentence"]}" was heard as "{f.get("transcript", "")}".' for f in failed)
        super().__init__(f"Narration check failed on slide {slide.number}: {heard}")


def _worker_args(tts_cmd: str):
    """For a csm-voice command line, the worker's arguments and the job's seed; otherwise (None, None).

    The text and audio placeholders are dropped (they become per-job fields), and so is --seed (a
    re-voice passes a new seed per job). Set PRESGEN_VOICE_WORKER=0 to always run one process per slide.
    """
    import shlex
    if os.environ.get("PRESGEN_VOICE_WORKER", "1") == "0" or "csm-voice" not in tts_cmd:
        return None, None
    parts = [p.strip('"') for p in shlex.split(tts_cmd, posix=False)]
    args, seed, i = [], None, 0
    while i < len(parts):
        p = parts[i]
        if p in ("-f", "--file", "-o", "--output") and i + 1 < len(parts) and parts[i + 1] in ("{text}", "{audio}"):
            i += 2
            continue
        if p == "--seed" and i + 1 < len(parts):
            seed, i = int(parts[i + 1]), i + 2
            continue
        args.append(p)
        i += 1
    return args, seed


def _voice_worker(args: list):
    key = tuple(args)
    w = _workers.get(key)
    if w is None or not w.alive():
        w = _workers[key] = VoiceWorker(args)
    return w


def close_voice_workers():
    for w in _workers.values():
        w.close()
    _workers.clear()


import atexit  # noqa: E402

atexit.register(close_voice_workers)


def synthesize(out: Path, slide, tts_cmd: str, retries: int = 3, force: bool = False):
    """Make audio_NN.wav from text_NN.txt. Reused while it is newer than the text. Returns the path or None.

    A csm-voice command runs in a persistent worker (the model loads once); any other command, or a
    worker that can't start, runs as one process per slide.
    """
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
    wargs, seed = _worker_args(tts_cmd)
    cmd = tts_cmd.format(text=text.name, audio=audio.name)
    for attempt in range(1, retries + 1):
        log.info(f"Slide {slide.number}: generating narration (attempt {attempt}/{retries})")
        t0 = time.time()
        if wargs is not None:
            try:
                ok, info = _voice_worker(wargs).run(
                    text, audio, seed, {"deck": out.name, "slide": slide.number, "stem": slide.stem})
            except Exception as e:  # the worker couldn't start: fall back to one process per slide
                log.warning(f"Voice worker unavailable ({e}); running csm-voice once per slide")
                wargs = None
            else:
                if ok and audio.exists() and info.get("failed"):
                    # Takes were already retried inside the voice; the wording is the problem. Don't render
                    # a clip with a gap in it: report the sentences so they can be reworded.
                    for f in info["failed"]:
                        log.warning(f"Slide {slide.number}: narration check failed: \"{f['sentence']}\" "
                                    f"was heard as \"{f.get('transcript', '')}\" (missing: {' '.join(f.get('missing', []))})")
                    audio.unlink(missing_ok=True)
                    raise TTSFailed(slide, info["failed"])
                if ok and audio.exists():
                    checked = info.get("sentences") or []
                    note = f", {len(checked)} sentence(s) passed the speech-to-text check" if checked else ""
                    log.info(f"Slide {slide.number}: {wav_secs(audio):.1f}s of narration in {time.time() - t0:.0f}s (worker{note})")
                    return audio
                log.warning(f"Slide {slide.number}: TTS failed in the voice worker: {info.get('error')}")
                continue
        r = subprocess.run(cmd, cwd=str(out), capture_output=True, text=True, encoding="utf-8",
                           errors="replace", env=_tts_env())
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
                # Right after CreateVideo the status can still read 0 (not started) for a moment; wait
                # through that too, or a slide that is about to queue is reported as failed.
                while pres.CreateVideoStatus in (0, STATUS_IN_PROGRESS, STATUS_QUEUED):
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
    # fps must come before tpad: setpts leaves the frame rate unknown, and tpad then pads no frames at all,
    # which put every slide's picture `lead` seconds ahead of its narration.
    chain = (f"[0:v]setpts=PTS*{factor},fps={fps},tpad=start_duration={lead}:start_mode=clone,format=yuv420p[vid];"
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
