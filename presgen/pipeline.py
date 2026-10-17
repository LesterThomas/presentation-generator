"""Narrate and render slides, and join the clips into the final video.

Shared by `presgen.py generate` (a batch from the command line) and the review server's background
generator (one slide at a time, as soon as it is approved).
"""

import json
import logging
import os
import random
import re
import shutil
from dataclasses import dataclass
from pathlib import Path

from . import render
from .deck import load_slides
from .review import ReviewState, output_dir, video_audio

log = logging.getLogger("presgen")


@dataclass
class Options:
    renderer: str = "powerpoint"
    tts: str = render.DEFAULT_TTS
    retries: int = 3
    lead: float = 1.0
    tail: float = 0.75
    quality: int = 100
    wps: float = 2.5          # narration speed the deck's animations were planned for
    max_stretch: float = 1.25
    # tts_cache: str = render.DEFAULT_TTS_CACHE  # cleared for a slide when it is re-voiced

    @classmethod
    def from_args(cls, a) -> "Options":
        return cls(**{k: getattr(a, k) for k in cls.__dataclass_fields__ if hasattr(a, k)})


VOICE_CACHE = Path(r"D:\Dev\lesterthomas\csm-lester-voice\cache")  # csm-voice's per-sentence takes


def forget_voice_takes(text: str, cache: Path = VOICE_CACHE) -> int:
    """Delete csm-voice's cached takes of these sentences, in every cache folder, so they are voiced afresh.

    csm-voice names a take <first three words, sanitised>_<sha256>.wav. Matching on the first three words
    can also drop another sentence that starts the same way; that only costs a retake.
    """
    if not cache.is_dir():
        return 0
    removed = 0
    for sentence in re.split(r"(?<=[.!?])\s+|\n+", text.strip()):
        words = sentence.split()
        if not words:
            continue
        stem = "".join(c if c.isalnum() or c == "_" else "_" for c in "_".join(words[:3])[:50])
        for folder in [cache, *[d for d in cache.iterdir() if d.is_dir()]]:
            for f in folder.glob(f"{stem}_*.wav"):
                if re.fullmatch(re.escape(stem) + r"_[0-9a-f]{64}\.wav", f.name):
                    f.unlink(missing_ok=True)
                    removed += 1
    return removed


def audio_newer(out: Path, e: dict) -> bool:
    """Has the slide's narration (e.g. a new recording) changed since its clip was made?"""
    audio, clip = out / (e.get("audio") or "-"), out / e.get("clip", "-")
    return audio.exists() and clip.exists() and audio.stat().st_mtime > clip.stat().st_mtime


def clip_current(out: Path, manifest: "render.Manifest", slide) -> bool:
    """Is there a clip made from exactly this version of the slide and its narration?"""
    e = manifest.data.get(slide.stem, {})
    return (e.get("fingerprint") == slide.fingerprint and "timing" in e  # clips before timing alignment are redone
            and (e.get("audio") or not slide.notes.strip())              # a narrated slide's clip must have narration
            and (out / e.get("clip", "-")).exists()
            and not audio_newer(out, e))


def render_slides(pptx: Path, slides: list, opts: Options, redo_audio=frozenset(), deck: Path = None,
                  on_stage=lambda slide, stage, **info: None):
    """Narrate and render each slide. Returns (rendered, failed) lists of slides.

    `deck` is the presentation PowerPoint renders from (default: the work copy from the last
    inspection). `on_stage(slide, stage, **info)` reports progress: narrating, rendering, failed
    (with `error`).
    """
    out = output_dir(pptx)
    deck = deck or out / ".work" / "deck.pptx"
    audio, failed, plan, jobs = {}, [], [], []
    manifest = render.Manifest(out)

    for s in slides:  # 1. narration: the slow part
        on_stage(s, "narrating")
        own = video_audio(out, s)
        if own:  # a presenter video: its own sound (EQ-matched) is the narration, never the voice
            audio[s.number] = own
            log.info(f"Slide {s.number}: narration is the sound of the slide's video")
            continue
        tts = opts.tts
        if s.number in redo_audio:
            # Re-voice: a genuinely new take. Forget the voice's cached takes of this slide's sentences and
            # sample with a new seed (the same seed and text would give the same take again).
            n = forget_voice_takes(s.notes)
            tts = f"{opts.tts} --seed {random.randint(1, 10**6)}"
            log.info(f"Slide {s.number}: re-voicing with a new seed ({n} cached sentence takes cleared)")
        try:
            audio[s.number] = render.synthesize(out, s, tts, opts.retries, force=s.number in redo_audio)
        except getattr(render, "TTSFailed", RuntimeError) as e:
            audio[s.number] = None
            failed.append(s)
            on_stage(s, "failed", error=str(e))
            continue
        if s.notes.strip() and audio[s.number] is None:  # never render a narrated slide without its narration
            failed.append(s)
            on_stage(s, "failed", error="Narration failed: see presgen.log for the voice's error")

    timings = {}
    for s in slides:  # 2. pictures, timed to the real narration
        if s in failed:
            continue
        t = render.timing(s, audio[s.number], opts.wps, opts.max_stretch)
        meta = out / f"audio_{s.stem}.json"
        if meta.exists() and json.loads(meta.read_text(encoding="utf-8")).get("sync"):
            # Narration recorded with the slide's own video (a presenter clip): never stretch, or lips drift.
            t = {**t, "factor": 1.0, "clamped": False, "sync": True}
        timings[s.number] = t
        if t["voice"]:
            log.info(f"Slide {s.number}: narration {t['voice']:.1f}s, planned {t['planned']:.1f}s"
                     f" -> animation timing x{t['factor']:.2f}"
                     + ("  (CLAMPED: retime this slide's animation at source)" if t["clamped"] else ""))
        secs = render.slide_secs(s, audio[s.number], opts.lead, opts.tail, t["factor"])
        key = manifest.key(s, audio[s.number], opts.renderer, secs, opts.lead, t["factor"])
        clip = f"clips/clip_{s.stem}.mp4"
        plan.append((s, secs, key, clip))
        if manifest.current(out, s, key):
            e = manifest.data.get(s.stem, {})
            if audio_newer(out, e):  # narration rewritten with identical audio: the clip is still right
                os.utime(out / e["clip"])
            log.info(f"Slide {s.number}: clip up to date")
            continue
        jobs.append((s, render.render_secs(secs, opts.lead, t["factor"]), out / ".work" / f"silent_{s.stem}.mp4"))
    for s, _, _ in jobs:
        on_stage(s, "rendering")
    if jobs:
        if opts.renderer == "powerpoint":
            render.powerpoint_videos(deck, jobs, quality=opts.quality)
        else:
            for s, rsecs, silent in jobs:
                render.still_video(out / f"slide_{s.stem}.png", rsecs, silent)

    fresh = {s.number for s, _, _ in jobs}  # 3. sound
    for s, secs, key, clip in plan:
        if s.number in fresh:
            silent = out / ".work" / f"silent_{s.stem}.mp4"
            render.add_narration(silent, audio[s.number], secs, opts.lead, out / clip, timings[s.number]["factor"])
            silent.unlink(missing_ok=True)
            meta = out / f"audio_{s.stem}.json"
            phrased = json.loads(meta.read_text(encoding="utf-8")).get("phrased") if meta.exists() else False
            manifest.record(s, key, clip, audio[s.number], secs, opts.renderer, {**timings[s.number], "phrased": phrased})
            manifest.save()
            log.info(f"Slide {s.number}: clip ready ({secs:.1f}s)")
    return [s for s, _, _, _ in plan], failed


def build_final(pptx: Path, partial: bool = False, ignore_approvals: bool = False, require_validation: bool = False):
    """Join the clips into the final video once every visible slide is approved and rendered.

    With require_validation (the review server), each slide's narration must also be validated.

    Returns the video's path, or None (with the reason logged) when the deck isn't ready.
    """
    out = output_dir(pptx)
    state = ReviewState(pptx)
    manifest = render.Manifest(out)
    visible = [s for s in load_slides(out) if not s.hidden]
    ready = [s for s in visible if clip_current(out, manifest, s) and (ignore_approvals or state.is_approved(s))
             and (not require_validation or state.validated(s, manifest.data.get(s.stem, {}).get("key")))]
    missing = [s for s in visible if s not in ready]
    final = out / f"{out.name}_video.mp4"
    if (missing and not partial) or not ready:
        if final.exists() and missing:
            log.info("The existing final video is out of date")
        log.info(f"Final video not built: {len(missing)} slide(s) not approved or not rendered yet: "
                 + ", ".join(str(s.number) for s in missing))
        return None
    built = {s.stem: s.fingerprint for s in ready}
    if manifest.data.get("_final") == built and final.exists():
        log.info(f"Final video up to date: {final}")
        return final
    render.concat([out / manifest.data[s.stem]["clip"] for s in ready], final)
    secs = [manifest.data[s.stem]["secs"] for s in ready]
    (out / f"{out.name}_chapters.txt").write_text(render.chapters(ready, secs), encoding="utf-8")
    manifest.data["_final"] = built
    manifest.save()
    total = sum(secs)
    log.info(f"Final video: {final}  ({len(ready)} slides, {int(total // 60)}:{int(total % 60):02d})")
    return final


def snapshot(out: Path, name: str) -> Path:
    """A private copy of the work deck, so a concurrent `refresh` can replace the work copy."""
    src = out / ".work" / "deck.pptx"
    dst = out / ".work" / f"{name}.pptx"
    shutil.copy2(src, dst)
    return dst
