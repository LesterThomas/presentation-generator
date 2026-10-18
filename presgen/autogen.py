"""Background generation for the review server: approve a slide and its narrated clip follows.

The generator keeps every approved slide's clip current. Every few seconds (or at once, when poked by
the review page) it looks for approved slides with no clip for their current version, and narrates
and renders them one at a time: the voice needs the GPU to itself. Approvals made anywhere count,
whether on the page or from the command line. When every visible slide is approved and rendered, it
rebuilds the final video.
"""

import logging
import threading
import time

import comtypes

from . import pipeline, render
from .deck import load_slides
from .review import ReviewState, output_dir

log = logging.getLogger("presgen")


class AutoGenerator(threading.Thread):
    def __init__(self, pptx, opts: pipeline.Options, interval: float = 3.0):
        super().__init__(daemon=True, name="presgen-autogen")
        self.pptx, self.opts, self.interval = pptx, opts, interval
        self.lock = threading.Lock()
        self.jobs = {}        # stem -> {"state": queued|narrating|rendering|failed, "error": ..., "at": ...}
        self.final = {}       # {"state": building|ready|waiting, ...}
        self.revoice = set()  # stems whose narration should be generated again
        self.failed_fp = {}   # stem -> fingerprint that failed (not retried until it changes or on request)
        self._wake = threading.Event()
        self._last_ready = None

    # ---- called from the web server
    def poke(self) -> None:
        self._wake.set()

    def request(self, stem: str, revoice: bool = False) -> None:
        with self.lock:
            self.failed_fp.pop(stem, None)
            if revoice:
                self.revoice.add(stem)
            # Show the request on the page at once, even while another slide is being narrated.
            if self.jobs.get(stem, {}).get("state") not in ("narrating", "rendering"):
                self.jobs[stem] = {"state": "queued", "at": time.strftime("%H:%M:%S")}
        self.poke()

    def snapshot(self) -> dict:
        with self.lock:
            return {"jobs": {k: dict(v) for k, v in self.jobs.items()}, "final": dict(self.final)}

    # ---- the worker
    def _set(self, stem, state, **extra) -> None:
        with self.lock:
            if state is None:
                self.jobs.pop(stem, None)
            else:
                self.jobs[stem] = {"state": state, "at": time.strftime("%H:%M:%S"), **extra}

    def run(self) -> None:
        comtypes.CoInitialize()  # PowerPoint COM from this thread
        log.info("Auto-generation is on: approved slides are narrated and rendered straight away")
        while True:
            self._wake.wait(self.interval)
            self._wake.clear()
            try:
                self._reconcile()
            except Exception as e:  # keep the server alive; report and try again later
                log.exception(f"Auto-generation: {e}")

    def _todo(self):
        out = output_dir(self.pptx)
        slides = load_slides(out)
        state = ReviewState(self.pptx)
        manifest = render.Manifest(out)
        todo = []
        for s in slides:
            if s.hidden or not state.is_approved(s):
                continue
            with self.lock:
                again = s.stem in self.revoice
                failed = self.failed_fp.get(s.stem) == s.fingerprint
            if again or (not failed and not pipeline.clip_current(out, manifest, s)):
                todo.append(s)
        with self.lock:  # slides the user asked to re-voice go first
            todo.sort(key=lambda s: s.stem not in self.revoice)
        return out, slides, state, todo

    def _reconcile(self) -> None:
        out, slides, state, todo = self._todo()
        for s in todo:
            if self.jobs.get(s.stem, {}).get("state") != "failed":
                self._set(s.stem, "queued")
        did = False
        for i, s in enumerate(todo):
            with self.lock:
                jump = i > 0 and bool(self.revoice - {x.stem for x in todo[:i]})
            if jump:
                # A re-voice was requested while this pass was running: start a new pass, so it goes
                # next instead of waiting for every slide queued before it.
                self.poke()
                break
            current = {x.slide_id: x for x in load_slides(out)}.get(s.slide_id)
            if current is None or current.fingerprint != s.fingerprint or not ReviewState(self.pptx).is_approved(current):
                self._set(s.stem, None)  # changed or un-approved while waiting
                continue
            with self.lock:
                redo = {s.number} if s.stem in self.revoice else set()
                self.revoice.discard(s.stem)
            deck = pipeline.snapshot(out, f"render_{s.stem}")
            try:
                done, failed = pipeline.render_slides(self.pptx, [s], self.opts, redo, deck,
                                                      on_stage=lambda sl, st, **info: self._set(sl.stem, st, **info))
            except Exception as e:
                log.exception(f"Slide {s.number}: could not render")
                done, failed = [], [s]
                self._set(s.stem, "failed", error=str(e)[:300])
            finally:
                deck.unlink(missing_ok=True)
            if failed:
                with self.lock:
                    self.failed_fp[s.stem] = s.fingerprint
                if self.jobs.get(s.stem, {}).get("state") != "failed":
                    self._set(s.stem, "failed", error="Narration failed after retries. See presgen.log.")
            else:
                self._set(s.stem, None)
                did = True
        self._final(did)

    def _final(self, did: bool) -> None:
        """Rebuild the final video whenever the set of approved, rendered, validated slides changes."""
        out = output_dir(self.pptx)
        state = ReviewState(self.pptx)
        manifest = render.Manifest(out)
        visible = [s for s in load_slides(out) if not s.hidden]
        ready = tuple((s.stem, s.fingerprint, manifest.data.get(s.stem, {}).get("key")) for s in visible
                      if state.is_approved(s) and pipeline.clip_current(out, manifest, s)
                      and state.validated(s, manifest.data.get(s.stem, {}).get("key")))
        if ready == self._last_ready and not did:
            return
        self._last_ready = ready
        if len(ready) < len(visible):
            with self.lock:
                self.final = {"state": "waiting", "ready": len(ready), "total": len(visible)}
            return
        with self.lock:
            self.final = {"state": "building"}
        video = pipeline.build_final(self.pptx, require_validation=True)
        with self.lock:
            self.final = {"state": "ready" if video else "waiting", "ready": len(ready), "total": len(visible)}
