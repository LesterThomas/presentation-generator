"""presgen: review a narrated deck slide by slide, then turn the approved slides into a video.

    uv run presgen.py review   DECK.pptx            # stills + previews, opens the review page;
                                                    # approving a slide narrates and renders it
    uv run presgen.py approve  DECK.pptx 1-4,7      # or approve on the review page
    uv run presgen.py changes  DECK.pptx 5 -m "..." # request changes, with a comment
    uv run presgen.py status   DECK.pptx
    uv run presgen.py reply    DECK.pptx 5 -m "..." # a note back to the reviewer, shown on the page
    uv run presgen.py refresh  DECK.pptx            # after a rebuild: the open page updates itself
    uv run presgen.py generate DECK.pptx            # narration + clips for approved slides, then the video
    uv run presgen.py run      DECK.pptx            # everything, ignoring approvals (the old behaviour)
"""

import argparse
import faulthandler
import json
import logging
import sys
from pathlib import Path

from . import pipeline, render
from .autogen import AutoGenerator
from .deck import inspect_deck, load_slides
from .review import (ACTIONS, REPLY, ReviewState, change_status, output_dir, parse_slides, serve,
                     state_path, write_page)

log = logging.getLogger("presgen")


def setup_logging(out=None) -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    handlers = [logging.StreamHandler(sys.stdout)]
    if out:
        out.mkdir(parents=True, exist_ok=True)
        handlers.append(logging.FileHandler(out / "presgen.log", encoding="utf-8"))
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s",
                        datefmt="%H:%M:%S", handlers=handlers, force=True)


def inspect(pptx: Path):
    out = output_dir(pptx)
    slides = inspect_deck(pptx, out)
    ReviewState(pptx).sync(slides)
    return out, slides


# ---------------------------------------------------------------- review


def previews(pptx: Path, out: Path, slides: list) -> None:
    """Silent PowerPoint renders of each slide, so animations can be reviewed on the page."""
    cache_path = out / "previews" / "previews.json"
    cache = json.loads(cache_path.read_text(encoding="utf-8")) if cache_path.exists() else {}
    jobs = []
    for s in slides:
        if s.hidden:
            continue
        mp4 = out / "previews" / f"slide_{s.stem}.mp4"
        if cache.get(s.stem) == s.fingerprint and mp4.exists():
            continue
        secs = round(max(3.0, s.planned_secs, s.anim_secs + 1.0), 2)
        jobs.append((s, secs, mp4))
    if jobs:
        log.info(f"Rendering {len(jobs)} silent preview(s) in PowerPoint")
        render.powerpoint_videos(out / ".work" / "deck.pptx", jobs)
        for s, _, _ in jobs:
            cache[s.stem] = s.fingerprint
        cache_path.write_text(json.dumps(cache, indent=2), encoding="utf-8")


def cmd_review(a) -> None:
    out, slides = inspect(a.deck)
    if not a.no_previews:
        previews(a.deck, out, slides)
    page = write_page(a.deck)
    print_status(a.deck)
    if a.no_serve:
        log.info(f"Review page (read-only): {page}")
        return
    cmd_serve(a)


def cmd_serve(a) -> None:
    # Record native crashes (PowerPoint COM, FFmpeg pipes) that would otherwise end the server silently.
    global _crash_log
    _crash_log = open(output_dir(a.deck) / "crash.log", "a", encoding="utf-8")
    faulthandler.enable(_crash_log)
    gen = None if a.no_auto else AutoGenerator(a.deck, pipeline.Options.from_args(a))
    serve(a.deck, a.port, open_browser=not a.no_open, gen=gen)


def cmd_change(a) -> None:
    if a.cmd in ("changes", REPLY) and not a.message:
        raise SystemExit(f"`{a.cmd}` needs a message: -m \"...\"")
    chosen = change_status(a.deck, a.slides, a.cmd, a.message or "", by=a.by)
    write_page(a.deck)
    log.info(f"{ACTIONS.get(a.cmd, 'replied')}: slide(s) {', '.join(str(s.number) for s in chosen)}")
    if a.cmd != REPLY:
        print_status(a.deck)


def cmd_refresh(a) -> None:
    """Re-read the deck after it was rebuilt; an open review page picks the changes up by itself."""
    out, slides = inspect(a.deck)
    if not a.no_previews:
        previews(a.deck, out, slides)
    write_page(a.deck)
    print_status(a.deck)


def print_status(pptx: Path, as_json: bool = False) -> dict:
    slides = load_slides(output_dir(pptx))
    state = ReviewState(pptx)
    rows = [{"number": s.number, "title": s.title, "status": state.status(s),
             "comment": state.data["slides"].get(str(s.slide_id), {}).get("comment", "")} for s in slides]
    counts = {}
    for r in rows:
        counts[r["status"]] = counts.get(r["status"], 0) + 1
    if as_json:
        print(json.dumps({"deck": pptx.name, "counts": counts, "slides": rows}, indent=2, ensure_ascii=False))
        return counts
    for r in rows:
        note = f"  - {r['comment']}" if r["comment"] and r["status"] != "approved" else ""
        print(f"  {r['number']:>3}  {r['status']:<18} {r['title']}{note}")
    print("  " + ", ".join(f"{v} {k}" for k, v in sorted(counts.items())) + f"   ({state_path(pptx).name})")
    return counts


def cmd_rematch(a) -> None:
    from .review import rematch_recordings
    numbers = {s.number for s in parse_slides(a.slides, load_slides(output_dir(a.deck)))} if a.slides else None
    done = rematch_recordings(a.deck, numbers)
    log.info(f"Re-processed {len(done)} recording(s): " + ", ".join(f"slide {d['slide']}" for d in done)
             + ". A running review server rebuilds their clips (approved slides).")


def cmd_status(a) -> None:
    if not load_slides(output_dir(a.deck)):
        raise SystemExit("The deck has not been inspected yet: run `review` first")
    print_status(a.deck, a.json)


# ---------------------------------------------------------------- generate


def cmd_generate(a, ignore_approvals: bool = False) -> None:
    pptx = a.deck
    out, slides = inspect(pptx)
    state = ReviewState(pptx)
    if ignore_approvals:
        a.partial = True
    visible = [s for s in slides if not s.hidden]
    wanted = parse_slides(a.slides, slides) if a.slides else visible
    redo = {s.number for s in parse_slides(a.redo_audio, slides)} if a.redo_audio else set()
    todo = [s for s in wanted if not s.hidden and (ignore_approvals or state.is_approved(s))]
    skipped = [s for s in wanted if not s.hidden and s not in todo]
    if skipped:
        log.info("Not approved, skipped: " + ", ".join(f"{s.number} ({state.status(s)})" for s in skipped))

    _, failed = pipeline.render_slides(pptx, todo, pipeline.Options.from_args(a), redo)
    write_page(pptx)
    if failed:
        log.error("Narration failed for slide(s) " + ", ".join(str(s.number) for s in failed)
                  + ": run `generate` again to retry")
    if a.no_build:
        log.info("Skipping the final video (--no-build)")
    else:
        pipeline.build_final(pptx, partial=a.partial, ignore_approvals=ignore_approvals)
        write_page(pptx)
    if failed:
        raise SystemExit(1)


def cmd_run(a) -> None:
    cmd_generate(a, ignore_approvals=True)


# ---------------------------------------------------------------- main


def main(argv=None) -> None:
    p = argparse.ArgumentParser(prog="presgen", description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest="cmd", required=True)

    def deck_arg(sp):
        sp.add_argument("deck", type=Path, help="the .pptx file")

    def gen_args(sp):
        sp.add_argument("--renderer", choices=["powerpoint", "still"], default="powerpoint",
                        help="powerpoint plays animations and clips; still shows each slide's final state")
        sp.add_argument("--tts", default=render.DEFAULT_TTS,
                        help="TTS command with {text} and {audio} placeholders (run in the output folder)")
        sp.add_argument("--retries", type=int, default=3, help="TTS attempts per slide")
        # sp.add_argument("--tts-cache", default=render.DEFAULT_TTS_CACHE,
        #                 help="the TTS's sentence cache, cleared for a slide when it is re-voiced ('' to leave it)")
        sp.add_argument("--lead", type=float, default=1.0, help="seconds before the narration starts")
        sp.add_argument("--tail", type=float, default=0.75, help="seconds after the narration ends")
        sp.add_argument("--quality", type=int, default=100, help="PowerPoint video quality, 1-100")
        sp.add_argument("--wps", type=float, default=2.5,
                        help="words per second the animations were planned for (timing is matched to the voice)")
        sp.add_argument("--max-stretch", type=float, default=1.25,
                        help="largest timing stretch or squeeze; beyond it the slide is flagged for retiming")

    sp = sub.add_parser("review", help="export stills and previews, then open the review page")
    deck_arg(sp)
    sp.add_argument("--port", type=int, default=8765)
    sp.add_argument("--no-serve", action="store_true", help="write review.html and exit (no buttons)")
    sp.add_argument("--no-open", action="store_true", help="don't open a browser")
    sp.add_argument("--no-previews", action="store_true", help="skip the silent PowerPoint previews")
    sp.add_argument("--no-auto", action="store_true", help="don't narrate and render slides when approved")
    gen_args(sp)
    sp.set_defaults(func=cmd_review)

    sp = sub.add_parser("serve", help="serve the review page again without re-inspecting the deck")
    deck_arg(sp)
    sp.add_argument("--port", type=int, default=8765)
    sp.add_argument("--no-open", action="store_true")
    sp.add_argument("--no-auto", action="store_true", help="don't narrate and render slides when approved")
    gen_args(sp)
    sp.set_defaults(func=cmd_serve)

    sp = sub.add_parser("refresh", help="re-read the deck after a rebuild (a running review page updates itself)")
    deck_arg(sp)
    sp.add_argument("--no-previews", action="store_true")
    sp.set_defaults(func=cmd_refresh)

    for name, text in (("approve", "approve slides"), ("changes", "request changes to slides"),
                       ("reset", "set slides back to pending"),
                       (REPLY, "post a note to the reviewer on the review page (status unchanged)")):
        sp = sub.add_parser(name, help=text)
        deck_arg(sp)
        sp.add_argument("slides", help="slide numbers: 3, 1-4,7 or all")
        sp.add_argument("-m", "--message", help="comment (required for `changes` and `reply`)")
        sp.add_argument("--by", default="cli", help="who is acting (shown in the history)")
        sp.set_defaults(func=cmd_change)

    sp = sub.add_parser("rematch", help="re-process recorded narration so it matches the generated voice")
    deck_arg(sp)
    sp.add_argument("--slides", help="only these slides (default: every slide with a recording)")
    sp.set_defaults(func=cmd_rematch)

    sp = sub.add_parser("status", help="show each slide's review status")
    deck_arg(sp)
    sp.add_argument("--json", action="store_true")
    sp.set_defaults(func=cmd_status)

    for name, func, text in (("generate", cmd_generate, "narrate and render approved slides, then build the video"),
                             ("run", cmd_run, "narrate and render every slide, ignoring approvals")):
        sp = sub.add_parser(name, help=text)
        deck_arg(sp)
        sp.add_argument("--slides", help="only these slides (default: all approved)")
        sp.add_argument("--redo-audio", metavar="SLIDES", help="regenerate narration for these slides")
        gen_args(sp)
        sp.add_argument("--partial", action="store_true", help="build the video from whatever is ready")
        sp.add_argument("--no-build", action="store_true", help="make the clips but not the final video")
        sp.set_defaults(func=func)

    a = p.parse_args(argv)
    if not a.deck.exists() or a.deck.suffix.lower() not in (".pptx", ".ppt"):
        raise SystemExit(f"Not a PowerPoint file: {a.deck}")
    a.deck = a.deck.resolve()
    setup_logging(output_dir(a.deck))
    try:
        a.func(a)
    except ValueError as e:
        log.error(str(e))
        raise SystemExit(2)
