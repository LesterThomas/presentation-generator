"""presgen: review a narrated deck slide by slide, then turn the approved slides into a video.

    uv run presgen.py review   DECK.pptx            # stills + silent previews, opens the review page
    uv run presgen.py approve  DECK.pptx 1-4,7      # or approve on the review page
    uv run presgen.py changes  DECK.pptx 5 -m "..." # request changes, with a comment
    uv run presgen.py status   DECK.pptx
    uv run presgen.py generate DECK.pptx            # narration + clips for approved slides, then the video
    uv run presgen.py run      DECK.pptx            # everything, ignoring approvals (the old behaviour)
"""

import argparse
import json
import logging
import sys
from pathlib import Path

from . import render
from .deck import inspect_deck, load_slides
from .review import (ACTIONS, ReviewState, change_status, output_dir, parse_slides, serve,
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
    serve(a.deck, a.port, open_browser=not a.no_open)


def cmd_serve(a) -> None:
    serve(a.deck, a.port, open_browser=not a.no_open)


def cmd_change(a) -> None:
    if a.cmd == "changes" and not a.message:
        raise SystemExit("`changes` needs a comment: -m \"what to change\"")
    chosen = change_status(a.deck, a.slides, a.cmd, a.message or "")
    write_page(a.deck)
    log.info(f"{ACTIONS[a.cmd]}: slide(s) {', '.join(str(s.number) for s in chosen)}")
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

    # 1. narration (the slow part: run only for approved slides)
    audio, failed = {}, []
    for s in todo:
        audio[s.number] = render.synthesize(out, s, a.tts, a.retries, force=s.number in redo)
        if s.notes.strip() and audio[s.number] is None:
            failed.append(s)

    # 2. pictures: PowerPoint renders each slide alone; unchanged clips are reused
    manifest = render.Manifest(out)
    jobs, plan = [], []
    for s in todo:
        if s in failed:
            continue
        secs = render.slide_secs(s, audio[s.number], a.lead, a.tail)
        key = manifest.key(s, audio[s.number], a.renderer, secs, a.lead)
        clip = f"clips/clip_{s.stem}.mp4"
        plan.append((s, secs, key, clip))
        if manifest.current(out, s, key):
            log.info(f"Slide {s.number}: clip up to date")
            continue
        jobs.append((s, secs, out / ".work" / f"silent_{s.stem}.mp4"))
    if jobs:
        if a.renderer == "powerpoint":
            render.powerpoint_videos(out / ".work" / "deck.pptx", jobs, quality=a.quality)
        else:
            for s, secs, silent in jobs:
                render.still_video(out / f"slide_{s.stem}.png", secs, silent)
    fresh = {s.number for s, _, _ in jobs}
    for s, secs, key, clip in plan:
        if s.number in fresh:
            silent = out / ".work" / f"silent_{s.stem}.mp4"
            render.add_narration(silent, audio[s.number], secs, a.lead, out / clip)
            silent.unlink(missing_ok=True)
            manifest.record(s, key, clip, audio[s.number], secs, a.renderer)
            manifest.save()
            log.info(f"Slide {s.number}: clip ready ({secs:.1f}s)")
    manifest.save()
    write_page(pptx)

    # 3. the whole video, once every visible slide is approved and rendered
    ready = {s.number: (secs, clip) for s, secs, _, clip in plan}
    for s in visible:
        e = manifest.data.get(s.stem, {})
        if s.number not in ready and e.get("fingerprint") == s.fingerprint and state.is_approved(s) \
                and (out / e.get("clip", "-")).exists():
            ready[s.number] = (e["secs"], e["clip"])
    missing = [s for s in visible if s.number not in ready]
    if failed:
        log.error("Narration failed for slide(s) " + ", ".join(str(s.number) for s in failed)
                  + ": run `generate` again to retry")
    if a.no_build:
        log.info("Skipping the final video (--no-build)")
    elif missing and not a.partial:
        if (out / f"{out.name}_video.mp4").exists():
            log.warning("The existing final video is out of date")
        log.info(f"Final video not built: {len(missing)} slide(s) not approved or not rendered yet: "
                 + ", ".join(str(s.number) for s in missing) + "  (use --partial to build from the rest)")
    else:
        parts = [s for s in visible if s.number in ready]
        final = out / f"{out.name}_video.mp4"
        render.concat([out / ready[s.number][1] for s in parts], final)
        secs = [ready[s.number][0] for s in parts]
        (out / f"{out.name}_chapters.txt").write_text(render.chapters(parts, secs), encoding="utf-8")
        manifest.data["_final"] = {s.stem: s.fingerprint for s in parts}
        manifest.save()
        total = sum(secs)
        log.info(f"Final video: {final}  ({len(parts)} slides, {int(total // 60)}:{int(total % 60):02d})")
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

    sp = sub.add_parser("review", help="export stills and previews, then open the review page")
    deck_arg(sp)
    sp.add_argument("--port", type=int, default=8765)
    sp.add_argument("--no-serve", action="store_true", help="write review.html and exit (no buttons)")
    sp.add_argument("--no-open", action="store_true", help="don't open a browser")
    sp.add_argument("--no-previews", action="store_true", help="skip the silent PowerPoint previews")
    sp.set_defaults(func=cmd_review)

    sp = sub.add_parser("serve", help="serve the review page again without re-inspecting the deck")
    deck_arg(sp)
    sp.add_argument("--port", type=int, default=8765)
    sp.add_argument("--no-open", action="store_true")
    sp.set_defaults(func=cmd_serve)

    for name, text in (("approve", "approve slides"), ("changes", "request changes to slides"),
                       ("reset", "set slides back to pending")):
        sp = sub.add_parser(name, help=text)
        deck_arg(sp)
        sp.add_argument("slides", help="slide numbers: 3, 1-4,7 or all")
        sp.add_argument("-m", "--message", help="comment (required for `changes`)")
        sp.set_defaults(func=cmd_change)

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
        sp.add_argument("--renderer", choices=["powerpoint", "still"], default="powerpoint",
                        help="powerpoint plays animations and clips; still shows each slide's final state")
        sp.add_argument("--tts", default=render.DEFAULT_TTS,
                        help="TTS command with {text} and {audio} placeholders (run in the output folder)")
        sp.add_argument("--retries", type=int, default=3, help="TTS attempts per slide")
        sp.add_argument("--lead", type=float, default=1.0, help="seconds before the narration starts")
        sp.add_argument("--tail", type=float, default=0.75, help="seconds after the narration ends")
        sp.add_argument("--quality", type=int, default=100, help="PowerPoint video quality, 1-100")
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
