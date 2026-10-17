"""Slide-by-slide review: approval state, the review page, and a small local server for it.

Approvals live next to the deck in ``<deck>.review.json``, so they can be committed with it. Each
approval records the slide's fingerprint. If the slide changes afterwards (visuals, notes, media or
animation timings), the approval no longer counts and the slide shows as "changed" until it is
approved again.
"""

import json
import logging
import mimetypes
import re
import threading
import webbrowser
from datetime import datetime, timezone
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from .deck import load_slides, write_atomic

log = logging.getLogger("presgen")

ACTIONS = {"approve": "approved", "changes": "changes-requested", "reset": "pending"}
PAGE_ACTIONS = set(ACTIONS)                 # what the review page may do
REPLY = "reply"                             # a note back to the reviewer; status unchanged
PAGE_TEMPLATE = Path(__file__).with_name("review_page.html")
_lock = threading.Lock()


def state_path(pptx: Path) -> Path:
    return pptx.with_name(pptx.stem + ".review.json")


def output_dir(pptx: Path) -> Path:
    return pptx.parent / pptx.stem


def _now() -> str:
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")


class ReviewState:
    def __init__(self, pptx: Path):
        self.pptx = pptx
        self.path = state_path(pptx)
        self.data = {"deck": pptx.name, "slides": {}}
        if self.path.exists():
            self.data = json.loads(self.path.read_text(encoding="utf-8"))

    def save(self) -> None:
        write_atomic(self.path, json.dumps(self.data, indent=2, ensure_ascii=False) + "\n")

    def entry(self, slide_id) -> dict:
        return self.data["slides"].setdefault(str(slide_id), {"status": "pending", "history": []})

    def sync(self, slides) -> None:
        """Record the current number, title and fingerprint of every slide after an inspection."""
        present = set()
        for s in slides:
            present.add(str(s.slide_id))
            e = self.entry(s.slide_id)
            e.update(number=s.number, title=s.title, hidden=s.hidden, fingerprint=s.fingerprint)
        for sid in list(self.data["slides"]):
            if sid not in present:
                del self.data["slides"][sid]  # the slide was deleted from the deck
        self.data["inspected_at"] = _now()
        self.save()

    def status(self, slide) -> str:
        """approved | changed | changes-requested | updated | pending | hidden"""
        if slide.hidden:
            return "hidden"
        e = self.data["slides"].get(str(slide.slide_id), {})
        st = e.get("status", "pending")
        if st == "approved" and e.get("approved_fingerprint") != slide.fingerprint:
            return "changed"
        asked = e.get("requested_fingerprint") or next(
            (h["fingerprint"] for h in reversed(e.get("history", [])) if h["action"] == "changes"), None)
        if st == "changes-requested" and asked not in (None, slide.fingerprint):
            return "updated"  # the slide changed after the reviewer asked for changes
        return st

    def is_approved(self, slide) -> bool:
        return self.status(slide) == "approved"

    def apply(self, slide, action: str, comment: str = "", by: str = "cli") -> None:
        if slide.hidden:
            raise ValueError(f"slide {slide.number} is hidden")
        if not slide.fingerprint:
            raise ValueError(f"slide {slide.number} has not been inspected yet: run `review` first")
        e = self.entry(slide.slide_id)
        if action == REPLY:
            e["reply"] = {"at": _now(), "text": comment}
            e["history"].append({"at": e["reply"]["at"], "action": REPLY, "by": by,
                                 "comment": comment, "fingerprint": slide.fingerprint})
            return
        e.pop("reply", None)
        e["status"] = ACTIONS[action]
        e["comment"] = comment
        e["updated_at"] = _now()
        e.pop("approved_fingerprint", None)
        e.pop("requested_fingerprint", None)
        if action == "approve":
            e["approved_fingerprint"] = slide.fingerprint
        elif action == "changes":
            e["requested_fingerprint"] = slide.fingerprint
        e["history"].append({"at": e["updated_at"], "action": action, "by": by,
                             "comment": comment, "fingerprint": slide.fingerprint})


def parse_slides(spec: str, slides) -> list:
    """'all', '3', '1-4,7' -> SlideInfo list (by PowerPoint slide number)."""
    visible = [s for s in slides if not s.hidden]
    if spec in ("all", "*"):
        return visible
    numbers = set()
    for part in re.split(r"\s*,\s*", spec.strip()):
        if not part:
            continue
        a, _, b = part.partition("-")
        numbers.update(range(int(a), int(b or a) + 1))
    by_num = {s.number: s for s in slides}
    unknown = sorted(n for n in numbers if n not in by_num)
    if unknown:
        raise ValueError(f"no such slide(s): {unknown}")
    return [by_num[n] for n in sorted(numbers)]


def change_status(pptx: Path, spec: str, action: str, comment: str = "", by: str = "cli") -> list:
    with _lock:
        slides = load_slides(output_dir(pptx))
        if not slides:
            raise ValueError("the deck has not been inspected yet: run `review` first")
        state = ReviewState(pptx)
        chosen = [s for s in parse_slides(spec, slides) if not s.hidden]
        if not chosen:
            raise ValueError(f"no visible slides in {spec!r}")
        for s in chosen:
            state.apply(s, action, comment, by)
        state.save()
        event = {"at": _now(), "action": action, "slides": [s.number for s in chosen],
                 "titles": [s.title for s in chosen], "comment": comment, "by": by}
        with open(events_path(pptx), "a", encoding="utf-8") as f:
            f.write(json.dumps(event, ensure_ascii=False) + "\n")
        return chosen


def events_path(pptx: Path) -> Path:
    """One JSON line per review action: what a watcher (e.g. Claude) follows to act on requests."""
    return output_dir(pptx) / "review-events.jsonl"


# ---------------------------------------------------------------- page data


def page_data(pptx: Path, gen=None) -> dict:
    """Everything the review page shows, built from the output folder and the review state.

    `gen` is the review server's AutoGenerator, if any: its job progress is included.
    """
    out = output_dir(pptx)
    live = gen.snapshot() if gen else {"jobs": {}, "final": {}}
    state = ReviewState(pptx)
    render = {}
    if (out / "render.json").exists():
        render = json.loads((out / "render.json").read_text(encoding="utf-8"))
    rows = []
    slides = load_slides(out)
    for s in slides:
        e = state.data["slides"].get(str(s.slide_id), {})
        r = render.get(s.stem, {})
        clip = r.get("clip") if (r.get("fingerprint") == s.fingerprint and "timing" in r
                                 and (out / r.get("clip", "-")).exists()) else None
        preview = f"previews/slide_{s.stem}.mp4"
        rows.append({
            "number": s.number, "slideId": s.slide_id, "title": s.title, "notes": s.notes,
            "hidden": s.hidden, "status": state.status(s), "comment": e.get("comment", ""),
            "updatedAt": e.get("updated_at", ""), "history": e.get("history", [])[-6:],
            "reply": e.get("reply"),
            "image": f"slide_{s.stem}.png" if not s.hidden else None,
            "preview": preview if (out / preview).exists() else None,
            "clip": clip, "audio": r.get("audio") if clip else None,
            "secs": r.get("secs") if clip else None,
            "timing": r.get("timing") if clip else None,
            "job": live["jobs"].get(s.stem),
        })
    final = out / f"{out.name}_video.mp4"
    built = render.get("_final", {})
    current = {s.stem: s.fingerprint for s in slides if not s.hidden}
    return {"deck": pptx.name, "slides": rows, "inspectedAt": state.data.get("inspected_at", ""),
            "video": final.name if final.exists() else None, "videoCurrent": built == current,
            "auto": gen is not None, "final": live["final"]}


def write_page(pptx: Path) -> Path:
    out = output_dir(pptx)
    html = PAGE_TEMPLATE.read_text(encoding="utf-8")
    data = json.dumps(page_data(pptx), ensure_ascii=False).replace("</", "<\\/")
    page = out / "review.html"
    page.write_text(html.replace("/*__DATA__*/null", data), encoding="utf-8")
    return page


# ---------------------------------------------------------------- server


def serve(pptx: Path, port: int = 8765, open_browser: bool = True, gen=None) -> None:
    """Serve the review page with live review controls until Ctrl+C.

    With `gen` (an AutoGenerator), approving a slide narrates and renders it straight away.
    """
    out = output_dir(pptx)

    class Handler(SimpleHTTPRequestHandler):
        def __init__(self, *a, **kw):
            super().__init__(*a, directory=str(out), **kw)

        def log_message(self, fmt, *args):
            pass

        def _json(self, code: int, body) -> None:
            raw = json.dumps(body, ensure_ascii=False).encode()
            self.send_response(code)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(raw)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(raw)

        def do_GET(self):
            if self.path.startswith("/api/state"):
                return self._json(200, page_data(pptx, gen))
            if self.path in ("/", "/index.html"):
                self.path = "/review.html"
            if "Range" in self.headers:
                return self._range()
            return super().do_GET()

        def end_headers(self):
            self.send_header("Accept-Ranges", "bytes")
            super().end_headers()

        def _range(self):
            """Byte ranges, so the browser can seek inside the slide videos."""
            path = Path(self.translate_path(self.path))
            if not path.is_file():
                return self.send_error(404)
            size = path.stat().st_size
            m = re.match(r"bytes=(\d*)-(\d*)", self.headers["Range"])
            start = int(m.group(1)) if m and m.group(1) else 0
            end = int(m.group(2)) if m and m.group(2) else size - 1
            end = min(end, size - 1)
            if start > end:
                self.send_response(416)
                self.send_header("Content-Range", f"bytes */{size}")
                return self.end_headers()
            self.send_response(206)
            self.send_header("Content-Type", mimetypes.guess_type(str(path))[0] or "application/octet-stream")
            self.send_header("Content-Range", f"bytes {start}-{end}/{size}")
            self.send_header("Content-Length", str(end - start + 1))
            self.end_headers()
            with open(path, "rb") as f:
                f.seek(start)
                left = end - start + 1
                while left > 0:
                    chunk = f.read(min(1 << 16, left))
                    if not chunk:
                        break
                    try:
                        self.wfile.write(chunk)
                    except (ConnectionResetError, BrokenPipeError):
                        return
                    left -= len(chunk)

        def do_POST(self):
            if self.path != "/api/review":
                return self.send_error(404)
            try:
                body = json.loads(self.rfile.read(int(self.headers.get("Content-Length", 0))) or b"{}")
                action = body.get("action")
                if action in ("revoice", "retry"):
                    if not gen:
                        raise ValueError("auto-generation is off: use `presgen.py generate`")
                    for n in str(body["slides"]).split(","):
                        gen.request(f"{int(n):02d}", revoice=action == "revoice")
                    log.info(f"Review page: {action} slide(s) {body['slides']}")
                    return self._json(200, page_data(pptx, gen))
                if action not in PAGE_ACTIONS:
                    raise ValueError(f"unknown action {action!r}")
                chosen = change_status(pptx, str(body["slides"]), action, body.get("comment", ""), by="review page")
                log.info(f"Review page: {action} slide(s) {', '.join(str(s.number) for s in chosen)}"
                         + (f" - {body.get('comment')}" if body.get("comment") else ""))
                write_page(pptx)
                if gen:
                    gen.poke()
                return self._json(200, page_data(pptx, gen))
            except Exception as e:
                return self._json(400, {"error": str(e)})

    write_page(pptx)
    if gen:
        gen.start()
    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    url = f"http://localhost:{port}/"
    log.info(f"Review page: {url}  (Ctrl+C to stop)")
    if open_browser:
        webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
