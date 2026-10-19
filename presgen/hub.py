"""The overview page: every deck in a folder, its review status, its final video, and a way into its review page.

    uv run presgen.py hub DECKS_DIR            # http://localhost:8760/

Each deck's review page is its own server (`presgen serve`), because each one owns a voice worker and
drives PowerPoint. The hub only reads state. "Open review" reuses a deck's running server (each server
records itself in <out>/server.json) or starts one on a free port, then opens it.
"""
from __future__ import annotations

import json
import logging
import os
import socket
import subprocess
import sys
import time
import webbrowser
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from .deck import load_slides
from .review import ReviewState, output_dir

log = logging.getLogger("presgen")
ROOT = Path(__file__).resolve().parent.parent
FIRST_REVIEW_PORT = 8765


def _deck_title(pptx: Path) -> str:
    try:
        from pptx import Presentation

        return Presentation(str(pptx)).core_properties.title or pptx.stem
    except Exception:
        return pptx.stem


def _pid_alive(pid: int) -> bool:
    if not pid:
        return False
    if os.name == "nt":
        r = subprocess.run(["tasklist", "/FI", f"PID eq {pid}", "/NH"], capture_output=True, text=True)
        return str(pid) in r.stdout
    try:
        os.kill(pid, 0)
        return True
    except OSError:
        return False


def _listening(port: int) -> bool:
    with socket.socket() as s:
        s.settimeout(0.3)
        return s.connect_ex(("127.0.0.1", port)) == 0


def running_server(pptx: Path) -> dict | None:
    """The deck's review server, if one is running (recorded by `serve` in <out>/server.json)."""
    f = output_dir(pptx) / "server.json"
    try:
        info = json.loads(f.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    if _pid_alive(info.get("pid", 0)) and _listening(info.get("port", 0)):
        return info
    return None


def _taken_ports(decks: list[Path]) -> set[int]:
    return {info["port"] for d in decks if (info := running_server(d))}


def deck_summary(pptx: Path, decks_for_titles: dict) -> dict:
    out = output_dir(pptx)
    counts: dict[str, int] = {}
    try:  # the same statuses the review page shows ("changed" is worked out from fingerprints)
        state = ReviewState(pptx)
        for sl in load_slides(out):
            if not sl.hidden:
                st = state.status(sl)
                counts[st] = counts.get(st, 0) + 1
    except (OSError, ValueError, KeyError):
        pass
    total = sum(counts.values())
    final = out / f"{out.name}_video.mp4"
    current = False
    secs = 0.0
    try:
        manifest = json.loads((out / "render.json").read_text(encoding="utf-8"))
        built = manifest.get("_final") or {}
        current = final.exists() and bool(built) and all(manifest.get(k, {}).get("fingerprint") == v for k, v in built.items()) \
            and len(built) == total and counts.get("approved", 0) == total
        secs = sum(manifest[k].get("secs", 0) for k in built if k in manifest)
    except (OSError, ValueError):
        pass
    youtube = None
    try:
        youtube = json.loads(pptx.with_suffix(".youtube.json").read_text(encoding="utf-8"))
    except (OSError, ValueError):
        pass
    server = running_server(pptx)
    return {
        "id": pptx.stem,
        "title": decks_for_titles.get(pptx.stem) or _deck_title(pptx),
        "slides": total,
        "counts": counts,
        "final": {"exists": final.exists(), "current": current, "secs": round(secs, 1),
                  "url": f"/video/{pptx.stem}" if final.exists() else None},
        "youtube": youtube,
        "review": {"url": server["url"] if server else None, "running": bool(server)},
    }


def start_review(pptx: Path, decks: list[Path]) -> str:
    """Start (or reuse) the deck's review server and return its URL."""
    if info := running_server(pptx):
        return info["url"]
    taken = _taken_ports(decks)
    port = next(p for p in range(FIRST_REVIEW_PORT, FIRST_REVIEW_PORT + 50) if p not in taken and not _listening(p))
    reviewed = (output_dir(pptx) / "review.html").exists()
    cmd = [sys.executable, str(ROOT / "presgen.py"), "serve" if reviewed else "review", str(pptx), "--port", str(port), "--no-open"]
    output_dir(pptx).mkdir(parents=True, exist_ok=True)
    logf = open(output_dir(pptx) / "hub-serve.log", "ab")
    flags = (subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.DETACHED_PROCESS) if os.name == "nt" else 0
    subprocess.Popen(cmd, cwd=str(ROOT), stdout=logf, stderr=subprocess.STDOUT, creationflags=flags)
    deadline = time.time() + (60 if reviewed else 900)
    while time.time() < deadline:
        if _listening(port):
            return f"http://localhost:{port}/"
        time.sleep(0.5)
    raise TimeoutError(f"the review server for {pptx.stem} did not start on port {port}")


def hub(folder: Path, port: int = 8760, open_browser: bool = True) -> None:
    folder = folder.resolve()
    page = (Path(__file__).with_name("hub_page.html")).read_bytes()
    titles: dict = {}

    def decks() -> list[Path]:
        return sorted(p for p in folder.glob("*.pptx") if not p.name.startswith("~$"))

    class Handler(SimpleHTTPRequestHandler):
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
            url = urlparse(self.path)
            if url.path in ("/", "/index.html"):
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(page)))
                self.end_headers()
                self.wfile.write(page)
                return
            if url.path == "/api/decks":
                found = decks()
                for d in found:
                    titles.setdefault(d.stem, _deck_title(d))
                return self._json(200, {"folder": str(folder), "decks": [deck_summary(d, titles) for d in found]})
            if url.path.startswith("/video/"):
                stem = url.path.split("/", 2)[2]
                pptx = folder / f"{stem}.pptx"
                video = output_dir(pptx) / f"{stem}_video.mp4"
                if not video.is_file():
                    return self.send_error(404)
                self.send_response(200)
                self._cors()
                self.send_header("Content-Type", "video/mp4")
                self.send_header("Content-Length", str(video.stat().st_size))
                self.end_headers()
                with open(video, "rb") as f:
                    try:
                        while chunk := f.read(1 << 16):
                            self.wfile.write(chunk)
                    except (ConnectionResetError, ConnectionAbortedError, BrokenPipeError):
                        pass
                return
            return self.send_error(404)

        def _cors(self):
            # Lets a browser page (for example YouTube Studio, during an upload driven by Claude in Chrome)
            # fetch a final video from this local server. Only the read-only /video route sends these.
            self.send_header("Access-Control-Allow-Origin", self.headers.get("Origin") or "*")
            self.send_header("Access-Control-Allow-Private-Network", "true")

        def do_OPTIONS(self):
            if not urlparse(self.path).path.startswith("/video/"):
                return self.send_error(404)
            self.send_response(204)
            self._cors()
            self.send_header("Access-Control-Allow-Methods", "GET")
            self.send_header("Access-Control-Allow-Headers", "*")
            self.end_headers()

        def do_POST(self):
            url = urlparse(self.path)
            if url.path != "/api/open":
                return self.send_error(404)
            stem = (parse_qs(url.query).get("deck") or [""])[0]
            pptx = folder / f"{stem}.pptx"
            if not pptx.is_file():
                return self._json(404, {"error": f"no deck {stem!r}"})
            try:
                return self._json(200, {"url": start_review(pptx, decks())})
            except Exception as e:
                return self._json(500, {"error": str(e)})

    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    url = f"http://localhost:{port}/"
    log.info(f"Overview page: {url}  ({len(decks())} decks in {folder}; Ctrl+C to stop)")
    if open_browser:
        webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
