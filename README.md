# Presentation Generator

Turns a PowerPoint deck into a narrated video. The speaker notes are read aloud in a cloned voice
(`csm-voice`), and PowerPoint itself renders the pictures, so entrance animations and embedded video
clips play just as they do in a slide show.

Every slide is reviewed and approved first. Only approved slides are narrated and rendered, which
matters because the voice takes one to three minutes per slide. If a slide changes after it was
approved, it has to be approved again.

## Requirements

- Windows with Microsoft PowerPoint (used through COM)
- Python 3.8+ and [uv](https://github.com/astral-sh/uv)
- [`csm-voice`](https://github.com/LesterThomas/csm-lester-voice) for the narration (any other
  command-line TTS works too, see `--tts`)

```bash
uv sync
```

Behind a TLS-intercepting proxy or antivirus (such as Norton), add `--native-tls` to `uv` commands.

## Workflow

```bash
uv run presgen.py review   "My Deck.pptx"               # 1. review page in the browser
uv run presgen.py generate "My Deck.pptx"               # 2. narrate + render approved slides
uv run presgen.py review   "My Deck.pptx"               # 3. listen, approve the rest, repeat 2
```

### 1. Review

`review` exports each slide as a 1920×1080 still, and has PowerPoint render a silent preview video
of each slide so you can check its animations. It then opens a review page at
http://localhost:8765. For each slide the page shows:

- the narrated clip (once generated), the silent preview, or the still,
- the narration text (speaker notes) with an estimated length,
- **Approve & next**, **Request changes** (with a comment) and **Reset**.

Keys: ← → move between slides, A approves, N jumps to the next slide that needs review. Stop the
server with Ctrl+C. `serve` reopens the page without re-reading the deck.

You can also review from the command line:

```bash
uv run presgen.py approve "My Deck.pptx" 1-4,7
uv run presgen.py changes "My Deck.pptx" 5 -m "Say 'O-D-A', not 'oda'"
uv run presgen.py reset   "My Deck.pptx" 5
uv run presgen.py status  "My Deck.pptx"          # add --json for scripts
```

Slide numbers are PowerPoint's own numbers. Hidden slides keep their number and are skipped.

Approvals are saved next to the deck in `My Deck.review.json`, which you can commit alongside the
deck. Each approval stores the slide's *fingerprint*: its still image, notes, embedded media and
animation timings. Change any of these and the slide shows as **changed since approval**. It is then
left out of the video until you approve it again.

### 2. Generate

`generate` works on approved slides only:

1. **Narration.** It runs the TTS on `text_NN.txt` to make `audio_NN.wav`. An audio file is reused
   while it is newer than its text. A failed run (csm-voice sometimes crashes) is retried 3 times.
2. **Pictures.** PowerPoint renders each slide on its own (`Presentation.CreateVideo`, 1080p, 30 fps),
   long enough for the narration, every animation and any embedded clip. That takes 2–3 seconds per
   slide.
3. **Sound.** FFmpeg adds the narration, starting 1 second in, mixed with any sound the slide already
   has. The result is `clips/clip_NN.mp4`. A clip is reused while its slide, audio and settings are
   unchanged (`render.json`).
4. **Final video.** Once every visible slide is approved and rendered, the clips are joined into
   `My Deck_video.mp4`. It also writes `My Deck_chapters.txt`, a YouTube-style chapter list built from
   the slide titles.

Useful options:

| Option | Effect |
|---|---|
| `--slides 3-5` | only these slides |
| `--redo-audio 4` | regenerate slide 4's narration (e.g. after a mispronunciation) |
| `--partial` | build the final video from whatever is ready |
| `--no-build` | make the clips but not the final video |
| `--renderer still` | show each slide's final state as a still (no PowerPoint animation) |
| `--lead 1.0` / `--tail 0.75` | seconds of silence before / after the narration |
| `--tts "cmd -f {text} -o {audio}"` | a different TTS command (run in the output folder) |

### One step, no review

```bash
uv run extract_slides.py "My Deck.pptx"     # same as: uv run presgen.py run "My Deck.pptx"
```

This narrates and renders every visible slide without waiting for approvals, which is the original
behaviour.

## Output

Everything goes into a folder named after the deck, next to it:

```
My Deck.pptx
My Deck.review.json          approvals (keep with the deck)
My Deck/
  review.html                review page (read-only if opened as a file)
  slide_NN.png  text_NN.txt  audio_NN.wav
  previews/slide_NN.mp4      silent PowerPoint renders for review
  clips/clip_NN.mp4          narrated clip per slide
  render.json  slides.json   caches
  My Deck_video.mp4          the final video
  My Deck_chapters.txt
  presgen.log
```

The deck itself is never changed. It is copied to `My Deck/.work/` and rendered from there, so it
can stay open in PowerPoint while you review it.

## Writing notes for the voice

- Write the narration as full spoken sentences. csm-voice generates one sentence at a time and splits
  on `.`, `!`, `?`, `;`, em-dashes and new lines, so avoid version numbers like `1.2.6` (say "one point
  two") and other mid-sentence full stops.
- Add `[pause]` for a one-second pause (`[pause][pause]` for two).
- Spell acronyms the way they should be said if the voice gets them wrong, e.g. "O-D-A".

## Configuration

The default TTS command is

```
"D:\Dev\lesterthomas\csm-lester-voice\.venv\Scripts\csm-voice.exe" -f {text} -o {audio}
```

It is set in `presgen/render.py` (`DEFAULT_TTS`). Override it per run with `--tts`.

## Code

| File | Purpose |
|---|---|
| `presgen/cli.py` | commands: review, serve, approve, changes, reset, status, generate, run |
| `presgen/deck.py` | PowerPoint COM: stills, notes, animation timings, fingerprints |
| `presgen/review.py` | approval state, review page data, local review server |
| `presgen/review_page.html` | the review page |
| `presgen/render.py` | TTS, PowerPoint rendering, FFmpeg muxing and concatenation |

## License

MIT
