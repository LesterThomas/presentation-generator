# Presentation Generator

Turns a PowerPoint deck into a narrated video. The speaker notes are read aloud in a cloned voice
(`csm-voice`), and PowerPoint itself renders the pictures, so entrance animations and embedded video
clips play just as they do in a slide show.

Every slide is reviewed on a local review page and switched from **Draft** to **Approved**. As
soon as a slide is approved, it is narrated and rendered in the background, so you can listen to it
a minute or two later. When every slide is approved, the final video rebuilds itself. If a slide
changes after it was approved, it goes back to needing approval.

Animations are planned for a narration speed of 2.5 words per second. Each slide's pictures are
stretched slightly to keep pace with the real voice, and narration and animation start together.

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
uv run presgen.py review "My Deck.pptx"     # review page; approving a slide narrates and renders it
```

That's the whole workflow. `generate` (below) does the same from the command line, without the page.

### 1. Review

`review` exports each slide as a 1920×1080 still, and has PowerPoint render a silent preview video
of each slide so you can check its animations. It then opens a review page at
http://localhost:8765. For each slide the page shows:

- the narrated clip (once generated), the silent preview, or the still,
- the narration text (speaker notes) with an estimated length,
- an **Approved / Draft** toggle. Switching it on queues the slide for narration and rendering.
  Progress shows on the page: queued, narrating, rendering, then "narrated clip ready".
- the clip's timing: the real narration length, the planned length, and how much the animation was
  stretched to match. A slide that needs more than ×1.25 is flagged for re-timing at source.
- **Re-voice** (a fresh take of the narration, for a bad take or a mispronunciation that a retry
  fixes) and **Retry** (after a failure),
- **Request changes** with a comment.

Keys: ← → move between slides, A toggles approval, N jumps to the next slide that needs review. Stop
the server with Ctrl+C. `serve` reopens the page without re-reading the deck. `--no-auto` turns
background generation off.

Generation runs one slide at a time, because the voice needs the GPU. Switching a slide back to Draft
keeps its clip, so approving it again is instant. Don't run `generate` while a review server with
auto-generation is running for the same deck.

You can also review from the command line:

```bash
uv run presgen.py approve "My Deck.pptx" 1-4,7
uv run presgen.py changes "My Deck.pptx" 5 -m "Say 'O-D-A', not 'oda'"
uv run presgen.py reset   "My Deck.pptx" 5
uv run presgen.py status  "My Deck.pptx"          # add --json for scripts
uv run presgen.py reply   "My Deck.pptx" 5 -m "Fixed the pronunciation"   # note shown on the page
uv run presgen.py refresh "My Deck.pptx"          # re-read the deck after editing it
```

Slide numbers are PowerPoint's own numbers. Hidden slides keep their number and are skipped.

### Getting change requests actioned straight away

Every review action is appended as one JSON line to `My Deck/review-events.jsonl`. A watcher (in the
ODA Canvas training-videos project, Claude follows this file) can pick up each **Request changes**
the moment it is made. It then fixes the slide, rebuilds the deck, and reports back:

```bash
uv run presgen.py refresh "My Deck.pptx"                     # re-read the rebuilt deck
uv run presgen.py reply   "My Deck.pptx" 5 -m "Changed ..." --by claude
```

`refresh` re-exports stills and previews (only changed slides are re-rendered). An open review page
picks the result up by itself within a few seconds. A slide that changed after you asked for changes
shows as **Updated: review again**, with the reply under your comment. Approve it, or ask again.

Approvals are saved next to the deck in `My Deck.review.json`, which you can commit alongside the
deck. Each approval stores the slide's *fingerprint*: its still image, notes, embedded media and
animation timings. Change any of these and the slide shows as **changed since approval**. It is then
left out of the video until you approve it again.

### 2. Generate (command line)

The review server does this by itself as slides are approved. `generate` is the same pipeline, as a
batch, for approved slides only:

1. **Narration.** It runs the TTS on `text_NN.txt` to make `audio_NN.wav`. An audio file is reused
   while it is newer than its text. A failed run (csm-voice sometimes crashes) is retried 3 times.
2. **Pictures.** PowerPoint renders each slide on its own (`Presentation.CreateVideo`, 1080p, 30 fps),
   long enough for the narration, every animation and any embedded clip. That takes 2–3 seconds per
   slide.
3. **Timing and sound.** The narration starts 1 second in, and so does the slide's animation (the
   first frame holds until then). The animation is planned at 2.5 words per second (`--wps`). It is
   stretched or squeezed by real narration ÷ planned narration so each build lands on its words,
   clamped at ×1.25 (`--max-stretch`). The log and the page report it. FFmpeg then mixes the
   narration with any sound the slide already has, giving `clips/clip_NN.mp4`. A clip is reused while
   its slide, audio and settings are unchanged (`render.json`).
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
| `--wps 2.5` / `--max-stretch 1.25` | planned narration speed, and the largest timing stretch |
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
  review-events.jsonl        one line per review action (for watchers)
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
| `presgen/cli.py` | commands: review, serve, refresh, approve, changes, reset, reply, status, generate, run |
| `presgen/deck.py` | PowerPoint COM: stills, notes, animation timings, fingerprints |
| `presgen/review.py` | approval state, review page data, local review server |
| `presgen/review_page.html` | the review page |
| `presgen/render.py` | TTS, timing, PowerPoint rendering, FFmpeg muxing and concatenation |
| `presgen/pipeline.py` | narrate and render slides, and build the final video (shared by CLI and server) |
| `presgen/autogen.py` | the review server's background generator |

## License

MIT
