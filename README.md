# Presentation Generator

A Python script that extracts slides from PowerPoint presentations as PNG images, extracts speaker notes, and generates audio narrations using text-to-speech.

## Features

- **Slide Export**: Exports each slide as a high-quality PNG image (`slide_01.png`, `slide_02.png`, etc.)
- **Speaker Notes Extraction**: Saves speaker notes to individual text files (`text_01.txt`, `text_02.txt`, etc.)
- **Audio Generation**: Creates audio narrations from speaker notes using `csm-voice` tool (`audio_01.wav`, `audio_02.wav`, etc.)
- **Error Logging**: Comprehensive logging to both console and `error.log` file
- **Automated Workflow**: Processes everything in one command

## Requirements

- Python 3.8+
- [uv](https://github.com/astral-sh/uv) package manager
- Microsoft PowerPoint (Windows)
- `csm-voice` tool for audio generation

## Installation

1. Clone this repository:
```bash
git clone https://github.com/LesterThomas/presentation-generator.git
cd presentation-generator
```

2. Install dependencies using uv:
```bash
uv sync
```

## Usage

Run the script with your PowerPoint presentation:

```bash
uv run extract_slides.py "path/to/your/presentation.pptx"
```

The script will:
1. Create a folder named after your presentation (without extension)
2. Export all slides as PNG images
3. Extract speaker notes to text files
4. Generate audio narrations from the speaker notes

### Example

```bash
uv run extract_slides.py "My Presentation.pptx"
```

This creates a `My Presentation` folder containing:
- `slide_01.png`, `slide_02.png`, ... (slide images)
- `text_01.txt`, `text_02.txt`, ... (speaker notes)
- `audio_01.wav`, `audio_02.wav`, ... (audio narrations)

## Configuration

The script uses a hardcoded path to `csm-voice`. To update it, edit line 148 in `extract_slides.py`:

```python
[r"D:\Dev\lesterthomas\csm-lester-voice\.venv\Scripts\csm-voice.exe", "-f", text_file, "-o", audio_file]
```

## Dependencies

- **comtypes**: Windows COM automation for PowerPoint
- **python-pptx**: Reading PowerPoint files and extracting speaker notes

## License

MIT
