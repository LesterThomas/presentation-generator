"""
PowerPoint presentation to narrated video, in one step (no review).

Kept for the original command line. It is the same as `presgen.py run`: every visible slide is
narrated from its speaker notes and rendered by PowerPoint (animations and embedded clips play),
whether or not it has been approved. For slide-by-slide review, use presgen.py (see README.md).

Usage:
    uv run extract_slides.py <path_to_presentation.pptx> [run options]
"""

import sys

from presgen.cli import main

if __name__ == "__main__":
    main(["run", *sys.argv[1:]])
