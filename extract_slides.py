"""
PowerPoint Slide Extractor

This script extracts slides from a PowerPoint presentation as PNG images
and saves speaker notes as text files.

Usage:
    python extract_slides.py <path_to_presentation.pptx>
"""

import argparse
import logging
import os
import subprocess
import sys
from pathlib import Path
from typing import Optional

import comtypes.client
from pptx import Presentation


# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('error.log'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)


def setup_output_folder(pptx_path: Path) -> Path:
    """
    Create output folder with the same name as the presentation.
    
    Args:
        pptx_path: Path to the PowerPoint file
        
    Returns:
        Path to the output folder
    """
    output_folder = pptx_path.parent / pptx_path.stem
    
    # Create folder (overwrite contents if exists)
    output_folder.mkdir(parents=True, exist_ok=True)
    logger.info(f"Output folder: {output_folder}")
    
    return output_folder


def export_slides_as_png(pptx_path: Path, output_folder: Path) -> int:
    """
    Export all slides from PowerPoint presentation as PNG images using COM automation.
    
    Args:
        pptx_path: Path to the PowerPoint file
        output_folder: Path to the output folder
        
    Returns:
        Number of slides exported
    """
    powerpoint = None
    presentation = None
    slide_count = 0
    
    try:
        # Convert to absolute path for COM
        abs_pptx_path = str(pptx_path.resolve())
        
        logger.info("Starting PowerPoint application...")
        powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
        powerpoint.Visible = 1  # Make PowerPoint visible
        
        logger.info(f"Opening presentation: {abs_pptx_path}")
        presentation = powerpoint.Presentations.Open(abs_pptx_path, WithWindow=False)
        
        slide_count = presentation.Slides.Count
        logger.info(f"Found {slide_count} slides")
        
        # Export each slide as PNG
        for i in range(1, slide_count + 1):
            try:
                slide = presentation.Slides(i)
                output_file = output_folder / f"slide_{i:02d}.png"
                
                logger.info(f"Exporting slide {i}/{slide_count} to {output_file.name}")
                slide.Export(str(output_file.resolve()), "PNG")
                
            except Exception as e:
                logger.error(f"Failed to export slide {i}: {e}")
        
        logger.info("Slide export completed")
        
    except Exception as e:
        logger.error(f"Error during PowerPoint automation: {e}")
        raise
        
    finally:
        # Clean up COM objects
        try:
            if presentation:
                presentation.Close()
                logger.info("Presentation closed")
        except Exception as e:
            logger.error(f"Error closing presentation: {e}")
            
        try:
            if powerpoint:
                powerpoint.Quit()
                logger.info("PowerPoint application closed")
        except Exception as e:
            logger.error(f"Error quitting PowerPoint: {e}")
    
    return slide_count


def extract_speaker_notes(pptx_path: Path, output_folder: Path, slide_count: int) -> None:
    """
    Extract speaker notes from each slide and save as text files.
    
    Args:
        pptx_path: Path to the PowerPoint file
        output_folder: Path to the output folder
        slide_count: Number of slides in the presentation
    """
    try:
        logger.info("Extracting speaker notes...")
        prs = Presentation(str(pptx_path))
        
        for i, slide in enumerate(prs.slides, start=1):
            try:
                # Get speaker notes
                notes_slide = slide.notes_slide
                notes_text = notes_slide.notes_text_frame.text if notes_slide else ""
                
                # Save to text file
                output_file = output_folder / f"text_{i:02d}.txt"
                output_file.write_text(notes_text, encoding='utf-8')
                
                if notes_text.strip():
                    logger.info(f"Extracted notes for slide {i}/{slide_count}")
                else:
                    logger.info(f"No notes found for slide {i}/{slide_count}")
                    
            except Exception as e:
                logger.error(f"Failed to extract notes for slide {i}: {e}")
        
        logger.info("Speaker notes extraction completed")
        
    except Exception as e:
        logger.error(f"Error reading presentation for notes: {e}")
        raise


def generate_audio_from_notes(output_folder: Path, slide_count: int) -> None:
    """
    Generate audio files from speaker notes text files using csm-voice tool.
    
    Args:
        output_folder: Path to the output folder containing text files
        slide_count: Number of slides in the presentation
    """
    try:
        logger.info("Generating audio from speaker notes...")
        
        for i in range(1, slide_count + 1):
            text_file = f"text_{i:02d}.txt"
            audio_file = f"audio_{i:02d}.wav"
            
            try:
                logger.info(f"Processing {text_file} -> {audio_file}")
                
                # Set up environment with UTF-8 encoding for csm-voice
                env = os.environ.copy()
                env['PYTHONIOENCODING'] = 'utf-8'
                
                # Run csm-voice from the output folder
                result = subprocess.run(
                    [r"D:\Dev\lesterthomas\csm-lester-voice\.venv\Scripts\csm-voice.exe", "-f", text_file, "-o", audio_file],
                    cwd=str(output_folder),
                    capture_output=True,
                    text=True,
                    check=True,
                    env=env
                )
                
                logger.info(f"Successfully generated {audio_file}")
                
            except subprocess.CalledProcessError as e:
                logger.error(f"Failed to generate audio for {text_file}: {e}")
                if e.stderr:
                    logger.error(f"Error output: {e.stderr}")
            except FileNotFoundError:
                logger.error("csm-voice tool not found. Please ensure it is installed and in your PATH.")
                break
            except Exception as e:
                logger.error(f"Unexpected error processing {text_file}: {e}")
        
        logger.info("Audio generation completed")
        
    except Exception as e:
        logger.error(f"Error during audio generation: {e}")
        raise


def main():
    """Main entry point for the script."""
    parser = argparse.ArgumentParser(
        description="Extract slides and speaker notes from PowerPoint presentations"
    )
    parser.add_argument(
        "presentation",
        type=str,
        help="Path to the PowerPoint presentation file (.pptx)"
    )
    
    args = parser.parse_args()
    
    # Validate input file
    pptx_path = Path(args.presentation)
    if not pptx_path.exists():
        logger.error(f"File not found: {pptx_path}")
        sys.exit(1)
    
    if not pptx_path.suffix.lower() in ['.pptx', '.ppt']:
        logger.error(f"Invalid file type: {pptx_path.suffix}. Expected .pptx or .ppt")
        sys.exit(1)
    
    logger.info(f"Processing presentation: {pptx_path}")
    
    try:
        # Set up output folder
        output_folder = setup_output_folder(pptx_path)
        
        # Export slides as PNG
        slide_count = export_slides_as_png(pptx_path, output_folder)
        
        # Extract speaker notes
        extract_speaker_notes(pptx_path, output_folder, slide_count)
        
        # Generate audio from speaker notes
        generate_audio_from_notes(output_folder, slide_count)
        
        logger.info(f"Successfully processed {slide_count} slides")
        logger.info(f"Output saved to: {output_folder}")
        
    except Exception as e:
        logger.error(f"Failed to process presentation: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
