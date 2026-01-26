#!/usr/bin/env python3
"""
PPT/PPTX to OpenLyrics 0.8 Converter

Converts PowerPoint presentation files containing song lyrics to OpenLyrics XML format.
Follows the rules in agent_instructions.md.
"""

import os
import re
import sys
import zipfile
import shutil
import subprocess
import tempfile
from pathlib import Path
from datetime import datetime
from xml.sax.saxutils import escape as xml_escape

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

# Constants
OPENLYRICS_NS = "http://openlyrics.info/namespace/2009/song"
LIBREOFFICE_PATH = r"C:\Program Files\LibreOffice\program\soffice.exe"
OPENLYRICS_VERSION = "0.8"
AUDIO_EXTENSIONS = {'.mp3', '.wav', '.m4a', '.wma', '.ogg', '.flac'}
SUPPORTED_EXTENSIONS = {'.pptx', '.ppt', '.odp'}


def convert_ppt_to_pptx(ppt_path: Path, temp_dir: Path) -> Path | None:
    """
    Convert legacy .ppt or .odp to .pptx using LibreOffice.
    Returns path to converted file or None on failure.
    """
    if not Path(LIBREOFFICE_PATH).exists():
        return None
    
    try:
        # Run LibreOffice in headless mode to convert to pptx
        result = subprocess.run(
            [
                LIBREOFFICE_PATH,
                '--headless',
                '--convert-to', 'pptx',
                '--outdir', str(temp_dir),
                str(ppt_path)
            ],
            capture_output=True,
            timeout=60
        )
        
        if result.returncode == 0:
            # Find the converted file
            converted = temp_dir / (ppt_path.stem + '.pptx')
            if converted.exists():
                return converted
    except Exception:
        pass
    
    return None

# Patterns for detection
VERSE_MARKER_PATTERN = re.compile(r'^(\d+)[°ª]?\s*strofa\b', re.IGNORECASE)
CHORUS_LABEL_PATTERN = re.compile(r'^(ritornello|rit\.?)\s*:?\s*', re.IGNORECASE)
TITLE_NUMBER_PATTERN = re.compile(r'^(\d+)\s*[-–—.]?\s*')
CHURCH_FILTER_PATTERN = re.compile(r'chiesa\s*(di\s+)?olgiate', re.IGNORECASE)


class ConversionLog:
    """Maintains a log file for conversion status."""
    
    def __init__(self, log_path: Path):
        self.log_path = log_path
        self._init_log()
    
    def _init_log(self):
        if not self.log_path.exists():
            with open(self.log_path, 'w', encoding='utf-8') as f:
                f.write(f"Conversion Log - Started {datetime.now().isoformat()}\n")
                f.write("=" * 60 + "\n\n")
    
    def log(self, filename: str, status: str, has_media: bool = False, 
            media_extracted: str = None, notes: str = None):
        with open(self.log_path, 'a', encoding='utf-8') as f:
            f.write(f"File: {filename}\n")
            f.write(f"  Status: {status}\n")
            f.write(f"  Has Media: {'Yes' if has_media else 'No'}\n")
            if media_extracted:
                f.write(f"  Media Extracted: {media_extracted}\n")
            if notes:
                f.write(f"  Notes: {notes}\n")
            f.write(f"  Processed: {datetime.now().isoformat()}\n")
            f.write("-" * 40 + "\n")


def clean_text(text: str) -> str:
    """
    Clean text by stripping control chars and normalizing whitespace.
    """
    if not text:
        return ""
    
    # Strip ASCII control chars (0x00-0x1F, 0x7F) except newline/tab
    result = []
    for char in text:
        code = ord(char)
        if code == 0x09:  # Tab -> newline
            result.append('\n')
        elif code == 0x0A:  # Newline - keep
            result.append(char)
        elif code == 0x0B or code == 0x0C:  # Vertical tab, form feed -> newline
            result.append('\n')
        elif code == 0x0D:  # Carriage return -> newline
            result.append('\n')
        elif 0x00 <= code <= 0x1F or code == 0x7F:  # Other control chars - remove
            continue
        else:
            result.append(char)
    
    text = ''.join(result)
    
    # Normalize multiple newlines
    text = re.sub(r'\n{3,}', '\n\n', text)
    
    # Trim trailing spaces per line
    lines = [line.rstrip() for line in text.split('\n')]
    
    return '\n'.join(lines)


def extract_title_from_filename(filename: str) -> str:
    """
    Extract song title from filename, stripping leading numbers and separators.
    Returns the title in Title Case.
    """
    # Remove extension
    name = Path(filename).stem
    
    # Remove leading numbers and separators
    name = TITLE_NUMBER_PATTERN.sub('', name)
    
    # Remove trailing underscores
    name = name.rstrip('_')
    
    # Clean up extra spaces
    name = ' '.join(name.split())
    
    # Convert to Title Case
    name = name.strip() or Path(filename).stem
    
    # Custom title case that handles apostrophes and accented chars better
    def title_case_word(word: str) -> str:
        if not word:
            return word
        # Handle words with apostrophes (e.g., "DELL'AMICIZIA" -> "Dell'Amicizia")
        if "'" in word:
            parts = word.split("'")
            return "'".join(part.capitalize() for part in parts)
        return word.capitalize()
    
    words = name.split()
    titled = ' '.join(title_case_word(word) for word in words)
    
    return titled


def extract_text_from_pptx(pptx_path: Path) -> list[list[str]]:
    """
    Extract text from PPTX file, returning list of slides, each slide is a list of text blocks.
    """
    slides_text = []
    
    try:
        prs = Presentation(str(pptx_path))
        
        for slide in prs.slides:
            slide_texts = []
            
            for shape in slide.shapes:
                if shape.has_text_frame:
                    text_parts = []
                    for paragraph in shape.text_frame.paragraphs:
                        para_text = ''.join(run.text for run in paragraph.runs)
                        if para_text.strip():
                            text_parts.append(para_text)
                    
                    if text_parts:
                        slide_texts.append('\n'.join(text_parts))
            
            if slide_texts:
                slides_text.append(slide_texts)
    
    except Exception as e:
        raise RuntimeError(f"Failed to extract text from {pptx_path}: {e}")
    
    return slides_text


def extract_media_from_pptx(pptx_path: Path, output_dir: Path, base_name: str) -> str | None:
    """
    Extract embedded audio from PPTX if present.
    Returns the extracted filename or None.
    """
    try:
        with zipfile.ZipFile(pptx_path, 'r') as zf:
            media_files = [n for n in zf.namelist() if n.startswith('ppt/media/')]
            
            for media_file in media_files:
                ext = Path(media_file).suffix.lower()
                if ext in AUDIO_EXTENSIONS:
                    # Check if audio already exists
                    existing_audio = find_existing_audio(output_dir, base_name)
                    if existing_audio:
                        return None  # Skip, already exists
                    
                    # Extract with song title
                    output_name = f"{base_name}{ext}"
                    output_path = output_dir / output_name
                    
                    with zf.open(media_file) as src:
                        with open(output_path, 'wb') as dst:
                            shutil.copyfileobj(src, dst)
                    
                    return output_name
    
    except Exception:
        pass
    
    return None


def find_existing_audio(folder: Path, base_name: str) -> Path | None:
    """
    Check if an audio file matching the base name already exists.
    """
    # Normalize base name for comparison
    base_lower = base_name.lower()
    
    for file in folder.iterdir():
        if file.suffix.lower() in AUDIO_EXTENSIONS:
            file_base = extract_title_from_filename(file.name).lower()
            if file_base == base_lower or base_lower in file.stem.lower():
                return file
    
    return None


def is_header_footer(text: str, title: str, all_slides_text: list[list[str]]) -> bool:
    """
    Check if text is a header/footer that should be ignored.
    Only filters out church names, repeated titles, and very short standalone headers.
    Does NOT filter out repeated lyric phrases (like refrains within verses).
    """
    text_stripped = text.strip()
    text_lower = text_stripped.lower()
    
    # Church name filter
    if CHURCH_FILTER_PATTERN.search(text_stripped):
        return True
    
    # Check if it's the title repeated
    title_lower = title.lower()
    if text_lower == title_lower:
        return True
    
    # Check for numbered title pattern (e.g., "47 - Là nel ciel...")
    title_match = TITLE_NUMBER_PATTERN.sub('', text_stripped).strip().lower()
    if title_match == title_lower:
        return True
    
    # Only filter very short text (<=3 words) that appears on most slides as standalone
    # This catches headers like "Page 1" or church names, but not lyric refrains
    words = text_stripped.split()
    if len(words) > 3:
        return False  # Longer text is likely lyrics, not header/footer
    
    # Check if this short text appears on most slides as a standalone element
    occurrence_count = 0
    total_slides = len(all_slides_text)
    
    for slide_texts in all_slides_text:
        for slide_text in slide_texts:
            # Only count if text is standalone (the entire text block) or at very start
            if text_stripped == slide_text.strip():
                occurrence_count += 1
                break
    
    if total_slides > 2 and occurrence_count >= total_slides * 0.7:
        return True
    
    return False


def parse_slides_to_blocks(slides_text: list[list[str]], title: str) -> list[dict]:
    """
    Parse slide text into verse/chorus blocks.
    Returns list of {'type': 'verse'|'chorus', 'lines': [...], 'number': int}
    """
    blocks = []
    verse_num = 0
    chorus_num = 0
    seen_choruses = {}  # For detecting repeated blocks
    
    for slide_texts in slides_text:
        # Combine all text blocks from this slide
        combined_text = '\n'.join(slide_texts)
        combined_text = clean_text(combined_text)
        
        if not combined_text.strip():
            continue
        
        lines = combined_text.split('\n')
        filtered_lines = []
        current_is_chorus = False
        explicit_verse_num = None
        
        for line in lines:
            line_stripped = line.strip()
            
            if not line_stripped:
                continue
            
            # Check for header/footer
            if is_header_footer(line_stripped, title, slides_text):
                continue
            
            # Check for verse marker (1° Strofa, 2° Strofa, etc.)
            verse_match = VERSE_MARKER_PATTERN.match(line_stripped)
            if verse_match:
                explicit_verse_num = int(verse_match.group(1))
                continue  # Skip the marker line
            
            # Check for chorus label
            chorus_match = CHORUS_LABEL_PATTERN.match(line_stripped)
            if chorus_match:
                # If we have accumulated lines, save them as verse first
                if filtered_lines and not current_is_chorus:
                    verse_num += 1
                    actual_num = explicit_verse_num if explicit_verse_num else verse_num
                    blocks.append({
                        'type': 'verse',
                        'lines': filtered_lines.copy(),
                        'number': actual_num
                    })
                    filtered_lines = []
                    explicit_verse_num = None
                
                current_is_chorus = True
                # Get remaining text after chorus label
                remaining = CHORUS_LABEL_PATTERN.sub('', line_stripped).strip()
                if remaining:
                    filtered_lines.append(remaining)
                continue
            
            filtered_lines.append(line_stripped)
        
        if not filtered_lines:
            continue
        
        # Check for repeated block (auto-detect chorus)
        block_key = '\n'.join(filtered_lines).lower()
        
        if current_is_chorus:
            if block_key not in seen_choruses:
                chorus_num += 1
                seen_choruses[block_key] = chorus_num
            
            blocks.append({
                'type': 'chorus',
                'lines': filtered_lines,
                'number': seen_choruses[block_key]
            })
        elif block_key in seen_choruses:
            # This is a repeated block - it's a chorus
            blocks.append({
                'type': 'chorus',
                'lines': filtered_lines,
                'number': seen_choruses[block_key]
            })
        else:
            # Check if this block repeats later (auto-detect chorus)
            is_repeated = False
            for future_slide_texts in slides_text[slides_text.index(slide_texts) + 1:]:
                future_combined = clean_text('\n'.join(future_slide_texts))
                future_lines = [l.strip() for l in future_combined.split('\n') if l.strip()]
                # Filter out headers
                future_lines = [l for l in future_lines if not is_header_footer(l, title, slides_text)]
                future_key = '\n'.join(future_lines).lower()
                
                if future_key == block_key:
                    is_repeated = True
                    break
            
            if is_repeated and len(filtered_lines) <= 8:  # Choruses are usually short
                if block_key not in seen_choruses:
                    chorus_num += 1
                    seen_choruses[block_key] = chorus_num
                
                blocks.append({
                    'type': 'chorus',
                    'lines': filtered_lines,
                    'number': seen_choruses[block_key]
                })
            else:
                verse_num += 1
                actual_num = explicit_verse_num if explicit_verse_num else verse_num
                blocks.append({
                    'type': 'verse',
                    'lines': filtered_lines,
                    'number': actual_num
                })
    
    return blocks


def deduplicate_blocks(blocks: list[dict]) -> list[dict]:
    """
    Remove consecutive duplicate blocks (same content).
    """
    if not blocks:
        return blocks
    
    result = [blocks[0]]
    
    for block in blocks[1:]:
        prev_key = '\n'.join(result[-1]['lines']).lower()
        curr_key = '\n'.join(block['lines']).lower()
        
        if curr_key != prev_key:
            result.append(block)
    
    return result


def generate_openlyrics_xml(title: str, blocks: list[dict]) -> str:
    """
    Generate OpenLyrics 0.8 XML from parsed blocks.
    """
    # Deduplicate consecutive identical blocks
    blocks = deduplicate_blocks(blocks)
    
    # Renumber verses and choruses sequentially
    verse_counter = 0
    chorus_counter = 0
    seen_verses = {}
    seen_choruses = {}
    
    verses_xml = []
    
    for block in blocks:
        lines_escaped = [xml_escape(line) for line in block['lines']]
        lines_content = '<br />'.join(lines_escaped)
        
        if block['type'] == 'chorus':
            key = '\n'.join(block['lines']).lower()
            if key not in seen_choruses:
                chorus_counter += 1
                seen_choruses[key] = chorus_counter
            name = f"c{seen_choruses[key]}"
        else:
            key = '\n'.join(block['lines']).lower()
            if key not in seen_verses:
                verse_counter += 1
                seen_verses[key] = verse_counter
            name = f"v{seen_verses[key]}"
        
        verses_xml.append(f'    <verse name="{name}">\n      <lines>{lines_content}</lines>\n    </verse>')
    
    title_escaped = xml_escape(title)
    
    xml = f'''<?xml version="1.0" encoding="UTF-8"?>
<song version="{OPENLYRICS_VERSION}" xmlns="{OPENLYRICS_NS}">
  <properties>
    <titles>
      <title>{title_escaped}</title>
    </titles>
  </properties>
  <lyrics>
{chr(10).join(verses_xml)}
  </lyrics>
</song>'''
    
    return xml


def convert_file(input_path: Path, output_dir: Path, log: ConversionLog) -> bool:
    """
    Convert a single PPT/PPTX file to OpenLyrics XML.
    Returns True if successful.
    """
    filename = input_path.name
    title = extract_title_from_filename(filename)
    temp_pptx = None
    
    try:
        # Extract text - handle different formats
        if input_path.suffix.lower() == '.pptx':
            slides_text = extract_text_from_pptx(input_path)
            actual_pptx = input_path
        elif input_path.suffix.lower() in {'.ppt', '.odp'}:
            # Legacy PPT or ODP - convert using LibreOffice first
            with tempfile.TemporaryDirectory() as temp_dir:
                temp_dir_path = Path(temp_dir)
                converted_path = convert_ppt_to_pptx(input_path, temp_dir_path)
                
                if converted_path and converted_path.exists():
                    slides_text = extract_text_from_pptx(converted_path)
                    actual_pptx = converted_path
                else:
                    # Fallback: try direct extraction (rarely works for .ppt)
                    try:
                        slides_text = extract_text_from_pptx(input_path)
                        actual_pptx = input_path
                    except Exception:
                        log.log(filename, "SKIPPED", notes="LibreOffice conversion failed and direct extraction not possible")
                        return False
                
                if not slides_text:
                    log.log(filename, "SKIPPED", notes="No text content found after conversion")
                    return False
                
                # Parse into blocks
                blocks = parse_slides_to_blocks(slides_text, title)
                
                if not blocks:
                    log.log(filename, "SKIPPED", notes="No verse/chorus blocks detected")
                    return False
                
                # Generate XML
                xml_content = generate_openlyrics_xml(title, blocks)
                
                # Write XML file
                output_path = output_dir / f"{title}.xml"
                with open(output_path, 'w', encoding='utf-8') as f:
                    f.write(xml_content)
                
                # Extract media from converted pptx if exists
                media_extracted = None
                has_media = False
                if converted_path and converted_path.exists():
                    media_extracted = extract_media_from_pptx(converted_path, output_dir, title)
                    has_media = media_extracted is not None
                
                log.log(filename, "SUCCESS", has_media=has_media, media_extracted=media_extracted, 
                       notes="Converted from legacy format via LibreOffice")
                return True
        else:
            log.log(filename, "SKIPPED", notes=f"Unsupported format: {input_path.suffix}")
            return False
        
        if not slides_text:
            log.log(filename, "SKIPPED", notes="No text content found")
            return False
        
        # Parse into blocks
        blocks = parse_slides_to_blocks(slides_text, title)
        
        if not blocks:
            log.log(filename, "SKIPPED", notes="No verse/chorus blocks detected")
            return False
        
        # Generate XML
        xml_content = generate_openlyrics_xml(title, blocks)
        
        # Write XML file
        output_path = output_dir / f"{title}.xml"
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(xml_content)
        
        # Extract media if present
        media_extracted = None
        has_media = False
        
        if input_path.suffix.lower() == '.pptx':
            media_extracted = extract_media_from_pptx(input_path, output_dir, title)
            has_media = media_extracted is not None
        
        log.log(filename, "SUCCESS", has_media=has_media, media_extracted=media_extracted)
        return True
    
    except Exception as e:
        log.log(filename, "ERROR", notes=str(e))
        return False


def main():
    """Main entry point."""
    # Determine input folder
    if len(sys.argv) > 1:
        input_folder = Path(sys.argv[1])
    else:
        input_folder = Path(r"c:\ppt_songs\1. Canti SDS - Chiesa Viva")
    
    # Output folder - create openLyrics subfolder
    output_folder = input_folder / "openLyrics"
    output_folder.mkdir(exist_ok=True)
    
    if not input_folder.exists():
        print(f"Error: Input folder not found: {input_folder}")
        sys.exit(1)
    
    # Initialize log
    log = ConversionLog(output_folder / "conversion_log.txt")
    
    # Find all presentation files
    ppt_files = []
    for ext in SUPPORTED_EXTENSIONS:
        ppt_files.extend(input_folder.glob(f"*{ext}"))
    
    ppt_files.sort(key=lambda p: p.name.lower())
    
    print(f"Found {len(ppt_files)} presentation files to convert")
    print(f"Output folder: {output_folder}")
    print("-" * 50)
    
    success_count = 0
    skip_count = 0
    error_count = 0
    
    for i, ppt_file in enumerate(ppt_files, 1):
        print(f"[{i}/{len(ppt_files)}] Processing: {ppt_file.name}")
        
        if convert_file(ppt_file, output_folder, log):
            success_count += 1
            print(f"  ✓ Converted successfully")
        else:
            # Check log for reason
            skip_count += 1
            print(f"  ⚠ Skipped or failed")
    
    print("-" * 50)
    print(f"Conversion complete!")
    print(f"  Success: {success_count}")
    print(f"  Skipped/Errors: {skip_count}")
    print(f"  Log file: {output_folder / 'conversion_log.txt'}")


if __name__ == "__main__":
    main()
