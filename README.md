
## Overview

The convert_to_openlyrics.py script converts PowerPoint presentations (`.pptx`, `.ppt`, `.odp`) containing song lyrics into **OpenLyrics 0.8 XML** format for use in projection applications like FreeShow.

**Key Features**:
- Recursive text extraction from group shapes
- Automatic case normalization (all-caps to sentence case)
- Organized output with separate folders for raw text and XML
- Comprehensive header/footer filtering
- Chorus auto-detection from repeated blocks
- Media extraction support

---

## Architecture

```
┌─────────────────┐     ┌──────────────────┐     ┌─────────────────┐
│  Input Files    │────▶│  Text Extraction │────▶│  Block Parsing  │
│ .pptx/.ppt/.odp │     │  & Cleaning      │     │  (verse/chorus) │
│  (incl. groups) │     │  (case norm.)    │     │  (auto-detect)  │
└─────────────────┘     └──────────────────┘     └─────────────────┘
                                                          │
                                                          ▼
┌─────────────────┐     ┌──────────────────┐     ┌─────────────────┐
│  Output Files   │◀────│  XML Generation  │◀────│  Deduplication  │
│  .xml + audio   │     │  OpenLyrics 0.8  │     │  & Renumbering  │
│  + raw .txt     │     │                  │     │                 │
└─────────────────┘     └──────────────────┘     └─────────────────┘
```

---

## Module Structure

### 1. Constants & Configuration

````python
OPENLYRICS_NS = "http://openlyrics.info/namespace/2009/song"
LIBREOFFICE_PATH = r"C:\Program Files\LibreOffice\program\soffice.exe"
OPENLYRICS_VERSION = "0.8"
AUDIO_EXTENSIONS = {'.mp3', '.wav', '.m4a', '.wma', '.ogg', '.flac'}
SUPPORTED_EXTENSIONS = {'.pptx', '.ppt', '.odp'}
````

- **OPENLYRICS_NS**: XML namespace for OpenLyrics format
- **LIBREOFFICE_PATH**: Path to LibreOffice for legacy file conversion
- **AUDIO_EXTENSIONS**: Recognized audio formats for extraction
- **SUPPORTED_EXTENSIONS**: Input file formats the script handles

---

### 2. Regex Patterns

````python
VERSE_MARKER_PATTERN = re.compile(r'^(\d+)[°ª]?\s*strofa\b', re.IGNORECASE)
CHORUS_LABEL_PATTERN = re.compile(r'^(ritornello|rit\.?)\s*:?\s*', re.IGNORECASE)
TITLE_NUMBER_PATTERN = re.compile(r'^(\d+)\s*[-–—.]?\s*')
CHURCH_FILTER_PATTERN = re.compile(r'chiesa\s*(di\s+)?olgiate', re.IGNORECASE)
````

| Pattern | Purpose | Example Match |
|---------|---------|---------------|
| `VERSE_MARKER_PATTERN` | Detects verse markers | `1° Strofa`, `2ª strofa` |
| `CHORUS_LABEL_PATTERN` | Detects chorus labels | `Ritornello:`, `Rit.` |
| `TITLE_NUMBER_PATTERN` | Strips leading numbers from filenames | `47 - `, `10 ` |
| `CHURCH_FILTER_PATTERN` | Filters church name headers | `Chiesa di Olgiate` |

---

### 3. ConversionLog Class

````python
class ConversionLog:
    """Maintains a log file for conversion status."""
    
    def __init__(self, log_path: Path):
        self.log_path = log_path
        self._init_log()
    
    def log(self, filename: str, status: str, has_media: bool = False, 
            media_extracted: str = None, notes: str = None):
        # Appends entry to log file
````

**Purpose**: Tracks conversion results for each file:
- Filename
- Status (SUCCESS / SKIPPED / ERROR)
- Media presence and extraction status
- Notes (errors or special handling)

**Output format**:
```
File: 47 - Camminando.pptx
  Status: SUCCESS
  Has Media: No
  Processed: 2026-01-26T10:30:00
----------------------------------------
```

---

### 4. Text Cleaning Functions

#### `clean_text(text: str) -> str`

````python
def clean_text(text: str) -> str:
    """
    Clean text by stripping control chars and normalizing whitespace.
    """
    # 1. Convert tabs, form feeds, vertical tabs to newlines
    # 2. Strip ASCII control chars (0x00-0x1F, 0x7F)
    # 3. Collapse multiple newlines to max 2
    # 4. Trim trailing spaces per line
    # 5. Convert all-uppercase lines to sentence case
````

**Transformations**:
| Input | Output |
|-------|--------|
| `\t` (tab) | `\n` |
| `\v` (vertical tab) | `\n` |
| `\f` (form feed) | `\n` |
| `\r` (carriage return) | `\n` |
| `\n\n\n\n` | `\n\n` |
| `TU SEI SANTO` | `Tu sei santo` |

---

#### `extract_title_from_filename(filename: str) -> str`

````python
def extract_title_from_filename(filename: str) -> str:
    """
    Extract song title from filename, stripping leading numbers and separators.
    Returns the title in Title Case.
    """
    # 1. Remove extension
    # 2. Remove leading numbers (47 - , 10 , etc.)
    # 3. Remove trailing underscores
    # 4. Convert to Title Case (handles apostrophes)
````

**Examples**:
| Input Filename | Output Title |
|----------------|--------------|
| `47 - Camminando.pptx` | `Camminando` |
| `10 CANTO PER CRISTO.pptx` | `Canto Per Cristo` |
| `DELL'AMICIZIA.pptx` | `Dell'Amicizia` |
| `UNICA VIA_.ppt` | `Unica Via` |

---

### 5. Text Extraction Functions

#### `extract_text_from_pptx(pptx_path: Path) -> list[list[str]]`

````python
def extract_text_from_pptx(pptx_path: Path) -> list[list[str]]:
    """
    Extract text from PPTX file.
    Returns: list of slides, each slide is a list of text blocks.
    """
    prs = Presentation(str(pptx_path))
    
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    # Extract text from each run
            elif hasattr(shape, 'shapes'):  # Group shape support
                # Recursively extract from grouped shapes
````

**Enhanced Features**:
- **Group Shape Support**: Recursively extracts text from shapes nested inside group elements (common in complex PPTX layouts)
- **Comprehensive Coverage**: Handles all text-containing shapes, including those in groups

**Return structure**:
```python
[
    # Slide 1
    ["Text block 1\nLine 2", "Text block 2"],
    # Slide 2
    ["Verse text here"],
    # ...
]
```

---

#### `convert_ppt_to_pptx(ppt_path: Path, temp_dir: Path) -> Path | None`

````python
def convert_ppt_to_pptx(ppt_path: Path, temp_dir: Path) -> Path | None:
    """
    Convert legacy .ppt or .odp to .pptx using LibreOffice.
    """
    subprocess.run([
        LIBREOFFICE_PATH,
        '--headless',
        '--convert-to', 'pptx',
        '--outdir', str(temp_dir),
        str(ppt_path)
    ])
````

**Flow**:
```
.ppt/.odp ──▶ LibreOffice (headless) ──▶ .pptx ──▶ python-pptx extraction
```

---

### 6. Header/Footer Detection

#### `is_header_footer(text: str, title: str, all_slides_text: list[list[str]]) -> bool`

````python
def is_header_footer(text: str, title: str, all_slides_text: list[list[str]]) -> bool:
    """
    Check if text is a header/footer that should be ignored.
    """
    # 1. Filter "Chiesa di Olgiate" or similar
    # 2. Filter repeated title (e.g., "47 - Camminando")
    # 3. Filter short text (≤3 words) appearing on 70%+ of slides as standalone
    # 4. Do NOT filter longer text (likely lyrics, not headers)
````

**Key logic**:
- Text with **more than 3 words** is assumed to be lyrics (not filtered)
- Short text must appear on **70%+ of slides** as a **standalone block** to be filtered
- This prevents filtering refrain lines like "camminando, camminando" that appear within verses

---

### 7. Block Parsing

#### `parse_slides_to_blocks(slides_text: list[list[str]], title: str) -> list[dict]`

````python
def parse_slides_to_blocks(slides_text: list[list[str]], title: str) -> list[dict]:
    """
    Parse slide text into verse/chorus blocks.
    Returns: list of {'type': 'verse'|'chorus', 'lines': [...], 'number': int}
    """
````

**Detection logic**:

```
┌─────────────────────────────────────────────────────────┐
│                    For each slide                        │
├─────────────────────────────────────────────────────────┤
│ 1. Is it a header/footer? → Skip                        │
│ 2. Does it match "1° Strofa"? → Mark as verse N         │
│ 3. Does it match "Ritornello:"? → Mark as chorus        │
│ 4. Does this block repeat later? → Auto-detect chorus   │
│ 5. Otherwise → Mark as verse                            │
└─────────────────────────────────────────────────────────┘
```

**Chorus auto-detection**:
- If a block appears **identically** on a later slide
- AND has **≤8 lines** (choruses are usually short)
- → Treat it as a chorus

**Output example**:
```python
[
    {'type': 'verse', 'lines': ['Line 1', 'Line 2'], 'number': 1},
    {'type': 'chorus', 'lines': ['Chorus line'], 'number': 1},
    {'type': 'verse', 'lines': ['Verse 2 text'], 'number': 2},
    {'type': 'chorus', 'lines': ['Chorus line'], 'number': 1},  # Same chorus repeated
]
```

---

### 8. Deduplication

#### `deduplicate_blocks(blocks: list[dict]) -> list[dict]`

````python
def deduplicate_blocks(blocks: list[dict]) -> list[dict]:
    """
    Remove consecutive duplicate blocks (same content).
    """
````

**Example**:
```
Input:  [v1, c1, c1, v2, c1]  # c1 appears twice consecutively
Output: [v1, c1, v2, c1]      # Consecutive duplicate removed
```

---

### 9. XML Generation

#### `generate_openlyrics_xml(title: str, blocks: list[dict]) -> str`

````python
def generate_openlyrics_xml(title: str, blocks: list[dict]) -> str:
    """
    Generate OpenLyrics 0.8 XML from parsed blocks.
    """
    # 1. Deduplicate consecutive blocks
    # 2. Renumber verses (v1, v2, ...) and choruses (c1, c2, ...)
    # 3. Escape XML special characters (& < > " ')
    # 4. Join lines with <br />
    # 5. Build final XML structure
````

**Output structure**:
```xml
<?xml version="1.0" encoding="UTF-8"?>
<song version="0.8" xmlns="http://openlyrics.info/namespace/2009/song">
  <properties>
    <titles>
      <title>Song Title</title>
    </titles>
  </properties>
  <lyrics>
    <verse name="v1">
      <lines>First line<br />Second line</lines>
    </verse>
    <verse name="c1">
      <lines>Chorus line 1<br />Chorus line 2</lines>
    </verse>
  </lyrics>
</song>
```

---

### 10. Media Extraction

#### `extract_media_from_pptx(pptx_path: Path, output_dir: Path, base_name: str) -> str | None`

````python
def extract_media_from_pptx(pptx_path: Path, output_dir: Path, base_name: str) -> str | None:
    """
    Extract embedded audio from PPTX if present.
    """
    with zipfile.ZipFile(pptx_path, 'r') as zf:
        media_files = [n for n in zf.namelist() if n.startswith('ppt/media/')]
        # Check for audio extensions
        # Skip if audio already exists
        # Extract with song title as filename
````

**PPTX internal structure**:
```
presentation.pptx (ZIP archive)
├── ppt/
│   ├── slides/
│   ├── media/
│   │   ├── audio1.mp3  ◀── Extracted
│   │   └── image1.png
```

---

### 11. Main Conversion Function

#### `convert_file(input_path: Path, output_dir: Path, log: ConversionLog) -> bool`

````python
def convert_file(input_path: Path, output_dir: Path, log: ConversionLog) -> bool:
    """
    Convert a single PPT/PPTX file to OpenLyrics XML.
    """
````

**Flow**:
```
┌─────────────┐
│ Input File  │
└──────┬──────┘
       │
       ▼
┌──────────────────────────────────┐
│ Is it .pptx?                     │
│   YES → Direct extraction        │
│   NO → Convert via LibreOffice   │
└──────┬───────────────────────────┘
       │
       ▼
┌──────────────────────────────────┐
│ Extract text from slides         │
└──────┬───────────────────────────┘
       │
       ▼
┌──────────────────────────────────┐
│ Parse into verse/chorus blocks   │
└──────┬───────────────────────────┘
       │
       ▼
┌──────────────────────────────────┐
│ Generate OpenLyrics XML          │
└──────┬───────────────────────────┘
       │
       ▼
┌──────────────────────────────────┐
│ Write .xml file                  │
│ Extract audio (if present)       │
│ Log result                       │
└──────────────────────────────────┘
```

---

### 12. Main Entry Point

#### `main()`

````python
def main():
    # 1. Parse input folder (CLI arg or default)
    # 2. Create openLyrics subfolder
    # 3. Initialize conversion log
    # 4. Find all .pptx/.ppt/.odp files
    # 5. Process each file
    # 6. Print summary
````

**Usage**:
```bash
# Default folder
python convert_to_openlyrics.py

# Custom folder
python convert_to_openlyrics.py "C:\path\to\songs"
```

### Environment Variables

The script supports the following environment variables for configuration:

- **`PPT_REMOVE_NUMBER_PREFIX`**: Controls whether to remove leading numbers from filenames
  - **Default**: `true` (removes numbers like "47 - " from "47 - Camminando.pptx")
  - **Values**: Set to `false`, `0`, `no`, or `off` to keep numbers in titles
  - **Example**: `PPT_REMOVE_NUMBER_PREFIX=false python convert_to_openlyrics.py`

---

## Recent Features

### Group Shape Support
- **Problem**: Some PPTX files use grouped shapes where text elements are nested inside group containers
- **Solution**: Recursive text extraction that handles group shapes (`hasattr(shape, 'shapes')`)
- **Benefit**: Captures text from complex slide layouts that were previously missed

### Case Normalization
- **Problem**: Lyrics in all caps (e.g., `TU SEI SANTO`) are harder to read
- **Solution**: Automatic conversion of all-uppercase lines to sentence case (`Tu sei santo`)
- **Benefit**: Improved readability of generated OpenLyrics XML

### Enhanced Output Organization
- **Raw Text Files**: Debug-friendly text files showing extracted content per slide
- **Structured Folders**: Separate `raw text/` and `openlyrics/` subfolders for better organization
- **Media Extraction**: Audio files now extracted to the `openlyrics/` folder alongside XML

---

**Input**: `47 - Camminando.pptx`

| Slide | Raw Text |
|-------|----------|
| 1 | `Chiesa di Olgiate`<br>`47 - Camminando`<br>`Felice io vado alla casa del ciel`<br>`camminando, camminando` |
| 2 | `Camminando, camminando`<br>`verso la città dove sta Gesù` |
| 3 | `Insieme noi andiamo`<br>`alla patria del ciel` |
| 4 | `Camminando, camminando`<br>`verso la città dove sta Gesù` |

**Processing**:
1. **Filter**: Remove "Chiesa di Olgiate" and "47 - Camminando"
2. **Parse**: Slides 2 and 4 are identical → detected as chorus
3. **Output**:

```xml
<song version="0.8" xmlns="http://openlyrics.info/namespace/2009/song">
  <properties>
    <titles><title>Camminando</title></titles>
  </properties>
  <lyrics>
    <verse name="v1">
      <lines>Felice io vado alla casa del ciel<br />camminando, camminando</lines>
    </verse>
    <verse name="c1">
      <lines>Camminando, camminando<br />verso la città dove sta Gesù</lines>
    </verse>
    <verse name="v2">
      <lines>Insieme noi andiamo<br />alla patria del ciel</lines>
    </verse>
  </lyrics>
</song>
```

---

## Error Handling

| Scenario | Handling |
|----------|----------|
| LibreOffice not found | Skip legacy files, log warning |
| No text in presentation | Log as SKIPPED |
| No verse/chorus detected | Log as SKIPPED |
| Extraction exception | Log as ERROR with message |
| Audio already exists | Skip extraction |

---

## Output Files

```
input_folder/
├── 47 - Camminando.pptx
├── 10 CANTO PER CRISTO.pptx
└── output/                          ◀── Output base folder
    ├── raw text/                    ◀── Raw extracted text files
    │   ├── Camminando.txt
    │   └── Canto Per Cristo.txt
    ├── openlyrics/                  ◀── OpenLyrics XML and media
    │   ├── Camminando.xml
    │   ├── Canto Per Cristo.xml
    │   └── Song Title.mp3           ◀── Extracted audio (if any)
    └── conversion_log.txt           ◀── Processing log
```