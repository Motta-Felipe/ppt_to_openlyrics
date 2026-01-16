Agent Instructions for PPT → OpenLyrics 0.8
================================================

Scope
- Input: PPT/PPTX files containing song slides. Extract slide text in reading order; treat each slide as one block of lines.

Cleaning & Normalization
- Decode as UTF-8. Strip ASCII control chars (0x00–0x1F, 0x7F). Normalize tabs/form-feeds/vertical tabs to newlines. Trim trailing spaces per line; keep meaningful leading spaces. Collapse multiple blank lines to a single blank separator.

Block Detection
- Within each slide, if a line starts with "Ritornello:" (case-insensitive, leading spaces ignored), treat the remainder of that slide as a chorus block; drop the label line itself.
- Number choruses in encounter order: c1, c2, ...
- All other blocks are verses in encounter order: v1, v2, ...
- If a slide has both verse text and a chorus marker, emit two consecutive blocks: verse first, then chorus.

Line Handling Inside a Block
- Preserve line order; join lines with <br />. Remove lines that are empty after cleaning; do not emit empty blocks.

XML Output (one file per song)
- File name: <ppt_basename>.xml (UTF-8).
- Envelope:
  <song version="0.8" xmlns="http://openlyrics.info/namespace/2009/song">
    <properties>
      <titles><title>{Title from filename}</title></titles>
      <!-- add author if known -->
    </properties>
    <lyrics>
      <verse name="v1"><lines>line1<br />line2</lines></verse>
      <verse name="c1"><lines>line1<br />line2</lines></verse>
      ...
    </lyrics>
  </song>

Ordering & Safety
- Emit verses/choruses in the order encountered while scanning slides. Ensure verse names are unique and sequential. Skip empty blocks. Escape XML special chars (& < > " ').

Validation Checklist
- No control chars remain.
- Namespace and version match OpenLyrics 0.8.
- Chorus and verse numbering consecutive from first occurrence.
- Well-formed XML; UTF-8 encoded.
