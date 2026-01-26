Agent Instructions for PPT → OpenLyrics 0.8
================================================

Scope
- Input: PPT/PPTX files containing song slides. Extract slide text in reading order; treat each slide as one block of lines.

Cleaning & Normalization
- Decode as UTF-8. Strip ASCII control chars (0x00–0x1F, 0x7F). Normalize tabs/form-feeds/vertical tabs to newlines. Trim trailing spaces per line; keep meaningful leading spaces. Collapse multiple blank lines to a single blank separator.

Header/Footer Filtering
- Ignore common header/footer text that is not part of the lyrics:
  - Church names such as "Chiesa OLGIATE" or similar recurring labels.
  - Song titles repeated on every slide (e.g., "47 - Là nel ciel qui sulla terra"); detect by checking if the same line appears across all or most slides and matches the filename pattern.
- Recognize verse markers like "1° Strofa", "2° Strofa", "3° Strofa", etc. (case-insensitive). Use them to number verses accordingly (v1, v2, v3, ...) and remove the marker line itself from the lyrics content.

Title Extraction
- When deriving the song title from the filename, strip any leading number and separator (e.g., "47 - ", "10 ", "178 - "). The title should be the text after the number/separator (e.g., "47 - Là nel ciel qui sulla terra" → "Là nel ciel qui sulla terra").

Embedded Audio Handling
- Some slides may have an embedded music/audio file.
- Before extracting, check if an audio file with the same name as the PPTX (or matching the song title) already exists in the folder (e.g., .mp3, .wav, .m4a).
- If the audio file already exists, skip extraction.
- If not present, extract the embedded audio to the same folder, naming it with the song title (derived from the PPTX filename).

Block Detection
- Within each slide, if it have the label "Ritornello" or "Rit." (case-insensitive, leading spaces ignored), treat the remainder of that slide as a chorus block; drop the label line itself.
- Also infer choruses even without the label: if a block of lines repeats after multiple verses (alternating verse/chorus pattern) or appears identically on multiple slides, classify that repeated block as the next chorus (c1, c2, ...) in encounter order.
- Number choruses in encounter order: c1, c2, ...
- All other blocks are verses in encounter order: v1, v2, ...
- If a slide has both verse text and a chorus, emit two consecutive blocks: verse first, then chorus.

Line Handling Inside a Block
- Preserve line order; join lines with <br />. Remove lines that are empty after cleaning; do not emit empty blocks.

XML Output (one file per song)
- Output folder: Create a subfolder named "openLyrics" in the source folder for all generated XML files and extracted audio.
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

Conversion Log
- Maintain a log file (e.g., conversion_log.txt) in the output folder, updated after processing each PPT file.
- For each file, record:
  - Filename of the PPT processed.
  - Whether it was successfully converted to XML.
  - If it contained embedded media (audio/video).
  - If media was extracted (and to what filename).
  - Any errors or notes (e.g., skipped due to existing audio, failed extraction).

References
- OpenLyrics repository: https://github.com/openlyrics/openlyrics/tree/master
- OpenLyrics RelaxNG XML schema: https://github.com/openlyrics/openlyrics/blob/master/openlyrics-0.9.rng
- Python validator example: https://github.com/openlyrics/openlyrics/blob/master/tools/validate.py
- Validation documentation: https://docs.openlyrics.org/en/latest/validation.html
- Official documentation: https://docs.openlyrics.org/en/latest/contents.html
- Song example (complex.xml): https://github.com/openlyrics/openlyrics/blob/master/examples/complex.xml
