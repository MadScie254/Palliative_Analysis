"""Build a comprehensive NVivo-friendly aggregate text file from all transcript DOCX files.

Current situation:
 - `AllTranscripts_NVivo.txt` only contains the first two transcripts and a placeholder.
 - Each transcript now exists as a valid `.docx` plus a backup `.docx.orig.txt` (original plaintext with ** markers).

Strategy:
 1. For each `transcript_PATIENT_XXX.docx` in numeric order:
    a. Prefer the corresponding `.orig.txt` backup (preserves ** markers as originally authored)
    b. Fallback: extract plain text from the DOCX (reconstruct bold paragraphs with ** if fully bold)
 2. Wrap each transcript with a clear delimiter line: `====TRANSCRIPT: PATIENT_XXX====`
 3. Append a consistent END section footer with generation timestamp and counts.

This produces a single UTF-8 text file suitable for import into qualitative analysis tools (e.g., NVivo / Atlas.ti). Bold markers are left in markdown-like form (`**bold**`) allowing downstream conversion or emphasis detection.

Idempotent: Running again overwrites the aggregate file.
"""

from __future__ import annotations
from pathlib import Path
from datetime import datetime
import re
import sys

try:
    from docx import Document  # type: ignore
except Exception:  # pragma: no cover - dependency issue
    Document = None  # type: ignore

ROOT = Path(__file__).parent
AGGREGATE_NAME = "AllTranscripts_NVivo.txt"

TRANSCRIPT_PATTERN = re.compile(r"transcript_(PATIENT_\d{3})\.docx$")


def iter_transcript_paths():
    files = []
    for p in ROOT.glob("transcript_PATIENT_*.docx"):
        m = TRANSCRIPT_PATTERN.search(p.name)
        if not m:
            continue
        pid = m.group(1)
        files.append((pid, p))
    # Sort numerically by the numeric portion
    files.sort(key=lambda t: int(t[0].split('_')[1]))
    return files


def read_from_backup(orig_path: Path) -> str | None:
    if orig_path.exists():
        try:
            return orig_path.read_text(encoding="utf-8")
        except Exception:
            return None
    return None


def extract_text_from_docx(docx_path: Path) -> str:
    if Document is None:
        raise RuntimeError("python-docx not available; cannot extract text")
    doc = Document(str(docx_path))
    lines: list[str] = []
    for para in doc.paragraphs:
        text = para.text.rstrip()
        if not text:
            lines.append("")
            continue
        # Determine if entire paragraph is bold (all runs bold & non-empty)
        runs = para.runs
        if runs and all((r.bold or False) for r in runs if r.text.strip()):
            text = f"**{text}**"
        lines.append(text)
    return "\n".join(lines)


def build_aggregate():
    entries = []
    transcripts = iter_transcript_paths()
    if not transcripts:
        print("No transcript files found.", file=sys.stderr)
        return 1

    for patient_id, path in transcripts:
        backup = path.with_suffix(path.suffix + ".orig.txt")
        content = read_from_backup(backup)
        if content is None:
            # Fallback to docx extraction
            content = extract_text_from_docx(path)

        # Ensure header presence (avoid duplicating if already has delimiter)
        header_line = f"====TRANSCRIPT: {patient_id}===="
        if header_line not in content.splitlines()[0:3]:
            content = f"{header_line}\n\n{content.strip()}\n"
        entries.append(content.strip() + "\n")

    footer = (
        "---\n\nEND OF ALL TRANSCRIPTS\n\n" \
        f"Generated: {datetime.utcnow():%Y-%m-%d %H:%M UTC}\n" \
        f"Total Patients: {len(transcripts)} ({', '.join(pid for pid, _ in transcripts)})\n" \
        "Source: palliative_data.csv\n" \
        "Format: Appendix I questionnaire structure with Esther transcript style\n"
    )

    aggregate_text = "\n".join(entries) + "\n" + footer
    out_path = ROOT / AGGREGATE_NAME
    out_path.write_text(aggregate_text, encoding="utf-8")
    print(f"Wrote aggregate file: {out_path} ({len(transcripts)} transcripts)")
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(build_aggregate())
