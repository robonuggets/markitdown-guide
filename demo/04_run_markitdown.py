"""Run MarkItDown across the Ember & Oak sample folder. Save .md outputs to _markdown/."""
from pathlib import Path
from markitdown import MarkItDown

ROOT = Path(__file__).parent
OUT = ROOT / "_markdown"
OUT.mkdir(exist_ok=True)

# Files to convert (skip the python scripts and the placeholder txt)
SKIP_NAMES = {"01_generate_demo.py", "04_run_markitdown.py", "02_generate_more_office.py",
              "03_generate_bulk.py", "05_ocr_image.py",
              "founder-voice-memo-NOTE.txt", "README.md"}
SKIP_DIRS  = {"_markdown"}

md = MarkItDown(enable_plugins=False)

results = []
for p in sorted(ROOT.iterdir()):
    if p.is_dir() and p.name in SKIP_DIRS: continue
    if not p.is_file() or p.name in SKIP_NAMES: continue
    try:
        result = md.convert(str(p))
        out_path = OUT / (p.stem + ".md")
        out_path.write_text(result.text_content, encoding="utf-8")
        results.append((p.name, len(result.text_content), "OK"))
    except Exception as e:
        results.append((p.name, 0, f"FAIL: {e}"))

print(f"\n{'File':<45} {'Chars':>8}  Status")
print("-" * 70)
for name, chars, status in results:
    print(f"{name:<45} {chars:>8}  {status}")

# Build a single master.md that concatenates everything — for a "drop the whole folder
# into a chat" demo
master_lines = ["# Ember & Oak — Full Folder as Markdown", "",
                "_All client files converted by MarkItDown into a single corpus._", ""]
for p in sorted(OUT.glob("*.md")):
    master_lines.append(f"\n---\n\n## {p.stem}\n")
    master_lines.append(p.read_text(encoding="utf-8"))
(OUT / "_master.md").write_text("\n".join(master_lines), encoding="utf-8")

total_chars = sum(r[1] for r in results)
print(f"\nTotal chars across all files: {total_chars:,}")
print(f"Master corpus: {OUT / '_master.md'}")
