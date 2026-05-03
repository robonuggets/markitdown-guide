"""Run MarkItDown with Claude vision on the candle PNG.

Setup: add this line to C:/ROBO/.secrets/.env, then run.
    ANTHROPIC_API_KEY=sk-ant-...

Cost: ~$0.002 per image with Haiku 4.5. This script does 1 image.
"""
import os
from pathlib import Path
from dotenv import load_dotenv
from openai import OpenAI
from markitdown import MarkItDown

load_dotenv(r"C:\ROBO\.secrets\.env")

key = os.environ.get("ANTHROPIC_API_KEY")
if not key:
    raise SystemExit("Add ANTHROPIC_API_KEY=... to C:/ROBO/.secrets/.env first")

client = OpenAI(
    base_url="https://api.anthropic.com/v1/",
    api_key=key,
)

md = MarkItDown(
    enable_plugins=False,
    llm_client=client,
    llm_model="claude-haiku-4-5-20251001",
)

ROOT = Path(__file__).parent
img_path = ROOT / "product-mockup-library-smoke.png"

print(f"Converting: {img_path.name}")
print(f"Model: claude-haiku-4-5-20251001 via Anthropic OpenAI-compat endpoint\n")

result = md.convert(str(img_path))

print("=" * 60)
print("MARKDOWN OUTPUT")
print("=" * 60)
print(result.text_content)
print("=" * 60)

# Save alongside the other markdown outputs, suffixed so we don't clobber
out = ROOT / "_markdown" / "product-mockup-library-smoke.WITH-VISION.md"
out.write_text(result.text_content, encoding="utf-8")
print(f"\nSaved: {out}")
