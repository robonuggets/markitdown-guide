# Install Guide — Flatten Any Folder Into a Browsable Knowledge Graph

Companion to the video. Three small tools, four prompts, and you'll go from a messy client folder of Word/Excel/PowerPoint/PDFs to a single Markdown corpus you can drop into any LLM — plus an interactive force-directed graph to browse it.

---

## What you'll build

1. A flattened **Markdown copy** of every file in a folder (docx, xlsx, pptx, pdf, png, eml, html, csv, json, etc.)
2. An interactive **force-directed graph** of those files in a single HTML page — click a bubble, see its content
3. A repeatable workflow you can run on any folder: client work, research, your own Drive

---

## The stack

| Tool | What it does | Link |
|---|---|---|
| **MarkItDown** | Microsoft's converter — any file → clean Markdown | https://github.com/microsoft/markitdown |
| **Claude Code** (or any agent) | Runs the prompts and builds the HTML | https://www.anthropic.com/claude-code |
| **personalise skill** | Adapts any external input (this guide, a tutorial, a tool) to *your* context | See Step 4 below |
| **Karpathy's LLM Wiki** | Why "LLM-readable" formats matter — required reading | https://gist.github.com/karpathy/442a6bf55591 |

---

## Step 1 — Install MarkItDown

```bash
pip install 'markitdown[all]'
```

The `[all]` extras include PDF, audio transcription, and image OCR support. If you only need Office docs, `pip install markitdown` is enough.

Check it works:

```bash
markitdown sample.docx > sample.md
```

---

## Step 2 — Flatten your folder

Point MarkItDown at every file in your folder and write the converted Markdown to a sibling `_markdown/` folder. Save this as `flatten.py` next to your files:

```python
from pathlib import Path
from markitdown import MarkItDown

ROOT = Path(__file__).parent
OUT = ROOT / "_markdown"
OUT.mkdir(exist_ok=True)

SKIP = {"flatten.py", "README.md"}
md = MarkItDown(enable_plugins=False)

for p in sorted(ROOT.iterdir()):
    if not p.is_file() or p.name in SKIP or p.name.startswith("_"):
        continue
    try:
        result = md.convert(str(p))
        (OUT / (p.stem + ".md")).write_text(result.text_content, encoding="utf-8")
        print(f"OK   {p.name}  ->  {len(result.text_content):,} chars")
    except Exception as e:
        print(f"FAIL {p.name}  -  {e}")
```

Run it:

```bash
python flatten.py
```

You'll get a `_markdown/` folder with one `.md` per source file. That's your corpus — you can now paste it into any LLM and ask questions across the whole folder.

> **Why this matters:** Office files are zipped XML archives. LLMs can't read them natively — you get binary noise. MarkItDown is the bridge. Same idea as Karpathy's LLM Wiki gist linked above: design your data so models can actually consume it.

---

## Step 3 — Visualise as a force-directed graph

Three prompts, copy-paste in order into Claude Code (or any coding agent). Each one builds on the last.

### Prompt 1 — build the graph

```
[path to your _markdown folder]

Create a one-page HTML slide that shows a force-directed graph similar to
Obsidian's interface, where the size of each bubble corresponds to the content
length of each MD file. Make it interactive — clicking a node should open that
MD file's content in a side panel so I can inspect it. Use D3 v7 + marked.js
from CDN. Embed the markdown content inline so it works fully offline (no
file:// fetch issues).
```

You'll get a working HTML file — open it in a browser, click bubbles, browse content.

### Prompt 2 — restyle dark + add gravity

```
Restyle with a polished dark theme:
- bg #000, surface #0a0a0a, border #1a1a1a, text #e0e0e0, muted #888
- Two accents: teal #50e3c2 and orange #ff6b1a
- Font: Outfit from Google Fonts
- Group bubble colours — pick five that read well on black
  (e.g. orange, teal, blue, amber, slate)

Add gravitational pull between nodes: a custom force on every D3 tick where
every pair attracts via F = G·m₁·m₂/r², with mass = bubble radius (so bigger
files pull harder). Layer it on top of charge/link/collide forces and weaken
charge so gravity dominates at medium range.
```

Now the graph clusters naturally — heavy files become anchors, related files orbit them.

### Prompt 3 — translucent bubbles with icons inside

```
Make the circles translucent (fill-opacity ~0.22 with a coloured stroke at
0.85, so they look like glass with glowing rims). On hover bump fill to 0.42,
on active state 0.5 with a contrasting ring.

Drop a simple Lucide-style stroke icon centred in every bubble, scaled to
r × 1.05 / 24 so it grows with the node. Pick an icon per group that matches
its meaning (e.g. megaphone for marketing, dollar for finance, package for
ops, user for people, document for other).
```

That's it — you now have the graph from the video.

---

## Step 4 — Personalise it for your world

The whole point of this workflow is to make *your* files queryable. The **personalise skill** takes any external input — a tutorial like this one, a tool, a framework, a prompt — and adapts it to your specific context (your business, your stack, your voice, your goals).

**What it does:** classifies the input, pulls only the context it needs from your setup, then returns a focused recommendation: take rate (high/medium/low/skip), what to keep, what to change, where it slots in your existing tools.

**To set it up** as a Claude Code skill, create `~/.claude/skills/personalise/SKILL.md` with frontmatter:

```yaml
---
name: personalise
description: Take any external input and recommend how to adapt it to my setup — workspace, business, stack, voice, goals. Triggers on "personalise this", "/personalise [thing]".
---
```

Then write the body to point at *your* context files (your CLAUDE.md, your goals doc, your tool inventory). The pattern is the value — the contents are yours.

**Use it on this guide:** once installed, run `/personalise [paste this URL]` and the skill will tell you which steps matter for your workflow vs which to skip.

---

## Further reading

**Andrej Karpathy — "On the LLM-readability of files"**
https://gist.github.com/karpathy/442a6bf55591

The thinking behind why MarkItDown exists. If your data isn't designed for LLMs to read, you're leaving capability on the table. Worth the 5 minutes.

---

## Recap

```
1. pip install 'markitdown[all]'
2. Run flatten.py on your folder       →  _markdown/
3. Three prompts in your coding agent  →  graph.html
4. Personalise the workflow            →  fits your world
```

Four moves. Any folder becomes a browsable, queryable knowledge base.

---

*Built with MarkItDown by Microsoft (MIT). Workflow originated by RoboNuggets.*
