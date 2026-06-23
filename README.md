# Pool Coder — Web App

Upload an Excel sheet → get back a finished presentation with teams coded
and pools balanced, in the exact format of `template.pptx`.

## Files

- `app.py` — the web server (Flask). One page, one upload form, one download.
- `pool_coder.py` — the coding + balancing logic (the part that must never guess).
- `read_excel.py` — reads the buyer's `.xlsx` into the format `pool_coder.py` needs.
- `build_presentation.py` — slots the result into `template.pptx`, preserving formatting exactly.
- `template.pptx` — the presentation template. Replace this with whatever template you're using.
- `requirements.txt` — Python dependencies.

## Run it locally

```bash
pip install -r requirements.txt
python3 app.py
```
Then open http://localhost:5000 — upload an Excel file with columns
`School`, `Category`, `Teams` in the first row, set the number of pools,
click Generate. You'll get a `.pptx` download.

## Deploy on Render (free to start, $7/mo for always-on)

1. Push this folder to a GitHub repo.
2. On Render: New → Web Service → connect the repo.
3. Build command: `pip install -r requirements.txt`
4. Start command: `gunicorn app:app`
5. Deploy. Render gives you a public URL — that's what you hand to a buyer.

## Swapping in a different buyer's template

`build_presentation.py` finds pool boxes by looking for text matching
"Pool X Teams ... so far ...". If a buyer's template uses different wording
or layout, you'll need to either:
- adjust their template's header text to match that pattern, or
- adjust the `POOL_HEADER_RE` regex in `build_presentation.py` to match theirs.

Everything else (header text, topic descriptions, branding, fonts) is left
completely untouched — only the team lists inside matching shapes are replaced.

## What's deliberately NOT AI-generated

The school→team coding and the pool-balancing are both plain, deterministic
code — no LLM involved, no chance of duplicate codes or miscounted teams.
If you later want a chat-style "agent" wrapper on top of this (so a buyer can
just describe their list in plain English instead of filling Excel columns),
that part *would* use an LLM — but only to restructure messy text into the
columns this code expects, never to do the counting itself.
