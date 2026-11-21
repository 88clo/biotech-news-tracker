# Biotech News Tracker

A self-contained, week-over-week dashboard of biotech news, built from weekly
`.docx` notes. The live site is a single HTML file (`index.html`) — no server,
no build step needed to view it.

## Live site

Once GitHub Pages is enabled (Settings → Pages → Build from `main` / root), the
tracker is served at:

```
https://<user>.github.io/<repo>/
```

## How it works

- **`notes/`** — one `.docx` per week, named `Month_Year_Wn.docx`
  (e.g. `Nov_2026_W2.docx`). The label and week ordering come from the filename.
- **`parse_notes.py`** — converts the notes into the dashboard (stdlib only;
  needs `pandoc` on the PATH).
- **`template.html`** — the dashboard's design shell; data is injected into it.
- **`index.html`** — the built dashboard (this is what Pages serves).

## Adding a week

1. Add the week's `.docx` to `notes/`.
2. Commit and push.

The included GitHub Action (`.github/workflows/build.yml`) installs pandoc,
re-runs the parser, and commits the rebuilt `index.html` automatically — so the
live site updates on its own. To rebuild locally instead:

```bash
python3 parse_notes.py notes --html-template template.html --html-out index.html
```

## Using the dashboard

- **Week carousel** — the focused week sits center; neighbors flank it. Move with
  the arrows, the ← → keys, the week chips, or by clicking a side column.
- **Tabs** — Diseases, Modalities, Themes.
- **Area → disease → stories**, with diseases ordered consistently across weeks.
- **Three toggles** — Areas, Indications, Details (Details collapsed by default).
- **Follow an indication** — click a story (or an indication header) and every
  week aligns that indication on the same row, so you can read it across time.
- **Market badge** — stories with a market size ≥ $1B show a green badge; the
  company name is highlighted in yellow. Negative stories read in red.
- **Search** — type a disease for a month-by-month timeline across all weeks.

Detection rules are editable near the top of `parse_notes.py` (`BIG_MKT`,
`NEGATIVE_RX`, and the disease-canonicalization tables).
