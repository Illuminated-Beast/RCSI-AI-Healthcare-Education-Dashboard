# AI in Formal Healthcare Education - Evidence Dashboard

Scoping review dashboard for an MSc dissertation on AI use in formal healthcare
education. This repo contains the analysis and visualisation tooling only.
**The underlying research dataset (paper titles, authors, links, abstracts) is
never included here and never will be** - see [Data policy](#data-policy) below.

## What's here

| File | What it is | Contains research data? |
|---|---|---|
| `dashboard.html` | Static dashboard, pre-built from the locked 507-paper corpus. Shows aggregate counts and the verified evidence scores. | **Aggregate numbers only** - no paper titles, authors, or links. |
| `dynamic_dashboard.html` | Interactive dashboard. Drop in a `.xlsx` file matching the input spec and it analyses it live, entirely in your browser. | **No** - ships empty. Nothing is embedded; nothing you upload leaves your machine (see below). |
| `raw_intake_cleaner.html` | Converts a raw literature-search export into the format the dashboards expect. Auto-fills what has a documented rule; flags what needs human/model input. | **No** - pure client-side code, no data embedded. |
| `build_dashboard.py` | Python script that regenerates `dashboard.html` from a local input file + `dashboard_template.html`. For offline/CLI use. | **No** - script contains only the locked aggregate evidence constants, not per-paper rows. |
| `dashboard_template.html` | The HTML/CSS/JS shell `build_dashboard.py` injects data into. | No. |

## How it works

**`dynamic_dashboard.html` and `raw_intake_cleaner.html` do all parsing in the
browser**, using SheetJS to read the uploaded `.xlsx` and Chart.js to render
charts. There is no backend, no server, no upload endpoint. A file you drag
into the page is read by `FileReader` into browser memory and never
transmitted anywhere - this is a structural guarantee of the architecture
(GitHub Pages serves static files only; there is nothing to send data to),
not a promise layered on top of code that could exfiltrate data.

**Topic assignment is locked, not computed here.** The six topic categories
and their evidence scores (D1/D2/D3, composite, tier) come from a BERTopic
model (sentence-transformer embeddings + KMeans) trained once on the full
corpus and then frozen. Neither dashboard re-runs that model - they display
its already-verified output and, for new papers, validate that an uploaded
`Topic_Label` matches one of the six locked labels exactly. If it doesn't,
the file is rejected with a specific error, not silently misprocessed.

**`raw_intake_cleaner.html`** takes a raw search export (`Title, Authors,
Journal, Date, Link, Matched AI, Matched education, Abstracts`) and derives
five fields using documented, exact rules (year extraction, pre/post-ChatGPT
period split, abstract-length flag, an outcome-reporting keyword regex, and
source-database detection from the link domain). It deliberately does **not**
guess the two fields that require human judgment or the real model
(Classification, Topic_Label) - those come out flagged and must be filled in
before the dashboards will accept the file.

## Running locally

```bash
# static dashboard - just open it, no build step needed to view
open dashboard.html        # macOS
start dashboard.html       # Windows

# regenerate the static dashboard from a new locked input file
pip install pandas openpyxl
python build_dashboard.py
```

`dynamic_dashboard.html` and `raw_intake_cleaner.html` need no install -
open them directly in any modern browser.

## Data policy

This project analyses a scoping review corpus that is not publicly released.
**No file in this repository, at any point in its history, will contain
paper-level data** (titles, authors, links, or abstracts). Aggregate,
de-identified statistics (counts, percentages, evidence tier ratings) are
published deliberately as the research output. If you are looking for the
underlying dataset: it is not here, and requests should go through the
research supervisor, not this repository.

## License

Code in this repository should not be reused under any (LICENSE) unless
stated otherwise. This does not extend to any research findings, which
remain the author's and their institution's.
