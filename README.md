# SpineResearch Studio Pro

A full-stack research workflow prototype for spine surgery manuscripts.

## What it does

- Accepts rough study notes and multiple files
- Parses CSV, Excel, DOCX, PDF text, TXT/MD/JSON, and ZIP bundles containing supported files
- Audits datasets for variables, missingness, plausibility, follow-up/outcome candidates, and data readiness
- Suggests feasible research questions and hypotheses centered on the user's initial idea
- Runs real exploratory statistical outputs from structured datasets:
  - Table 1
  - missingness table/figure
  - paired pre/post screening
  - group comparison table when exposure/outcome are selected
  - regression feasibility warnings
- Generates a journal-style manuscript package after analysis
- Exports DOCX manuscript and DOCX statistical report
- Optionally uses `OPENAI_API_KEY` for more polished manuscript drafting, while constraining the model to verified analysis outputs

## Run locally

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

Open: http://localhost:8000

## Optional AI

Set an environment variable before running:

```bash
export OPENAI_API_KEY="sk-..."
```

Without the key, the app still performs parsing, audit, analysis, tables, figures, and structured manuscript drafting without invented citations.

## Deployment note

This is a backend application. Deploy to Render, Railway, Fly.io, DigitalOcean, or a server that supports Python packages. This is not a no-dependency static Vercel build.

## Guardrails

- Does not invent p-values, sample sizes, or references
- Labels placeholders and unresolved items
- Flags analyses that are not feasible from the uploaded data
- Manuscript results are based on computed tables only
