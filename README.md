---
title: kf-pdf-liberator
emoji: 🚀
colorFrom: green
colorTo: blue
sdk: streamlit
sdk_version: 1.44.1
app_file: app.py
pinned: false
---

# KF-PDFLiberator

> Extract tables from PDFs into Excel/CSV — no more manual retyping.

## The Problem

52% of AP (Accounts Payable) teams spend 10+ hours per week on invoice processing. 37% cite manual data entry as their biggest challenge. Retyping PDF tables into Excel is one of the most universally hated office tasks.

## How It Works

1. Upload a PDF containing tables
2. Choose extraction method (auto-detect, with borders, or without borders)
3. Preview extracted tables
4. Download as Excel (.xlsx) or CSV — individual tables or all at once

## Libraries Used

- **pdfplumber** — PDF table extraction with cell-level coordinate detection
- **openpyxl** — Excel file generation
- **pandas** — Data manipulation and CSV export

## Development

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deployment

Hosted on [Hugging Face Spaces](https://huggingface.co/spaces/mitoi/kf-pdf-liberator).

---

Part of the [KaleidoFuture AI-Driven Development Research](https://kaleidofuture.com) — proving that everyday problems can be solved with existing libraries, no AI model required.
