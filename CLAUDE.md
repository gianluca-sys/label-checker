# Healf Label Checker

Compares two nutrition label PDFs using Claude Vision and writes a formatted Excel report.

## Quick start

```bash
cd C:\Users\GianlucaLyons\label-checker
python label_compare.py --label1 labels/current.pdf --label2 labels/new.pdf --sku "MOM-PROT-001"
```

The report is saved to `label_comparison_report.xlsx` in this folder.
On the first run it creates the file. Every subsequent run appends a row to the Change Log
and creates or replaces a tab for the SKU.

## Arguments

| Argument   | Required | Description                                           |
|------------|----------|-------------------------------------------------------|
| `--label1` | Yes      | Path to the current / reference label PDF             |
| `--label2` | Yes      | Path to the new / supplier label PDF                  |
| `--sku`    | Yes      | Product SKU — becomes the tab name in the report      |
| `--output` | No       | Custom output path (default: label_comparison_report.xlsx) |

## Folder layout

```
label-checker/
├── label_compare.py
├── CLAUDE.md
├── labels/                          ← drop PDFs here
│   ├── current.pdf
│   └── new_supplier.pdf
└── label_comparison_report.xlsx     ← auto-created / updated
```

## What it checks (word for word)

- Brand name and product name
- Net weight
- Serving size and servings per container
- Calories
- Every nutrient — amount and % Daily Value, in the exact label order
- Nutrients added or removed from the facts table
- Ingredients — full text including order (any reordering is flagged as a change)
- Allergen / "Contains" statement
- Health claims, certifications, storage instructions, and any other label text

## Output tabs in the Excel report

| Tab          | Contents                                                          |
|--------------|-------------------------------------------------------------------|
| Change Log   | Every run: timestamp, SKU, file names, change count, MATCH/CHANGES |
| [SKU name]   | Side-by-side field comparison. Red = changed. Green = match.      |
| How To Use   | Quick reference guide                                             |

## Requirements

- Python 3.8+
- `anthropic` and `openpyxl` packages installed
- `ANTHROPIC_API_KEY` set as a User environment variable

## Running from inside Claude Code

Just tell Claude what you want to compare, for example:

> Compare labels/current.pdf and labels/new_supplier.pdf for SKU MOM-PROT-001

Claude will run the script, print the differences, and confirm the report was saved.
