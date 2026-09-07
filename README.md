# Retail Inventory Automation

Python CLI tool for managing sales, inventory, and daily operations of an artisanal retail business — built with `openpyxl` over a single shared Excel workbook.

## Background

The business ran its entire operation (stock tracking, sales logging, custom orders, team task board) through a manually maintained Excel file shared on Google Drive. As the product catalog grew past 80 SKUs and monthly sales volume increased, manually maintaining formulas, dropdowns, and sheet styling became error-prone and time-consuming.

This tool automates the most repetitive and fragile operations, making the workbook reliable and easy to extend.

## Impact

- **Formula errors eliminated**: the `validate` command detects product name mismatches that previously caused COUNTIFS to silently return 0, understating sold units and overstating available stock
- **New months in seconds**: adding a new monthly sales sheet went from a 10-15 minute manual copy-paste process to a single CLI command
- **Dropdown consistency enforced**: dropdowns are regenerated from the live product catalog, not hardcoded — new products appear automatically
- **Styling applied in one command**: header colors, alternating row fills, and status-based highlights applied consistently across all sheets without manual work

## Commands

```bash
python inventory_tools.py styling              # apply consistent colors to all sheets
python inventory_tools.py fix_formulas         # regenerate VENDIDOS (units sold) formulas
python inventory_tools.py add_month "SEP 26"   # create a new monthly sales sheet, ready to use
python inventory_tools.py stock_report         # print stock status to console
python inventory_tools.py fix_desc_stock       # normalize TRUE/FALSE → SI/NO
python inventory_tools.py validate             # detect entries with typos vs product catalog
python inventory_tools.py validate --fix       # detect AND auto-correct typos
python inventory_tools.py fix_dropdowns        # regenerate all dropdowns from current catalog
python inventory_tools.py reset_desc_stock     # reset all dispatched flags (for stock recount)
```

## Workbook structure

| Sheet | Purpose |
|---|---|
| `STOCK` | Master catalog — one row per SKU (item type + model + material). Tracks initial stock, sold units (formula-driven), losses, workshop units, and available local stock. |
| `ENERO 26`, `FEB 26`, … | Monthly sales sheets — one row per sale. Each drives the VENDIDOS formula via COUNTIFS. |
| `ENCARGOS` | Custom made-to-order pieces — tracked separately from stock. |
| `TAREAS` | Async task board — staff leaves requests, operator replies in RESPUESTA column. |
| `DASHBOARD` | Visual summary — auto-calculated from STOCK. |

## How stock calculation works

```
VENDIDOS = COUNTIFS across all monthly sheets WHERE
  PIEZA matches AND MODELO matches AND METAL matches AND DESC_STOCK = "SI"

STOCK LOCAL = STOCK INICIAL - VENDIDOS - PERDIDOS - TALLER - FERIA
```

`DESC STOCK = "SI"` means the unit was physically dispatched and counts against inventory.
`"NO"` means the row is reserved, on hold, or a non-dispatched adjustment entry.

## Key design decisions

- **PIEZA + MODELO + METAL must match exactly** between STOCK and monthly sheets — any mismatch silently returns 0 in COUNTIFS with no error. The `validate` command detects this using `difflib` fuzzy matching and can auto-correct.
- **Compatible with Google Sheets** — all formulas use COUNTIFS/SUMIFS. No Excel-only functions.
- **Column mapping is centralized** in `STOCK_COLS` and `DESC_STOCK_COL` — updating the column layout requires changing one dict, not dozens of hardcoded references.
- **New months auto-detected** — `_month_sheets()` scans sheet names for year patterns so newly added months are picked up by `fix_formulas` without touching the source code.

## Setup

```bash
pip install openpyxl
```

Run `python create_template.py` to generate a sample workbook with fictional data.

## Tech

- Python 3.x
- [openpyxl](https://openpyxl.readthedocs.io/) for reading and writing `.xlsx`
- `difflib` (stdlib) for fuzzy matching in the `validate` command