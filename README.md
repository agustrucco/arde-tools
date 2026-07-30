# ARDE Tools

Python utilities for managing sales and inventory in an Excel/Google Sheets workbook for **ARDE**, an Argentine artisanal jewelry brand.

The system lives in a single `.xlsx` file shared on Google Drive, used daily to log sales, manage custom orders, and track stock across 80+ SKUs.

---

## Commands

```bash
python arde_tools.py styling              # apply consistent colors to all sheets
python arde_tools.py fix_formulas         # regenerate VENDIDOS formulas in STOCK
python arde_tools.py add_month "SEP 26"   # create a new month sheet ready to use
python arde_tools.py stock_report         # print stock status to console
python arde_tools.py fix_desc_stock       # normalize TRUE/FALSE → SI/NO in DESC STOCK
python arde_tools.py validate             # detect entries with typos vs STOCK catalog
python arde_tools.py validate --fix       # detect AND auto-correct typos
```

---

## Workbook structure

| Sheet | Description |
|---|---|
| `STOCK` | Master catalog — one row per SKU (PIEZA + MODELO + METAL). Tracks initial stock, sold units (formula), losses, workshop, and local stock. |
| `ENERO 26`, `FEB 26`, … | Monthly sales sheets — one row per sale. Drives VENDIDOS formula via COUNTIFS. |
| `ENCARGOS` | Custom orders (made-to-order pieces). |
| `TAREAS` | Task board — team leaves requests, Agustín replies in RESPUESTA column. |
| `DASHBOARD` | Visual summary — auto-calculated from STOCK. |

---

## Setup

```bash
pip install -r requirements.txt
```

The included `NUEVO ARDE AGUSTIN.xlsx` is a **template with sample data**. Replace it with your actual workbook or run `create_template.py` to regenerate it.

---

## Key design decisions

- **PIEZA, MODELO, METAL must match exactly** between STOCK and monthly sheets — any mismatch silently breaks the COUNTIFS formulas (no error shown).
- **DESC STOCK = "SI"** means the unit was physically dispatched and counts against stock. "NO" = reserved, on hold, or an adjustment entry.
- **Compatible with Google Sheets** — all formulas use COUNTIFS/SUMIFS/INDEX-MATCH. No Excel-only functions.
- When adding a new month, always run `fix_formulas` afterwards so VENDIDOS includes the new sheet.
