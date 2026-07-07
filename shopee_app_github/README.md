# Shopee Bulk Stock Update — Streamlit App

Sync available stock from the **MM Masterlist** to **Shopee** using your confirmed **SKU registry**.
Only the **Seller Stock** column is changed; the original Shopee template format is preserved.

## Run
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Upload (4 files)
1. **Masterlist Excel** — needs `Stock Type ID` and `Total` (Available Quantity) columns.
2. **Shopee bulk stock — File 1 of 2**
3. **Shopee bulk stock — File 2 of 2**
4. **SKU Registry Excel** — tabs: `Locked Matches`, `New Masterlist SKUs`, `Match Review`,
   `Skipped (SKU=NO or 0 Qty)`, `Accessories (not matched)`, `Not Selling in Shopee`, `Not on Shopee Yet`.

## Stock rules (matched by Variation ID)
| Source | Seller Stock |
|---|---|
| Locked Matches | Sum of Available Qty of all linked Masterlist SKUs (supports one variation → many SKUs) |
| Skipped | 0 |
| Accessories | 10 |
| Match Review / not in registry | unchanged |

Precedence: Locked → Skipped → Accessories → unchanged.

## Output
- Two updated Shopee `.xlsx` files (same layout) → upload in Seller Centre → Mass Update → Stock.
- **Validation summary** (rows changed per rule) and **Error report**:
  - Locked variations whose Masterlist ID isn't in the Masterlist (counted as 0).
  - Registry Variation IDs not found in the uploaded Shopee files.

## Notes
- Shopee exports contain a malformed `activePane` attribute; the app auto-repairs it on load.
- The Masterlist `Total` column is used as Available Quantity.
