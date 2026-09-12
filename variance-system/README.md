# Dispatch Variance Builder

A small local web app that does exactly what we built manually, automatically:
upload a **Projection** file and a **Dispatch** file (same shape as your
`demo_wps.xlsx` / `demo_2_wps.xlsx`), and it hands back one workbook with:

- Each site's header merged into a single cell (site name only)
- 3 sub-columns per site: **Projection | Dispatch | Variance** (`= Dispatch - Projection`)
- Header/site names matched even with minor spelling differences (e.g. `HACKET` vs `HACKETT`)
- Rows matched by **Item Name**, not row position — so the two files don't need to be in the same order
- **Total Projection / Total Dispatch / Total Variance** columns at the end

## Setup

```bash
cd backend
pip install -r requirements.txt
python app.py
```

Then open **http://127.0.0.1:5000** in your browser.

## Using it

1. Drop your Projection file into the first box, Dispatch file into the second.
2. Click **Build variance report**.
3. It shows which sites matched (and flags any that didn't — check for
   spelling differences or a site missing from one of the files).
4. Click **Download variance_report.xlsx**.

## How matching works (backend/processor.py)

- `ItemType`, `ItemName`, `UOMName` are treated as label columns — everything else
  is treated as a site.
- Sites are matched first by exact name, then by fuzzy string match (handles
  small typos automatically — tune `FUZZY_MATCH_THRESHOLD` in `processor.py`
  if it's too loose/strict for your data).
- Items (rows) are matched by `ItemName`, so row order differences between the
  two files don't cause mismatched data.
- Variance is written as a **live Excel formula**, not a hardcoded number — it
  recalculates if you edit either the Projection or Dispatch value.

## Notes

- This runs Flask's built-in dev server, fine for local/team use on one machine.
  For hosting it for a wider team, put it behind a real WSGI server (gunicorn, etc.)
  — ask me if you want that set up.
- Max upload size is 20 MB per file (adjustable in `app.py`).
