"""
Core logic for the Projection vs Dispatch Variance system.

Given two workbooks:
  - file1 ("Projection" file): label columns (ItemType, ItemName, [UOMName], ...),
    then one column per site
  - file2 ("Dispatch" file): same shape, values may differ, headers may differ slightly

Produces one workbook, styled after the team's final reference format:

  - Row 1: bold white-on-blue banner over every "named" column (labels, sites, the
    repeated ItemName column). The strip above the Totals columns is left as a plain,
    unfilled merged box, since the Totals don't have a site name to show there.
  - Row 2: sub-header row. Any cell carrying real header text - "Projection" /
    "Dispatch" / "Variance" per matched site, or "Total Projection" / "Total Dispatch" /
    "Total Variance" - is filled yellow. Blank filler cells (spacer columns, unmatched
    sites' subheader, the now-empty cell under the ItemName repeat) stay light blue.
  - Every site (matched or not) is followed by one blank spacer column.
  - Sites found in the Projection file but with no match in the Dispatch file are kept as
    a single raw-value column (no Dispatch/Variance split) and are EXCLUDED from the Total
    columns.
  - A repeated ItemName column (header in row 1, left-aligned data) sits right before the
    totals, for readability on wide sheets.
  - Total Projection / Total Dispatch / Total Variance columns at the end.
  - Full thin grid lines throughout, with a heavier "medium" border framing the outside
    edge of the whole table.
  - Alternating row banding on data rows.
  - Column widths size to each column's own header text rather than a flat width.

The whole layout - number of site blocks, number of item rows, column widths - is computed
from the actual uploaded files. Nothing here assumes a fixed number of sites or rows.
"""

import io
import difflib
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

LABEL_HEADERS = {"itemtype", "itemname", "uomname"}
FUZZY_MATCH_THRESHOLD = 0.85

# ---- Styling ----
FONT_NAME = "Calibri"
FONT_SIZE = 11

HEADER_FILL = PatternFill("solid", fgColor="4472C4")        # row 1 banner (blue)
HEADER_FONT = Font(name=FONT_NAME, size=FONT_SIZE, bold=True, color="FFFFFF")

PLAIN_HEADER_FONT = Font(name=FONT_NAME, size=FONT_SIZE, bold=False, color="000000")

SUBHEADER_FILL = PatternFill("solid", fgColor="D9E1F2")     # row 2 filler (light blue)
LABEL_FILL = PatternFill("solid", fgColor="FFFF00")         # row 2 real labels (yellow)
SUBHEADER_FONT = Font(name=FONT_NAME, size=FONT_SIZE, bold=False, color="000000")

BAND_FILL = PatternFill("solid", fgColor="D9E1F2")          # alternating data-row banding

DATA_FONT = Font(name=FONT_NAME, size=FONT_SIZE)

CENTER = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center")


def _norm(s):
    return str(s).strip().lower() if s is not None else ""


def _read_sheet(file_bytes):
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    return wb.active


def _label_columns(ws):
    """[(col, header_text), ...] for every label column, in original left-to-right order."""
    cols = []
    for c in range(1, ws.max_column + 1):
        v = ws.cell(1, c).value
        if v is not None and _norm(v) in LABEL_HEADERS:
            cols.append((c, str(v).strip()))
    return cols


def _header_map(ws):
    """column index (1-based) -> header text, excluding label columns."""
    headers = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(1, c).value
        if v is not None and _norm(v) not in LABEL_HEADERS:
            headers[c] = str(v).strip()
    return headers


def _item_row_map(ws, name_col):
    """ItemName (normalized) -> row index, for row lookups regardless of order."""
    mapping = {}
    for r in range(2, ws.max_row + 1):
        name = ws.cell(r, name_col).value
        if name is not None:
            mapping[_norm(name)] = r
    return mapping


def _fuzzy_match(header, candidates):
    """Return the best candidate header matching `header` above threshold, or None."""
    best, best_score = None, 0.0
    for cand in candidates:
        score = difflib.SequenceMatcher(None, _norm(header), _norm(cand)).ratio()
        if score > best_score:
            best, best_score = cand, score
    if best_score >= FUZZY_MATCH_THRESHOLD:
        return best
    return None


def _border(row, col, last_row, last_col):
    """Thin grid everywhere, medium on the true outer edge of the whole table."""
    top = "medium" if row == 1 else "thin"
    bottom = "medium" if row == last_row else "thin"
    left = "medium" if col == 1 else "thin"
    right = "medium" if col == last_col else "thin"
    return Border(
        top=Side(style=top), bottom=Side(style=bottom),
        left=Side(style=left), right=Side(style=right),
    )


def build_variance_workbook(file1_bytes, file2_bytes):
    """Returns (xlsx_bytes, report_dict) where report_dict lists matched/unmatched sites."""
    proj_ws = _read_sheet(file1_bytes)
    disp_ws = _read_sheet(file2_bytes)

    proj_labels = _label_columns(proj_ws)          # [(col, "ItemType"), (col, "ItemName"), ...]
    disp_labels = _label_columns(disp_ws)
    proj_headers = _header_map(proj_ws)            # {col: site_name} in projection file
    disp_headers = _header_map(disp_ws)            # {col: site_name} in dispatch file

    proj_name_col = next((c for c, h in proj_labels if _norm(h) == "itemname"), None)
    disp_name_col = next((c for c, h in disp_labels if _norm(h) == "itemname"), None)

    if proj_name_col is None or disp_name_col is None:
        raise ValueError("Both files need an 'ItemName' column.")

    disp_item_rows = _item_row_map(disp_ws, disp_name_col)

    # Items, in the Projection file's original top-to-bottom order
    item_rows = []  # (proj_row, item_name, {label_header: value})
    for r in range(2, proj_ws.max_row + 1):
        name = proj_ws.cell(r, proj_name_col).value
        if name is None:
            continue
        label_values = {h: proj_ws.cell(r, c).value for c, h in proj_labels}
        item_rows.append((r, name, label_values))

    # Match each projection-file site to a dispatch-file site (exact, then fuzzy),
    # preserving the projection file's original left-to-right site order.
    proj_cols_sorted = sorted(proj_headers.keys())
    disp_header_values = list(disp_headers.values())

    site_plan = []
    matched_sites, unmatched_sites = [], []

    for col in proj_cols_sorted:
        site_name = proj_headers[col]
        disp_col = None
        for dcol, dname in disp_headers.items():
            if _norm(dname) == _norm(site_name):
                disp_col = dcol
                break
        if disp_col is None:
            match_name = _fuzzy_match(site_name, disp_header_values)
            if match_name:
                for dcol, dname in disp_headers.items():
                    if dname == match_name:
                        disp_col = dcol
                        break
        if disp_col is not None:
            matched_sites.append(site_name)
            site_plan.append({"name": site_name, "proj_col": col, "disp_col": disp_col, "matched": True})
        else:
            unmatched_sites.append(site_name)
            site_plan.append({"name": site_name, "proj_col": col, "disp_col": None, "matched": False})

    # ---- Lay out output columns: matched sites get 4 cols (Proj/Disp/Var/gap),
    # unmatched sites get 2 cols (value/gap). ----
    num_labels = len(proj_labels)
    next_col = num_labels + 1
    for block in site_plan:
        if block["matched"]:
            block["out_proj_col"] = next_col
            block["out_disp_col"] = next_col + 1
            block["out_var_col"] = next_col + 2
            block["out_gap_col"] = next_col + 3
            next_col += 4
        else:
            block["out_val_col"] = next_col
            block["out_gap_col"] = next_col + 1
            next_col += 2

    itemname_repeat_col = next_col
    total_proj_col = next_col + 1
    total_disp_col = next_col + 2
    total_var_col = next_col + 3
    max_col = total_var_col

    header_row = 1
    subheader_row = 2
    first_data_row = 3
    last_data_row = first_data_row + len(item_rows) - 1 if item_rows else first_data_row

    # ---- Build the workbook from scratch ----
    out_wb = Workbook()
    out_ws = out_wb.active
    out_ws.title = "Variance Report"

    # Label headers (row 1)
    for i, (_, header_text) in enumerate(proj_labels, start=1):
        out_ws.cell(header_row, i, header_text)

    # Site headers / sub-headers
    label_subheader_cols = set()  # row-2 cells that carry real header text -> yellow
    for block in site_plan:
        if block["matched"]:
            c0 = block["out_proj_col"]
            out_ws.merge_cells(start_row=header_row, start_column=c0, end_row=header_row, end_column=c0 + 2)
            out_ws.cell(header_row, c0, block["name"])
            out_ws.cell(subheader_row, c0, "Projection")
            out_ws.cell(subheader_row, c0 + 1, "Dispatch")
            out_ws.cell(subheader_row, c0 + 2, "Variance")
            label_subheader_cols.update([c0, c0 + 1, c0 + 2])
        else:
            out_ws.cell(header_row, block["out_val_col"], block["name"])

    item_name_header = next((h for _, h in proj_labels if _norm(h) == "itemname"), "ItemName")
    # Repeated ItemName label now lives in row 1, alongside the other named headers
    out_ws.cell(header_row, itemname_repeat_col, item_name_header)

    # Totals: row 1 above them is a plain, unfilled merged box (no site name to show there)
    out_ws.merge_cells(start_row=header_row, start_column=total_proj_col, end_row=header_row, end_column=total_var_col)
    out_ws.cell(subheader_row, total_proj_col, "Total Projection")
    out_ws.cell(subheader_row, total_disp_col, "Total Dispatch")
    out_ws.cell(subheader_row, total_var_col, "Total Variance")
    label_subheader_cols.update([total_proj_col, total_disp_col, total_var_col])

    # Style rows 1-2 across the full width
    for c in range(1, max_col + 1):
        top = out_ws.cell(header_row, c)
        top.border = _border(header_row, c, last_data_row, max_col)
        top.alignment = CENTER
        top.font = HEADER_FONT
        top.fill = HEADER_FILL

        sub = out_ws.cell(subheader_row, c)
        sub.font = SUBHEADER_FONT
        sub.border = _border(subheader_row, c, last_data_row, max_col)
        sub.alignment = CENTER
        sub.fill = LABEL_FILL if c in label_subheader_cols else SUBHEADER_FILL

    # ---- Data rows ----
    proj_cols_for_total = [b["out_proj_col"] for b in site_plan if b["matched"]]
    disp_cols_for_total = [b["out_disp_col"] for b in site_plan if b["matched"]]
    var_cols_for_total = [b["out_var_col"] for b in site_plan if b["matched"]]

    text_cols = set(range(1, num_labels + 1)) | {itemname_repeat_col}

    for i, (proj_row, item_name, label_values) in enumerate(item_rows):
        out_row = first_data_row + i

        for col_idx, (_, header_text) in enumerate(proj_labels, start=1):
            out_ws.cell(out_row, col_idx, label_values.get(header_text))

        drow = disp_item_rows.get(_norm(item_name))

        for block in site_plan:
            if block["matched"]:
                proj_val = proj_ws.cell(proj_row, block["proj_col"]).value
                disp_val = disp_ws.cell(drow, block["disp_col"]).value if drow else None
                out_ws.cell(out_row, block["out_proj_col"], proj_val)
                out_ws.cell(out_row, block["out_disp_col"], disp_val)
                pl = get_column_letter(block["out_proj_col"])
                dl = get_column_letter(block["out_disp_col"])
                out_ws.cell(out_row, block["out_var_col"], f"={dl}{out_row}-{pl}{out_row}")
            else:
                proj_val = proj_ws.cell(proj_row, block["proj_col"]).value
                out_ws.cell(out_row, block["out_val_col"], proj_val)

        out_ws.cell(out_row, itemname_repeat_col, item_name)

        if proj_cols_for_total:
            out_ws.cell(out_row, total_proj_col,
                        "=" + "+".join(f"{get_column_letter(c)}{out_row}" for c in proj_cols_for_total))
            out_ws.cell(out_row, total_disp_col,
                        "=" + "+".join(f"{get_column_letter(c)}{out_row}" for c in disp_cols_for_total))
            out_ws.cell(out_row, total_var_col,
                        "=" + "+".join(f"{get_column_letter(c)}{out_row}" for c in var_cols_for_total))
        else:
            out_ws.cell(out_row, total_proj_col, 0)
            out_ws.cell(out_row, total_disp_col, 0)
            out_ws.cell(out_row, total_var_col, 0)

        # Banding: 2nd, 4th, ... data row shaded
        band = (i % 2 == 1)
        for c in range(1, max_col + 1):
            cell = out_ws.cell(out_row, c)
            cell.font = DATA_FONT
            cell.border = _border(out_row, c, last_data_row, max_col)
            cell.alignment = LEFT if c in text_cols else CENTER
            if band:
                cell.fill = BAND_FILL

    # ---- Column widths: size to each column's own header text, not a flat number ----
    item_name_col_idx = next((i for i, (_, h) in enumerate(proj_labels, start=1) if _norm(h) == "itemname"), 2)
    max_name_len = max([len(str(n)) for _, n, _ in item_rows] + [len(item_name_header)])
    name_width = min(max(max_name_len + 4, 20), 60)

    def header_width(text):
        return max(len(str(text)) + 2, 8)

    for i, (_, header_text) in enumerate(proj_labels, start=1):
        width = name_width if i == item_name_col_idx else header_width(header_text)
        out_ws.column_dimensions[get_column_letter(i)].width = width

    for block in site_plan:
        if block["matched"]:
            out_ws.column_dimensions[get_column_letter(block["out_proj_col"])].width = header_width("Projection")
            out_ws.column_dimensions[get_column_letter(block["out_disp_col"])].width = header_width("Dispatch")
            out_ws.column_dimensions[get_column_letter(block["out_var_col"])].width = header_width("Variance")
        else:
            out_ws.column_dimensions[get_column_letter(block["out_val_col"])].width = header_width(block["name"])

    out_ws.column_dimensions[get_column_letter(itemname_repeat_col)].width = name_width
    out_ws.column_dimensions[get_column_letter(total_proj_col)].width = header_width("Total Projection")
    out_ws.column_dimensions[get_column_letter(total_disp_col)].width = header_width("Total Dispatch")
    out_ws.column_dimensions[get_column_letter(total_var_col)].width = header_width("Total Variance")

    buf = io.BytesIO()
    out_wb.save(buf)
    buf.seek(0)

    report = {
        "matched_sites": matched_sites,
        "unmatched_sites": unmatched_sites,
        "total_sites_in_projection_file": len(proj_headers),
        "total_sites_in_dispatch_file": len(disp_headers),
    }
    return buf.read(), report