"""
Core logic for the Projection vs Dispatch Variance system.

Given two workbooks:
  - file1 ("Projection" file): ItemType, ItemName, [UOMName], then one column per site
  - file2 ("Dispatch" file): same shape, values may differ, headers may differ slightly

Produces one workbook where, for every site found in BOTH files:
  - a merged header with just the site name
  - 3 sub-columns: Projection | Dispatch | Variance (= Dispatch - Projection)
  - 3 gap columns are left after most site groups for manual notes (matches original ask)
  - Total Projection / Total Dispatch / Total Variance columns at the very end
"""

import io
import difflib
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font

LABEL_HEADERS = {"itemtype", "itemname", "uomname"}
FUZZY_MATCH_THRESHOLD = 0.85


def _norm(s):
    return str(s).strip().lower() if s is not None else ""


def _read_sheet(file_bytes):
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    return wb.active


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


def _find_label_col(ws, label):
    for c in range(1, ws.max_column + 1):
        if _norm(ws.cell(1, c).value) == label:
            return c
    return None


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


def build_variance_workbook(file1_bytes, file2_bytes):
    """Returns (xlsx_bytes, report_dict) where report_dict lists matched/unmatched sites."""
    proj_ws = _read_sheet(file1_bytes)
    disp_ws = _read_sheet(file2_bytes)

    proj_headers = _header_map(proj_ws)          # {col: site_name} in projection file
    disp_headers = _header_map(disp_ws)          # {col: site_name} in dispatch file

    proj_name_col = _find_label_col(proj_ws, "itemname")
    disp_name_col = _find_label_col(disp_ws, "itemname")
    disp_item_rows = _item_row_map(disp_ws, disp_name_col)

    # Build fresh output workbook starting from the projection file (keeps its label columns)
    out_wb = load_workbook(io.BytesIO(file1_bytes))
    out_ws = out_wb.active
    max_row = out_ws.max_row
    max_col = out_ws.max_column

    # Match each projection-file site to a dispatch-file site (exact, then fuzzy)
    disp_header_values = list(disp_headers.values())
    matched_sites = []      # (proj_col, proj_site_name, disp_col)
    unmatched_sites = []    # proj_site_name with no match

    proj_cols_sorted = sorted(proj_headers.keys())
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
            matched_sites.append((col, site_name, disp_col))
        else:
            unmatched_sites.append(site_name)

    # Insert 3 blank columns after EVERY site column (right to left, so indices stay valid)
    for col in reversed(proj_cols_sorted):
        out_ws.insert_cols(col + 1, amount=3)

    # After insertion, each site's own column shifts right by 3 for every OTHER site
    # that sits to its left (each of those got 3 fresh columns inserted before this one).
    # rank = position of this site among all sites, sorted ascending by original column.
    rank_by_col = {col: i for i, col in enumerate(proj_cols_sorted)}
    site_groups = []  # (actual_col, matched_col, variance_col, site_name)
    for col, site_name, disp_col in matched_sites:
        final_col = col + 3 * rank_by_col[col]
        site_groups.append((final_col, final_col + 1, final_col + 2, site_name))

    # Fill matched (Dispatch) values by ItemName lookup, and write variance formulas
    disp_col_by_site = {name: dc for _, name, dc in matched_sites}
    for actual_col, matched_col, variance_col, site_name in site_groups:
        disp_col = disp_col_by_site[site_name]
        for r in range(2, max_row + 1):
            item_name = out_ws.cell(r, proj_name_col).value
            drow = disp_item_rows.get(_norm(item_name))
            out_ws.cell(r, matched_col).value = disp_ws.cell(drow, disp_col).value if drow else None
        actual_letter = get_column_letter(actual_col)
        matched_letter = get_column_letter(matched_col)
        for r in range(2, max_row + 1):
            out_ws.cell(r, variance_col).value = f"={matched_letter}{r}-{actual_letter}{r}"

        # Merge the 3-column header into one cell named just the site name
        out_ws.merge_cells(start_row=1, start_column=actual_col, end_row=1, end_column=variance_col)
        top_left = out_ws.cell(1, actual_col)
        top_left.value = site_name
        top_left.alignment = Alignment(horizontal="center", vertical="center")
        top_left.font = Font(bold=True)

    # Insert sub-header row (Projection / Dispatch / Variance) below row 1
    out_ws.insert_rows(2)
    new_max_row = out_ws.max_row
    for actual_col, matched_col, variance_col, site_name in site_groups:
        out_ws.cell(2, actual_col).value = "Projection"
        out_ws.cell(2, matched_col).value = "Dispatch"
        out_ws.cell(2, variance_col).value = "Variance"
        for c in (actual_col, matched_col, variance_col):
            out_ws.cell(2, c).alignment = Alignment(horizontal="center")
            out_ws.cell(2, c).font = Font(bold=True)
        # Rewrite variance formulas for the shifted data rows (now start at row 3)
        actual_letter = get_column_letter(actual_col)
        matched_letter = get_column_letter(matched_col)
        for r in range(3, new_max_row + 1):
            out_ws.cell(r, variance_col).value = f"={matched_letter}{r}-{actual_letter}{r}"

    # Totals at the very end
    out_max_col = out_ws.max_column
    total_proj_col = out_max_col + 1
    total_disp_col = out_max_col + 2
    total_var_col = out_max_col + 3
    proj_cols = [g[0] for g in site_groups]
    disp_cols = [g[1] for g in site_groups]
    var_cols = [g[2] for g in site_groups]

    headers = {
        total_proj_col: "Total Projection",
        total_disp_col: "Total Dispatch",
        total_var_col: "Total Variance",
    }
    for col, label in headers.items():
        out_ws.cell(1, col).value = label
        out_ws.merge_cells(start_row=1, start_column=col, end_row=2, end_column=col)
        cell = out_ws.cell(1, col)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.font = Font(bold=True)

    def _formula(cols, row):
        return "=" + "+".join(f"{get_column_letter(c)}{row}" for c in cols)

    for r in range(3, new_max_row + 1):
        out_ws.cell(r, total_proj_col).value = _formula(proj_cols, r)
        out_ws.cell(r, total_disp_col).value = _formula(disp_cols, r)
        out_ws.cell(r, total_var_col).value = _formula(var_cols, r)

    buf = io.BytesIO()
    out_wb.save(buf)
    buf.seek(0)

    report = {
        "matched_sites": [name for _, name, _ in matched_sites],
        "unmatched_sites": unmatched_sites,
        "total_sites_in_projection_file": len(proj_headers),
        "total_sites_in_dispatch_file": len(disp_headers),
    }
    return buf.read(), report
