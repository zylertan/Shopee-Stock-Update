"""Core logic for the Shopee Bulk Stock Update app (no Streamlit dependency)."""
import io, re, zipfile
import openpyxl

# ---------- Shopee file repair (Shopee exports have an invalid activePane attr) ----------
def fix_shopee_bytes(data: bytes) -> io.BytesIO:
    """Return a BytesIO of the workbook with invalid pane attributes fixed so openpyxl can load it."""
    zin = zipfile.ZipFile(io.BytesIO(data), "r")
    out = io.BytesIO()
    zout = zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        raw = zin.read(item.filename)
        if item.filename.endswith(".xml"):
            txt = raw.decode("utf-8", errors="ignore")
            txt = re.sub(r'activePane="[^"]*"', 'activePane="topLeft"', txt)
            raw = txt.encode("utf-8")
        zout.writestr(item, raw)
    zin.close(); zout.close()
    out.seek(0)
    return out

def _norm_id(v):
    if v is None: return ""
    s = str(v).strip()
    if s.endswith(".0") and s[:-2].isdigit(): s = s[:-2]
    return s

def _split_ids(cell):
    if cell is None: return []
    return [x for x in re.split(r"[;,/]", str(cell)) if x.strip()]

# ---------- Masterlist ----------
def load_masterlist(file) -> dict:
    """Return {stock_type_id: available_qty_int}. Available Qty = the 'Total' column."""
    wb = openpyxl.load_workbook(file, data_only=True)
    ws = wb.active
    # find header row & columns
    id_col = qty_col = header_row = None
    for r in range(1, min(6, ws.max_row) + 1):
        vals = [str(c.value).strip().lower() if c.value is not None else "" for c in ws[r]]
        if "stock type id" in vals:
            header_row = r
            id_col = vals.index("stock type id") + 1
            if "total" in vals: qty_col = vals.index("total") + 1
            break
    if id_col is None:
        raise ValueError("Could not find 'Stock Type ID' column in Masterlist.")
    if qty_col is None:
        raise ValueError("Could not find 'Total' (Available Quantity) column in Masterlist.")
    out = {}
    for r in range(header_row + 1, ws.max_row + 1):
        sid = _norm_id(ws.cell(r, id_col).value)
        if not sid: continue
        q = ws.cell(r, qty_col).value
        try: q = int(float(q))
        except (TypeError, ValueError): q = 0
        out[sid] = q
    return out

# ---------- SKU Registry ----------
def _find_col(ws, header_row, *keywords):
    vals = [str(c.value).strip().lower() if c.value is not None else "" for c in ws[header_row]]
    for kw in keywords:
        for i, v in enumerate(vals):
            if kw in v: return i + 1
    return None

def load_registry(file) -> dict:
    """Return {'locked': {vid:[mid,...]}, 'skipped': set(vid), 'accessories': set(vid)}."""
    wb = openpyxl.load_workbook(file, data_only=True)
    reg = {"locked": {}, "skipped": set(), "accessories": set()}
    names = {n.lower(): n for n in wb.sheetnames}
    def sheet_like(*subs):
        for low, real in names.items():
            if all(s in low for s in subs): return wb[real]
        return None
    # Locked Matches
    ws = sheet_like("locked")
    if ws is not None:
        vcol = _find_col(ws, 1, "variation id")
        mcol = _find_col(ws, 1, "masterlist id")
        if vcol and mcol:
            for r in range(2, ws.max_row + 1):
                vid = _norm_id(ws.cell(r, vcol).value)
                if not vid: continue
                mids = [_norm_id(x) for x in _split_ids(ws.cell(r, mcol).value)]
                reg["locked"].setdefault(vid, [])
                for m in mids:
                    if m and m not in reg["locked"][vid]: reg["locked"][vid].append(m)
    # Skipped
    ws = sheet_like("skipped")
    if ws is not None:
        vcol = _find_col(ws, 1, "variation id")
        if vcol:
            for r in range(2, ws.max_row + 1):
                vid = _norm_id(ws.cell(r, vcol).value)
                if vid: reg["skipped"].add(vid)
    # Accessories
    ws = sheet_like("accessor")
    if ws is not None:
        vcol = _find_col(ws, 1, "variation id")
        if vcol:
            for r in range(2, ws.max_row + 1):
                vid = _norm_id(ws.cell(r, vcol).value)
                if vid: reg["accessories"].add(vid)
    return reg

# ---------- Shopee processing ----------
ACC_PRESET = 10
def process_shopee(data: bytes, avail_map: dict, reg: dict, file_label="file"):
    """Update only Seller Stock by Variation ID. Returns (output_bytes, stats, errors)."""
    fixed = fix_shopee_bytes(data)
    wb = openpyxl.load_workbook(fixed)  # keep formatting
    ws = wb.active
    # locate machine-header row and columns
    var_col = stock_col = None
    hdr_row = None
    for r in range(1, min(4, ws.max_row) + 1):
        vals = [str(c.value).strip().lower() if c.value is not None else "" for c in ws[r]]
        for i, v in enumerate(vals):
            if v == "et_title_variation_id" or v == "variation id": var_col = i + 1
            if v.startswith("et_title_variation_sto") or v == "seller stock": stock_col = i + 1
        if var_col and stock_col: hdr_row = r; break
    if not (var_col and stock_col):
        raise ValueError(f"Could not locate Variation ID / Seller Stock columns in Shopee {file_label}.")
    stats = {"total_rows":0,"locked":0,"skipped":0,"accessories":0,"unchanged":0}
    errors = []
    seen_vids = set()
    for r in range(hdr_row + 1, ws.max_row + 1):
        vid = _norm_id(ws.cell(r, var_col).value)
        if not vid or not vid.isdigit():  # skip meta/blank rows
            continue
        stats["total_rows"] += 1
        seen_vids.add(vid)
        cell = ws.cell(r, stock_col)
        # precedence: Locked > Skipped > Accessories > unchanged
        if vid in reg["locked"]:
            mids = reg["locked"][vid]
            missing = [m for m in mids if m not in avail_map]
            total = sum(avail_map.get(m, 0) for m in mids)
            cell.value = total
            stats["locked"] += 1
            if missing:
                errors.append({"file":file_label,"variation_id":vid,"issue":"Masterlist ID(s) not in Masterlist (counted as 0)","detail":",".join(missing)})
        elif vid in reg["skipped"]:
            cell.value = 0; stats["skipped"] += 1
        elif vid in reg["accessories"]:
            cell.value = ACC_PRESET; stats["accessories"] += 1
        else:
            stats["unchanged"] += 1
    out = io.BytesIO(); wb.save(out); out.seek(0)
    return out.getvalue(), stats, errors, seen_vids

def registry_coverage(reg, seen_all):
    """Registry Variation IDs (locked/skipped/accessories) not present in the uploaded Shopee files."""
    missing = []
    for cat in ("locked","skipped","accessories"):
        ids = reg[cat].keys() if isinstance(reg[cat], dict) else reg[cat]
        for vid in ids:
            if vid not in seen_all: missing.append({"category":cat,"variation_id":vid})
    return missing
