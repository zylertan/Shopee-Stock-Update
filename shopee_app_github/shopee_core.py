"""Core logic for the Shopee Bulk Stock Update app (no Streamlit dependency)."""
import io, re, zipfile, copy
import openpyxl

# ---------- Shopee file repair (invalid activePane attr) ----------
def fix_shopee_bytes(data: bytes) -> io.BytesIO:
    zin = zipfile.ZipFile(io.BytesIO(data), "r")
    out = io.BytesIO(); zout = zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        raw = zin.read(item.filename)
        if item.filename.endswith(".xml"):
            raw = re.sub(r'activePane="[^"]*"', 'activePane="topLeft"', raw.decode("utf-8", "ignore")).encode("utf-8")
        zout.writestr(item, raw)
    zin.close(); zout.close(); out.seek(0); return out

def _norm_id(v):
    if v is None: return ""
    s = str(v).strip()
    if s.endswith(".0") and s[:-2].isdigit(): s = s[:-2]
    return s

def _split_ids(cell):
    return [x for x in re.split(r"[;,/]", str(cell))] if cell is not None else []

def _is_freebie(model): return "FREEB" in (model or "").upper()

# ---------- Masterlist ----------
def load_masterlist_full(file):
    wb = openpyxl.load_workbook(file, data_only=True); ws = wb.active
    id_col = qty_col = header_row = cat_col = brand_col = model_col = color_col = None
    for r in range(1, min(6, ws.max_row) + 1):
        vals = [str(c.value).strip().lower() if c.value is not None else "" for c in ws[r]]
        if "stock type id" in vals:
            header_row = r; id_col = vals.index("stock type id") + 1
            for name, key in [("category","cat"),("brand","brand"),("model","model"),("color","color"),("total","qty")]:
                if name in vals:
                    idx = vals.index(name) + 1
                    if key=="cat": cat_col=idx
                    elif key=="brand": brand_col=idx
                    elif key=="model": model_col=idx
                    elif key=="color": color_col=idx
                    elif key=="qty": qty_col=idx
            break
    if id_col is None: raise ValueError("Could not find 'Stock Type ID' column in Masterlist.")
    if qty_col is None: raise ValueError("Could not find 'Total' (Available Quantity) column in Masterlist.")
    rows = []
    for r in range(header_row + 1, ws.max_row + 1):
        sid = _norm_id(ws.cell(r, id_col).value)
        if not sid: continue
        q = ws.cell(r, qty_col).value
        try: q = int(float(q))
        except (TypeError, ValueError): q = 0
        rows.append({"id":sid,
            "category":(ws.cell(r,cat_col).value if cat_col else "") or "",
            "brand":(ws.cell(r,brand_col).value if brand_col else "") or "",
            "model":(ws.cell(r,model_col).value if model_col else "") or "",
            "color":(ws.cell(r,color_col).value if color_col else "") or "",
            "avail":q})
    return rows

def load_masterlist(file):
    return {m["id"]: m["avail"] for m in load_masterlist_full(file)}

# ---------- Registry ----------
def _find_col(ws, header_row, *keywords):
    vals = [str(c.value).strip().lower() if c.value is not None else "" for c in ws[header_row]]
    for kw in keywords:
        for i, v in enumerate(vals):
            if kw in v: return i + 1
    return None

def load_registry(file):
    wb = openpyxl.load_workbook(file, data_only=True)
    reg = {"locked": {}, "skipped": set(), "accessories": set(),
           "not_selling": set(), "not_on_shopee": set(), "locked_mids": set()}
    names = {n.lower(): n for n in wb.sheetnames}
    def sheet_like(*subs):
        for low, real in names.items():
            if all(s in low for s in subs): return wb[real]
        return None
    ws = sheet_like("locked")
    if ws is not None:
        vcol = _find_col(ws, 1, "variation id"); mcol = _find_col(ws, 1, "masterlist id")
        if vcol and mcol:
            for r in range(2, ws.max_row + 1):
                vid = _norm_id(ws.cell(r, vcol).value)
                if not vid: continue
                mids = [_norm_id(x) for x in _split_ids(ws.cell(r, mcol).value) if _norm_id(x)]
                reg["locked"].setdefault(vid, [])
                for m in mids:
                    if m not in reg["locked"][vid]: reg["locked"][vid].append(m)
                    reg["locked_mids"].add(m)
    for key, subs in [("skipped",("skipped",)),("accessories",("accessor",))]:
        ws = sheet_like(*subs)
        if ws is not None:
            vcol = _find_col(ws, 1, "variation id")
            if vcol:
                for r in range(2, ws.max_row + 1):
                    vid = _norm_id(ws.cell(r, vcol).value)
                    if vid: reg[key].add(vid)
    for key, subs in [("not_selling",("not selling",)),("not_on_shopee",("not on shopee",))]:
        ws = sheet_like(*subs)
        if ws is not None:
            scol = _find_col(ws, 1, "stock type id", "masterlist id")
            if scol:
                for r in range(2, ws.max_row + 1):
                    sid = _norm_id(ws.cell(r, scol).value)
                    if sid: reg[key].add(sid)
    return reg

def new_masterlist_skus(ml_full, reg):
    covered = set(reg["locked_mids"]) | reg["not_selling"] | reg["not_on_shopee"]
    return [m for m in ml_full if (m["id"] not in covered) and not _is_freebie(m["model"])]

# ---------- Build the UPDATED Match Review workbook ----------
def build_updated_registry(reg_bytes, avail_map, new_ml_rows):
    wb = openpyxl.load_workbook(io.BytesIO(reg_bytes))  # keep formatting
    names = {n.lower(): n for n in wb.sheetnames}
    def sheet_like(*subs):
        for low, real in names.items():
            if all(s in low for s in subs): return wb[real]
        return None
    # refresh Locked Matches Target Stock + ML Available Qty from current Masterlist
    lk = sheet_like("locked")
    if lk is not None:
        mcol = _find_col(lk, 1, "masterlist id"); acol = _find_col(lk, 1, "available qty"); tcol = _find_col(lk, 1, "target stock")
        if mcol and tcol:
            for r in range(2, lk.max_row + 1):
                mids = [_norm_id(x) for x in _split_ids(lk.cell(r, mcol).value) if _norm_id(x)]
                if not mids: continue
                avs = [avail_map.get(m) for m in mids]
                if acol: lk.cell(r, acol).value = "; ".join(str(a) if a is not None else "?" for a in avs)
                lk.cell(r, tcol).value = sum(a for a in avs if a is not None)
    # rewrite New Masterlist SKUs sheet
    nm = sheet_like("new masterlist")
    if nm is not None:
        if nm.max_row >= 2: nm.delete_rows(2, nm.max_row - 1)
        for i, m in enumerate(new_ml_rows, start=1):
            nm.append([i, m["id"], m["category"], m["brand"], m["model"], m["color"], m["avail"], "", "", "", ""])
    out = io.BytesIO(); wb.save(out); out.seek(0); return out.getvalue()

# ---------- Shopee processing ----------
ACC_PRESET = 10
def process_shopee(data, avail_map, reg, file_label="file"):
    wb = openpyxl.load_workbook(fix_shopee_bytes(data)); ws = wb.active
    var_col = stock_col = hdr_row = None
    for r in range(1, min(4, ws.max_row) + 1):
        vals = [str(c.value).strip().lower() if c.value is not None else "" for c in ws[r]]
        for i, v in enumerate(vals):
            if v == "et_title_variation_id" or v == "variation id": var_col = i + 1
            if v.startswith("et_title_variation_sto") or v == "seller stock": stock_col = i + 1
        if var_col and stock_col: hdr_row = r; break
    if not (var_col and stock_col):
        raise ValueError(f"Could not locate Variation ID / Seller Stock columns in Shopee {file_label}.")
    stats = {"total_rows":0,"locked":0,"skipped":0,"accessories":0,"unchanged":0}
    seen = set()
    for r in range(hdr_row + 1, ws.max_row + 1):
        vid = _norm_id(ws.cell(r, var_col).value)
        if not vid or not vid.isdigit(): continue
        stats["total_rows"] += 1; seen.add(vid)
        cell = ws.cell(r, stock_col)
        if vid in reg["locked"]:
            cell.value = sum(avail_map.get(m, 0) for m in reg["locked"][vid]); stats["locked"] += 1
        elif vid in reg["skipped"]:
            cell.value = 0; stats["skipped"] += 1
        elif vid in reg["accessories"]:
            cell.value = ACC_PRESET; stats["accessories"] += 1
        else:
            stats["unchanged"] += 1
    out = io.BytesIO(); wb.save(out); out.seek(0)
    return out.getvalue(), stats, seen
