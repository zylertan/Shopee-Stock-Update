import io
import pandas as pd
import streamlit as st
import shopee_core as core

st.set_page_config(page_title="Mister Mobile · Shopee Stock Sync", page_icon="📱", layout="wide")

# ---------- Mister Mobile theme (Yellow #FFEB00 / Black / Cream) ----------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@600;700;800&family=Open+Sans:wght@400;600&display=swap');
html, body, [class*="css"], .stMarkdown, .stDataFrame { font-family: 'Open Sans', sans-serif; }
h1, h2, h3, h4 { font-family: 'Montserrat', sans-serif !important; font-weight: 800 !important; color:#111; letter-spacing:-.2px; }
#MainMenu, footer {visibility: hidden;}
.block-container { padding-top: 1.2rem; max-width: 1150px; }

/* Brand header banner */
.mm-header { background:#FFEB00; border-radius:16px; padding:22px 26px; display:flex; align-items:center;
  gap:18px; box-shadow:0 2px 0 #000; margin-bottom:6px; }
.mm-badge { width:52px; height:52px; border-radius:14px; background:#000; color:#FFEB00; font-family:'Montserrat';
  font-weight:800; font-size:22px; display:flex; align-items:center; justify-content:center; flex:0 0 auto; }
.mm-title { font-family:'Montserrat'; font-weight:800; font-size:26px; color:#000; line-height:1.1; }
.mm-sub   { font-family:'Open Sans'; font-size:13px; color:#333; margin-top:2px; }
.mm-tag   { margin-left:auto; background:#000; color:#fff; font-size:11px; font-weight:600; padding:5px 10px; border-radius:20px; }

/* Step chips */
.mm-step { display:flex; align-items:center; gap:10px; margin:22px 0 10px; }
.mm-step .num { width:28px; height:28px; border-radius:50%; background:#000; color:#FFEB00; font-weight:800;
  font-family:'Montserrat'; display:flex; align-items:center; justify-content:center; font-size:14px; }
.mm-step .lbl { font-family:'Montserrat'; font-weight:800; font-size:18px; color:#111; }

/* Cards / metrics */
div[data-testid="stMetric"] { background:#F9F4E1; border:1px solid #efe6c4; border-radius:14px; padding:14px 16px; }
div[data-testid="stMetricValue"] { font-family:'Montserrat'; font-weight:800; }
.stButton>button { font-family:'Montserrat'; font-weight:700; border-radius:10px; border:2px solid #000; }
.stButton>button[kind="primary"] { background:#FFEB00; color:#000; }
.stDownloadButton>button { border-radius:10px; font-weight:700; border:2px solid #000; background:#fff; color:#000; }
.mm-note { background:#F9F4E1; border-left:5px solid #FFEB00; border-radius:8px; padding:10px 14px; font-size:14px; }
.mm-foot { text-align:center; color:#6D6962; font-size:12px; margin-top:34px; }
hr { border-color:#eee; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="mm-header">
  <div class="mm-badge">MM</div>
  <div>
    <div class="mm-title">Mister Mobile · Shopee Stock Sync</div>
    <div class="mm-sub">Sync available stock from the MM Masterlist to Shopee — only the Seller Stock column changes.</div>
  </div>
  <div class="mm-tag">build v7 · core %s</div>
</div>
""" % getattr(core, "CORE_VERSION", "?"), unsafe_allow_html=True)

def step(n, label):
    st.markdown(f'<div class="mm-step"><div class="num">{n}</div><div class="lbl">{label}</div></div>', unsafe_allow_html=True)

with st.expander("ℹ️  How it works & stock rules"):
    st.markdown("""
**Stock rules (matched by Variation ID)** — Locked → sum of linked Masterlist Available Qty · Skipped (SKU=NO/0Qty) → `0`
· Accessories → `10` · everything else → unchanged.

**Each run reconciles your registry:** decisions you filled in are applied (New Masterlist SKUs → *Linked* to **Locked Matches**,
*Not Selling* / *Not on Shopee Yet* to their tabs; *Match Review* links → **Locked Matches**). New Shopee variations appear in **Match Review**;
new Masterlist SKUs appear in **New Masterlist SKUs**.

**Loop:** run → download updated sheet → fill decisions → re-upload it as the registry → run again.
""")

# ---------- Step 1: uploads ----------
step(1, "Upload your files")
c1, c2 = st.columns(2)
with c1:
    ml_file  = st.file_uploader("📗  Masterlist Excel", type=["xlsx"], key="ml")
    reg_file = st.file_uploader("📘  SKU Registry (Shopee Match Review sheet)", type=["xlsx"], key="reg")
with c2:
    sh1 = st.file_uploader("📦  Shopee bulk stock — File 1 of 2", type=["xlsx"], key="s1")
    sh2 = st.file_uploader("📦  Shopee bulk stock — File 2 of 2", type=["xlsx"], key="s2")

# Decisions-detected readout (prevents uploading the wrong lookalike file)
reg_bytes = None
if reg_file is not None:
    try:
        reg_bytes = reg_file.getvalue()
        _l, _ns, _nos, _su = core.read_registry(reg_bytes)
        found = _su["linked"] + _su["not_selling"] + _su["not_on_shopee"] + _su["match_review_locked"]
        if found:
            st.markdown(f'<div class="mm-note">📋 <b>Decisions detected in this registry:</b> '
                        f'{_su["linked"]} linked · {_su["not_selling"]} not-selling · {_su["not_on_shopee"]} not-on-Shopee '
                        f'· {_su["match_review_locked"]} match-review links — these will be applied on Process.</div>',
                        unsafe_allow_html=True)
        else:
            st.markdown('<div class="mm-note">📋 No new decisions detected in this registry (that\'s fine if you have '
                        'nothing new to link — otherwise check you uploaded the file you filled in).</div>', unsafe_allow_html=True)
    except Exception as e:
        st.warning(f"Couldn't pre-read the registry: {e}")

ready = ml_file and reg_file and (sh1 or sh2)
st.write("")
run = st.button("🚀  Process files", type="primary", disabled=not ready, use_container_width=False)

if run:
    try:
        ml_full = core.load_masterlist_full(ml_file)
        if reg_bytes is None: reg_bytes = reg_file.getvalue()
        locked, not_selling, not_on_shopee, summ = core.read_registry(reg_bytes)
    except Exception as e:
        st.error(f"Could not read Masterlist/registry: {e}"); st.stop()
    raw = [(f.getvalue(), lbl, out) for f, lbl, out in
           [(sh1,"File 1 of 2","Shopee_Stock_Updated_1of2.xlsx"),(sh2,"File 2 of 2","Shopee_Stock_Updated_2of2.xlsx")] if f]
    svars = []
    for data, lbl, _ in raw:
        try: svars += core.read_shopee_variations(data)
        except Exception as e: st.error(f"{lbl}: {e}")
    cls, new_ml, avail, mlmap = core.reconcile(ml_full, locked, not_selling, not_on_shopee, svars)
    updated = core.build_registry_workbook(cls, new_ml, mlmap, not_selling, not_on_shopee, summ)
    outputs, stats = [], []
    for data, lbl, out in raw:
        try:
            ob, s = core.process_shopee(data, avail, locked, file_label=lbl)
            outputs.append((out, ob)); stats.append((lbl, s))
        except Exception as e: st.error(f"{lbl}: {e}")
    st.session_state["res"] = {"updated":updated,"new_ml":new_ml,"outputs":outputs,"stats":stats,
                               "summary":summ,"match_review":len(cls["match_review"])}
    st.session_state.pop("reviewed", None)

res = st.session_state.get("res")
if res:
    su = res["summary"]
    if any(su.values()):
        st.success(f"✅  Applied from your review — Linked to Locked: {su['linked']+su['match_review_locked']} · "
                   f"Not Selling: {su['not_selling']} · Not on Shopee Yet: {su['not_on_shopee']}")

    step(2, "Stock summary")
    tot = {k:0 for k in ("total_rows","locked","skipped","accessories","unchanged")}
    for _, s in res["stats"]:
        for k in tot: tot[k]+=s[k]
    m = st.columns(5)
    m[0].metric("Variations", f"{tot['total_rows']:,}")
    m[1].metric("🔒 Locked → synced", f"{tot['locked']:,}")
    m[2].metric("Skipped → 0", f"{tot['skipped']:,}")
    m[3].metric("Accessories → 10", f"{tot['accessories']:,}")
    m[4].metric("Unchanged", f"{tot['unchanged']:,}")
    df = pd.DataFrame([{"File":l,"Rows":s["total_rows"],"Locked":s["locked"],"Skipped":s["skipped"],
                        "Accessories":s["accessories"],"Unchanged":s["unchanged"]} for l,s in res["stats"]])
    if not df.empty: st.dataframe(df, use_container_width=True, hide_index=True)
    st.caption(f"🔎 {res['match_review']:,} Shopee variation(s) await a match in the Match Review tab (new or not-yet-linked) — their stock is left unchanged.")

    step(3, "Updated Shopee Match Review sheet")
    st.download_button("⬇️  Download updated Match Review sheet", res["updated"],
        file_name="Shopee_Match_Review_UPDATED.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.caption("This becomes your registry for next time. Fill any decisions, then re-upload it here.")

    step(4, "Review new Masterlist SKUs")
    new_ml = res["new_ml"]
    if new_ml:
        st.markdown(f'<div class="mm-note">⚠️ <b>{len(new_ml)} new Masterlist SKU(s)</b> not yet linked. Review the '
                    f'“New Masterlist SKUs” tab in the sheet above, link them, and re-upload. Confirm below to unlock the Shopee files.</div>',
                    unsafe_allow_html=True)
        st.dataframe(pd.DataFrame([{"Stock Type ID":m["id"],"Category":m["category"],"Brand":m["brand"],
                     "Model":m["model"],"Color":m["color"],"Available Qty":m["avail"]} for m in new_ml]),
                     use_container_width=True, hide_index=True)
        reviewed = st.checkbox("✅  I have reviewed the New Masterlist SKUs tab", key="reviewed")
    else:
        st.success("👍  No new Masterlist SKUs to review — all accounted for.")
        reviewed = True

    step(5, "Download Shopee bulk-upload files")
    if reviewed:
        d = st.columns(len(res["outputs"]) or 1)
        for i,(out, ob) in enumerate(res["outputs"]):
            d[i].download_button(f"⬇️  {out}", ob, file_name=out,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=out, use_container_width=True)
        st.caption("Upload these in Shopee Seller Centre → Mass Update → Stock.")
    else:
        st.markdown('<div class="mm-note">🔒 Tick the review confirmation above to unlock the two Shopee files.</div>', unsafe_allow_html=True)

st.markdown('<div class="mm-foot">Mister Mobile · Shopee Stock Sync — internal tool</div>', unsafe_allow_html=True)
