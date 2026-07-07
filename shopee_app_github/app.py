import io, datetime
import pandas as pd
import streamlit as st
import shopee_core as core

st.set_page_config(page_title="Shopee Bulk Stock Update", page_icon="📦", layout="wide")
st.title("📦 Shopee Bulk Stock Update")
st.caption("Sync available stock from the MM Masterlist to Shopee, using your confirmed SKU registry. "
           "Only the **Seller Stock** column is changed; the original template format is preserved.")

with st.expander("ℹ️ How stock is decided (rules)", expanded=False):
    st.markdown("""
| Registry tab | Seller Stock written |
|---|---|
| **Locked Matches** | Sum of Available Qty across all linked Masterlist SKUs |
| **Skipped (SKU=NO / 0 Qty)** | `0` |
| **Accessories (Not Matched)** | `10` (preset) |
| **Match Review / not in registry** | left unchanged |
Matching is by **Variation ID**. Precedence: Locked → Skipped → Accessories → unchanged.
""")

st.subheader("1 · Upload files")
c1, c2 = st.columns(2)
with c1:
    ml_file  = st.file_uploader("Masterlist Excel", type=["xlsx"], key="ml")
    reg_file = st.file_uploader("SKU Registry Excel (tabs: Locked Matches / Skipped / Accessories …)", type=["xlsx"], key="reg")
with c2:
    sh1 = st.file_uploader("Shopee bulk stock — File 1 of 2", type=["xlsx"], key="s1")
    sh2 = st.file_uploader("Shopee bulk stock — File 2 of 2", type=["xlsx"], key="s2")

run = st.button("🚀 Generate updated Shopee files", type="primary",
                disabled=not (ml_file and reg_file and (sh1 or sh2)))

if run:
    try:
        with st.spinner("Reading Masterlist & registry…"):
            avail = core.load_masterlist(ml_file)
            reg = core.load_registry(reg_file)
    except Exception as e:
        st.error(f"Could not read Masterlist/registry: {e}"); st.stop()

    st.success(f"Masterlist: {len(avail):,} SKUs · Registry — Locked: {len(reg['locked']):,}, "
               f"Skipped: {len(reg['skipped']):,}, Accessories: {len(reg['accessories']):,}")

    outputs, all_stats, all_errors, seen_all = [], [], [], set()
    files = [(sh1, "File 1 of 2", "Shopee_Stock_Updated_1of2.xlsx"),
             (sh2, "File 2 of 2", "Shopee_Stock_Updated_2of2.xlsx")]
    for f, label, outname in files:
        if not f: continue
        try:
            data = f.read()
            outb, stats, errs, seen = core.process_shopee(data, avail, reg, file_label=label)
            outputs.append((outname, outb)); all_stats.append((label, stats))
            all_errors += errs; seen_all |= seen
        except Exception as e:
            st.error(f"{label}: {e}")

    if all_stats:
        st.subheader("2 · Validation summary")
        df = pd.DataFrame([{"File": l, "Rows": s["total_rows"], "Locked→sum": s["locked"],
                            "Skipped→0": s["skipped"], "Accessories→10": s["accessories"],
                            "Unchanged": s["unchanged"]} for l, s in all_stats])
        tot = {"File": "TOTAL", **{k: int(df[k].sum()) for k in df.columns if k != "File"}}
        df = pd.concat([df, pd.DataFrame([tot])], ignore_index=True)
        st.dataframe(df, use_container_width=True, hide_index=True)
        m1,m2,m3,m4 = st.columns(4)
        m1.metric("Locked updated", int(df.iloc[-1]["Locked→sum"]))
        m2.metric("Skipped → 0", int(df.iloc[-1]["Skipped→0"]))
        m3.metric("Accessories → 10", int(df.iloc[-1]["Accessories→10"]))
        m4.metric("Left unchanged", int(df.iloc[-1]["Unchanged"]))

        st.subheader("3 · Error report")
        missing_cov = core.registry_coverage(reg, seen_all)
        if all_errors:
            edf = pd.DataFrame(all_errors)
            st.warning(f"{len(all_errors)} locked variations reference Masterlist ID(s) not found in the Masterlist "
                       f"(their missing SKUs counted as 0).")
            st.dataframe(edf, use_container_width=True, hide_index=True)
        else:
            st.success("No Masterlist-ID resolution errors.")
        if missing_cov:
            st.info(f"{len(missing_cov)} registry Variation IDs were not found in the uploaded Shopee files "
                    f"(nothing to update for them).")
            st.dataframe(pd.DataFrame(missing_cov), use_container_width=True, hide_index=True)

        # combined error report download
        if all_errors or missing_cov:
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as xw:
                (pd.DataFrame(all_errors) if all_errors else pd.DataFrame(columns=["file","variation_id","issue","detail"])).to_excel(xw, sheet_name="Masterlist ID errors", index=False)
                (pd.DataFrame(missing_cov) if missing_cov else pd.DataFrame(columns=["category","variation_id"])).to_excel(xw, sheet_name="Registry not in Shopee", index=False)
            st.download_button("⬇️ Download error report (Excel)", buf.getvalue(),
                               file_name="Shopee_Stock_Update_Error_Report.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.subheader("4 · Download updated Shopee files")
        for outname, outb in outputs:
            st.download_button(f"⬇️ {outname}", outb, file_name=outname,
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               key=outname)
        st.caption("Upload these in Shopee Seller Centre → Mass Update → Stock.")
