import io
import pandas as pd
import streamlit as st
import shopee_core as core

st.set_page_config(page_title="Shopee Bulk Stock Update", page_icon="📦", layout="wide")
st.title("📦 Shopee Bulk Stock Update")
st.caption("Sync stock from the MM Masterlist to Shopee. The app returns an updated **Shopee Match Review** sheet "
           "plus two ready-to-upload Shopee files. Only the **Seller Stock** column is changed.")

with st.expander("ℹ️ How stock is decided (rules)"):
    st.markdown("""
| Registry tab | Seller Stock written |
|---|---|
| **Locked Matches** | Sum of Available Qty across all linked Masterlist SKUs |
| **Skipped (SKU=NO / 0 Qty)** | `0` |
| **Accessories (Not Matched)** | `10` (preset) |
| **Match Review / not in registry** | left unchanged |
Matched by **Variation ID**. Precedence: Locked → Skipped → Accessories → unchanged.
""")

st.subheader("1 · Upload files")
c1, c2 = st.columns(2)
with c1:
    ml_file  = st.file_uploader("Masterlist Excel", type=["xlsx"], key="ml")
    reg_file = st.file_uploader("SKU Registry Excel (Shopee Match Review sheet)", type=["xlsx"], key="reg")
with c2:
    sh1 = st.file_uploader("Shopee bulk stock — File 1 of 2", type=["xlsx"], key="s1")
    sh2 = st.file_uploader("Shopee bulk stock — File 2 of 2", type=["xlsx"], key="s2")

run = st.button("🚀 Process", type="primary", disabled=not (ml_file and reg_file and (sh1 or sh2)))

if run:
    try:
        ml_full = core.load_masterlist_full(ml_file)
        avail = {m["id"]: m["avail"] for m in ml_full}
        reg_bytes = reg_file.read()
        reg = core.load_registry(io.BytesIO(reg_bytes))
    except Exception as e:
        st.error(f"Could not read Masterlist/registry: {e}"); st.stop()

    new_ml = core.new_masterlist_skus(ml_full, reg)
    updated_registry = core.build_updated_registry(reg_bytes, avail, new_ml)

    outputs, all_stats = [], []
    for f, label, outname in [(sh1,"File 1 of 2","Shopee_Stock_Updated_1of2.xlsx"),
                              (sh2,"File 2 of 2","Shopee_Stock_Updated_2of2.xlsx")]:
        if not f: continue
        try:
            outb, stats, _ = core.process_shopee(f.read(), avail, reg, file_label=label)
            outputs.append((outname, outb)); all_stats.append((label, stats))
        except Exception as e:
            st.error(f"{label}: {e}")

    st.session_state["results"] = {
        "updated_registry": updated_registry, "new_ml": new_ml,
        "outputs": outputs, "stats": all_stats,
        "reg_counts": {k: (len(v) if hasattr(v,'__len__') else v) for k,v in
                       {"Locked":reg["locked"],"Skipped":reg["skipped"],"Accessories":reg["accessories"]}.items()},
    }
    st.session_state.pop("reviewed", None)

res = st.session_state.get("results")
if res:
    st.divider()
    st.subheader("2 · Updated Shopee Match Review sheet")
    st.download_button("⬇️ Download updated Shopee Match Review sheet", res["updated_registry"],
                       file_name="Shopee_Match_Review_UPDATED.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.subheader("3 · Validation summary")
    df = pd.DataFrame([{"File": l, "Rows": s["total_rows"], "Locked→sum": s["locked"],
                        "Skipped→0": s["skipped"], "Accessories→10": s["accessories"],
                        "Unchanged": s["unchanged"]} for l, s in res["stats"]])
    if not df.empty:
        tot = {"File":"TOTAL", **{k:int(df[k].sum()) for k in df.columns if k!="File"}}
        st.dataframe(pd.concat([df, pd.DataFrame([tot])], ignore_index=True),
                     use_container_width=True, hide_index=True)

    st.subheader("4 · Review: New Masterlist SKUs")
    new_ml = res["new_ml"]
    if new_ml:
        st.warning(f"⚠️ {len(new_ml)} new Masterlist SKU(s) are not yet linked. "
                   f"Review them (see the 'New Masterlist SKUs' tab in the updated sheet above), then confirm below to unlock the Shopee files.")
        st.dataframe(pd.DataFrame([{"Stock Type ID":m["id"],"Category":m["category"],"Brand":m["brand"],
                                    "Model":m["model"],"Color":m["color"],"Available Qty":m["avail"]} for m in new_ml]),
                     use_container_width=True, hide_index=True)
        reviewed = st.checkbox("✅ I have reviewed the New Masterlist SKUs tab", key="reviewed")
    else:
        st.success("No new Masterlist SKUs to review — all Masterlist SKUs are accounted for.")
        reviewed = True

    st.subheader("5 · Download Shopee bulk-upload files")
    if reviewed:
        for outname, outb in res["outputs"]:
            st.download_button(f"⬇️ {outname}", outb, file_name=outname,
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=outname)
        st.caption("Upload these in Shopee Seller Centre → Mass Update → Stock.")
    else:
        st.info("🔒 Tick the review confirmation above to unlock the two Shopee files.")
