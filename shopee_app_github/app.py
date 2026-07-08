import io
import pandas as pd
import streamlit as st
import shopee_core as core

st.set_page_config(page_title="Shopee Bulk Stock Update", page_icon="📦", layout="wide")
st.title("📦 Shopee Bulk Stock Update")
st.caption(f"build v6 · core {core.CORE_VERSION}")
st.caption("Reconciles your SKU registry with the Masterlist + Shopee export. Returns an updated "
           "**Shopee Match Review** sheet and two ready-to-upload Shopee files. Only **Seller Stock** changes.")

with st.expander("ℹ️ How it works"):
    st.markdown("""
**Stock rules (by Variation ID):** Locked → sum of linked Masterlist Available Qty · Skipped (SKU=NO/0Qty) → `0`
· Accessories → `10` · everything else → unchanged.

**Reconciliation each run:**
- Decisions you filled in the registry are **applied**: *New Masterlist SKUs* → Linked move to **Locked Matches**
  (stock combines), *Not applicable* → **Not Selling**, *Not on Shopee yet* → **Not on Shopee Yet**;
  *Match Review* rows you link (fill Corrected Masterlist ID + Decision) also move to **Locked Matches**.
- **New Shopee variations** not yet in the registry appear in the **Match Review** tab.
- **New Masterlist SKUs** (not yet linked) appear in the **New Masterlist SKUs** tab for review.

**Loop:** run → download updated sheet → fill decisions → re-upload it as the registry → run again to lock them in.
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
        reg_bytes = reg_file.read()
        locked, not_selling, not_on_shopee, summ = core.read_registry(reg_bytes)
    except Exception as e:
        st.error(f"Could not read Masterlist/registry: {e}"); st.stop()

    shopee_inputs = [(sh1,"File 1 of 2","Shopee_Stock_Updated_1of2.xlsx"),
                     (sh2,"File 2 of 2","Shopee_Stock_Updated_2of2.xlsx")]
    raw = [(f.read(), label, outname) for f, label, outname in shopee_inputs if f]
    svars = []
    for data, label, _ in raw:
        try: svars += core.read_shopee_variations(data)
        except Exception as e: st.error(f"{label}: {e}")

    cls, new_ml, avail, mlmap = core.reconcile(ml_full, locked, not_selling, not_on_shopee, svars)
    updated_registry = core.build_registry_workbook(cls, new_ml, mlmap, not_selling, not_on_shopee, summ)

    outputs, stats = [], []
    for data, label, outname in raw:
        try:
            ob, st_ = core.process_shopee(data, avail, locked, file_label=label)
            outputs.append((outname, ob)); stats.append((label, st_))
        except Exception as e:
            st.error(f"{label}: {e}")

    st.session_state["results"] = {"updated_registry":updated_registry,"new_ml":new_ml,
        "outputs":outputs,"stats":stats,"summary":summ,"match_review":len(cls["match_review"])}
    st.session_state.pop("reviewed", None)

res = st.session_state.get("results")
if res:
    su = res["summary"]
    if any(su.values()):
        st.success(f"Applied from your review — Linked to Locked: {su['linked']+su['match_review_locked']}, "
                   f"Not Selling: {su['not_selling']}, Not on Shopee Yet: {su['not_on_shopee']}.")
    st.divider()
    st.subheader("2 · Updated Shopee Match Review sheet")
    st.download_button("⬇️ Download updated Shopee Match Review sheet", res["updated_registry"],
                       file_name="Shopee_Match_Review_UPDATED.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.subheader("3 · Validation summary")
    df = pd.DataFrame([{"File":l,"Rows":s["total_rows"],"Locked→sum":s["locked"],
                        "Skipped→0":s["skipped"],"Accessories→10":s["accessories"],"Unchanged":s["unchanged"]}
                       for l,s in res["stats"]])
    if not df.empty:
        tot={"File":"TOTAL", **{k:int(df[k].sum()) for k in df.columns if k!="File"}}
        st.dataframe(pd.concat([df,pd.DataFrame([tot])],ignore_index=True), use_container_width=True, hide_index=True)
    st.caption(f"🆕 {res['match_review']} Shopee variation(s) are in the Match Review tab awaiting a match "
               f"(new or not-yet-linked). Their stock is left unchanged until you link them.")

    st.subheader("4 · Review: New Masterlist SKUs")
    new_ml = res["new_ml"]
    if new_ml:
        st.warning(f"⚠️ {len(new_ml)} new Masterlist SKU(s) not yet linked. Review the 'New Masterlist SKUs' tab "
                   f"in the sheet above; link them and re-upload to lock them in. Confirm below to unlock the Shopee files.")
        st.dataframe(pd.DataFrame([{"Stock Type ID":m["id"],"Category":m["category"],"Brand":m["brand"],
                     "Model":m["model"],"Color":m["color"],"Available Qty":m["avail"]} for m in new_ml]),
                     use_container_width=True, hide_index=True)
        reviewed = st.checkbox("✅ I have reviewed the New Masterlist SKUs tab", key="reviewed")
    else:
        st.success("No new Masterlist SKUs to review.")
        reviewed = True

    st.subheader("5 · Download Shopee bulk-upload files")
    if reviewed:
        for outname, ob in res["outputs"]:
            st.download_button(f"⬇️ {outname}", ob, file_name=outname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=outname)
        st.caption("Upload these in Shopee Seller Centre → Mass Update → Stock.")
    else:
        st.info("🔒 Tick the review confirmation above to unlock the two Shopee files.")
