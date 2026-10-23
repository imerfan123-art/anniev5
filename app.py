"""Annie v4 - Streamlit UI (belongs to annie_core.py).   Run:  streamlit run app.py"""

# =====================================================================================
# 1. IMPORTS AND SETUP
# =====================================================================================
import base64
from datetime import date
from pathlib import Path
from streamlit_pdf_viewer import pdf_viewer

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

import annie_core as annie

st.set_page_config(page_title="Annie v4", layout="wide")
con = annie.connect()
SAMPLES = Path(__file__).parent / "samples"
ss = st.session_state
ss.setdefault("results", [])
ss.setdefault("view_case", None)
ss.setdefault("view_po", None)
UPLOAD_TYPES = sorted(e.lstrip(".") for e in annie.SUPPORTED_EXTENSIONS)
ROUTE_ICON = {annie.R_M2: "✅", annie.R_REVIEW: "🕵️", annie.R_DUP_FILE: "⛔", annie.R_DUP_CONTENT: "⛔", "REJECTED": "🚫"}
M2_ICON = {"REGISTERED": "📄", "ALLOCATED": "🔗", "UNALLOCATED": "⏳", "CALLOFF_OPENED": "📞"}
DEFAULT_ORDER = "Default order"


# =====================================================================================
# 2. SHARED UI HELPERS (previews, summaries, filter/sort controls, tab bubbles)
# =====================================================================================
def file_view(doc, height=620, key=""):
    """Preview a stored file: PDF inline, image inline, anything else as download."""
    path = Path(doc["file_path"]) if doc and doc.get("file_path") else None
    if not path or not path.exists():
        st.info("No file stored for this document.")
        return
    data = path.read_bytes()
    ext = path.suffix.lower()
    if ext == ".pdf":
        pdf_viewer(data, height=height)
        #b64 = base64.b64encode(data).decode()
        #components.html(f'<iframe src="data:application/pdf;base64,{b64}" width="100%" height="{height - 10}" '
        #                f'style="border:1px solid #ccc"></iframe>', height=height)
    elif ext in annie.IMAGE_EXTENSIONS and ext not in (".heic", ".heif", ".tif", ".tiff"):
        st.image(data)
    else:
        st.caption(f"Preview not available for {ext} files.")
    st.download_button("Download original file", data, file_name=path.name, key=f"dl_{key}_{path.name}")


BASIS_LABEL = {"QTY": "Qty", "AMOUNT": "Amount", "SINGLE": "Single invoice (tick)"}


def summary(doc):
    p = doc["payload"]
    is_po = doc["doc_type"] == "PO"
    rows = {"Reference": p.get("po_reference") if is_po else p.get("invoice_reference"),
            "Vendor": p.get("vendor_name"), "Date": p.get("date"), "Total": annie.doc_total(p)}
    if is_po:
        rows.update({"Cost centre": p.get("cost_centre"), "Type": p.get("category"),
                     "Approver": f"{p.get('approver_name')} ({p.get('approver_title')})",
                     "Prepared by": f"{p.get('prepared_by')} ({p.get('prepared_by_title')})",
                     "Signed": "Yes" if p.get("signature_present") else "No"})
    else:
        rows["PO/PR cited on invoice"] = ", ".join(p.get("po_references") or [p.get("po_reference") or "-"])
    st.table(pd.DataFrame({"Field": list(rows), "Value": [str(v) for v in rows.values()]}).set_index("Field"))
    if p.get("items"):
        st.dataframe(pd.DataFrame(p["items"]), hide_index=True)
    elif p.get("service_description"):
        st.caption("Service: " + str(p["service_description"]))


def _sort_key(v):
    if v is None:
        return (1, 0, 0)
    if isinstance(v, (int, float)):
        return (0, 0, v)
    return (0, 1, str(v).lower())


def controls(key, rows, sort_fields):
    """Filter box + sort selector + order toggle + 'x of y entries' counter. Returns the filtered, sorted rows."""
    c1, c2, c3, c4 = st.columns([3, 2, 1.4, 2])
    q = c1.text_input("Filter", key=f"flt_{key}", placeholder="🔎 Filter entries...", label_visibility="collapsed")
    field = c2.selectbox("Sort by", [DEFAULT_ORDER] + list(sort_fields), key=f"srt_{key}",
                         label_visibility="collapsed", format_func=lambda f: f if f == DEFAULT_ORDER else f"Sort by: {f}")
    desc = c3.toggle("Descending", key=f"dsc_{key}")
    out = list(rows)
    if q.strip():
        out = [r for r in out if q.strip().lower() in " ".join(str(v) for v in r.values()).lower()]
    if field == DEFAULT_ORDER:
        out = out[::-1] if desc else out
    else:
        out = sorted(out, key=lambda r: _sort_key(r.get(field)), reverse=desc)
    c4.markdown(f"**{len(out)}** of {len(rows)} entries")
    return out


def bubble(n):
    """Small number bubble for tab names (circled digits; plain (n) above 20)."""
    if not n:
        return ""
    return " " + (chr(0x2460 + n - 1) if n <= 20 else f"({n})")


def render_results(results):
    rows = list(results)[::-1]
    rows = controls("results", rows, ["source_file", "m1_status", "route", "record_id"])
    for r in rows:
        m2 = r.get("m2") or {}
        label = (f"{ROUTE_ICON.get(r.get('route'), 'ℹ️')} {r.get('source_file') or r.get('record_id')} - "
                 f"{r.get('m1_status')} - {r.get('route')}")
        if m2:
            label += f" -> {M2_ICON.get(m2.get('outcome'), '')} {m2.get('outcome')} {m2.get('case_id') or ''} {m2.get('status') or ''}"
        with st.expander(label):
            st.json(r)


# =====================================================================================
# 3. CASE PAGE (one invoice: allocation, GR/SES batches, Post)
# =====================================================================================
def gr_form(case_id, lines):
    """One GR/SES batch. Quantities (goods) or amounts (services) are entered per allocated PO/PR line."""
    cat = "Services" if all(l["basis"] in ("AMOUNT", "SINGLE") for l in lines) else "Goods"
    preparers = annie.allowed_preparers(con, case_id)
    vendor = annie.case_po_docs(con, case_id)[0]["payload"].get("vendor_name")
    with st.form(f"gr_{case_id}"):
        st.caption("A GR/SES can be raised in batches. Enter what was actually received or performed in this batch; "
                   "the case completes once receipts cover everything billed on this invoice.")
        a, b = st.columns(2)
        a.text_input("PO/PR reference(s)", ", ".join(sorted({l["po_reference"] for l in lines})), disabled=True)
        b.text_input("Vendor", vendor or "", disabled=True)
        a.text_input("SES/GR number", annie.next_ses_number(con, cat) + " (auto-generated on submit)", disabled=True)
        prepared = b.selectbox("Prepared by (name) - must match the PO/PR", preparers)
        d1, d2 = st.columns(2)
        f_from = d1.date_input("Performance / service date - from", date.today())
        f_to = d2.date_input("Performance / service date - to", date.today())
        df = pd.DataFrame([dict(line_id=l["line_id"], po_reference=l["po_reference"], item=l["description"],
                                basis=BASIS_LABEL[l["basis"]], ordered=l["ordered"], received_for_this_invoice=l["received"],
                                billed_this_invoice=l["billed_this"], receive_now=min(l["needed"], l["room"])) for l in lines])
        edited = st.data_editor(df, hide_index=True, key=f"gr_ed_{case_id}",
                                disabled=[c for c in df.columns if c != "receive_now"],
                                column_config={"receive_now": st.column_config.NumberColumn("Receive now (single invoice: 1 = done)", min_value=0.0)})
        submitted = st.form_submit_button("Submit GR/SES", type="primary")
    if submitted:
        totals = {}
        for r in edited.to_dict("records"):
            totals[int(r["line_id"])] = totals.get(int(r["line_id"]), 0.0) + float(r["receive_now"] or 0)
        gr = dict(service_date_from=f_from.isoformat(), service_date_to=f_to.isoformat(), prepared_by=prepared,
                  lines=[dict(line_id=k, receive=v) for k, v in totals.items()])
        try:
            res = annie.submit_gr(con, case_id, gr)
            msg = f"GR/SES {res['ses_gr_number']} saved. Case status: {res['status']}."
            ss.flash = [msg]
            st.rerun()
        except ValueError as e:
            st.error(str(e))


def manual_panel(case_id):
    """Stage 1: match the invoice by hand (goods by qty, services by amount, or a single-invoice service)."""
    with st.expander("Match manually"):
        inv_lines, cands = annie.invoice_line_list(con, case_id), annie.manual_candidates(con, case_id)
        if not cands:
            st.info("No PO/PR line has anything left to bill.")
            return
        by_label = {c["label"]: c for c in cands}
        st.caption("For each invoice line choose the PO/PR line it bills. Enter a **quantity** for goods, a **money amount** for services "
                   "tracked by amount, or any number above 0 for a single-invoice service (it takes the whole invoice line). "
                   "To split one invoice line over several PO/PR lines, add rows with the same invoice-line number; each invoice line must reach 100%.")
        seed = []
        for ln in inv_lines:
            best = max(cands, key=lambda c: (annie.line_score(ln, c, False, False), c["vendor_sim"]))
            val = ln["qty"] if best["basis"] == "QTY" else (1.0 if best["basis"] == "SINGLE" else ln["amount"])
            seed.append(dict(inv_line=ln["index"], invoice_item=ln["description"], po_line=best["label"], allocate=min(val, best["open"]) if best["basis"] != "SINGLE" else 1.0))
        edited = st.data_editor(pd.DataFrame(seed), num_rows="dynamic", hide_index=True, key=f"mm_{case_id}", disabled=["invoice_item"],
                                column_config={"inv_line": st.column_config.NumberColumn("Invoice line", min_value=1, max_value=len(inv_lines), step=1),
                                               "invoice_item": "Invoice item",
                                               "po_line": st.column_config.SelectboxColumn("PO/PR line", options=list(by_label), width="large", required=True),
                                               "allocate": st.column_config.NumberColumn("Qty / amount", min_value=0.0)})
        rows = []
        for r in edited.to_dict("records"):
            if pd.isna(r.get("inv_line")) or r.get("po_line") not in by_label or pd.isna(r.get("allocate")):
                continue
            rows.append(dict(inv_line=int(r["inv_line"]), po_line_id=by_label[r["po_line"]]["line_id"], value=float(r["allocate"])))
        if st.button("Check match", key=f"mm_chk_btn_{case_id}"):
            ss[f"mm_chk_{case_id}"] = annie.manual_match(con, case_id, rows, dry_run=True)
        chk = ss.get(f"mm_chk_{case_id}")
        if chk:
            for e in chk["errors"]:
                st.error(e)
            if chk["rows"]:
                st.dataframe(pd.DataFrame(chk["rows"]).drop(columns=["po_line_id", "qty"]), hide_index=True)
            if chk["ok"] and chk["below_threshold"]:
                st.warning(f"Some choices score below the automatic threshold ({annie.LINK_THRESHOLD}) or the vendor name looks different. "
                           "You can still confirm, with a reason.")
            elif chk["ok"]:
                st.success("Checks passed.")
        reason = st.text_input("Reason for the manual match (required)", key=f"mm_reason_{case_id}")
        if st.button("Confirm manual match", type="primary", key=f"mm_go_{case_id}"):
            try:
                res = annie.manual_match(con, case_id, rows, reason=reason)
                ss.pop(f"mm_chk_{case_id}", None)
                ss.flash = [f"{case_id} matched manually. Status: {res['status']}."]
                st.rerun()
            except ValueError as e:
                st.error(str(e))


def unmatch_panel(case_id):
    """Stage 2: undo the match but keep the invoice, so it can be re-matched by hand in Stage 1."""
    with st.expander("Unmatch (send back to Stage 1)"):
        st.caption("Reverses the billing on the PO/PR lines (a single-invoice PO/PR becomes open again) and returns this invoice to "
                   "Stage 1 so you can match it manually. It will not be re-matched automatically. GR/SES batches raised for this match are voided.")
        reason = st.text_input("Reason (required)", key=f"um_reason_{case_id}")
        if st.button("Unmatch", key=f"um_go_{case_id}"):
            try:
                annie.unmatch_case(con, case_id, reason)
                ss.flash = [f"{case_id} unmatched and returned to Stage 1."]
                st.rerun()
            except ValueError as e:
                st.error(str(e))


def dismiss_panel(case_id):
    """Stage 1 / Stage 2: delete the invoice (billing reversed, file forgotten so it can be resubmitted)."""
    with st.expander("Dismiss invoice (delete)"):
        st.warning("This deletes the invoice and its stored file. Any billing it made on PO/PR lines is reversed "
                   "(many-to-many lines get their balance back; a single-invoice PO/PR returns to Stage 1). "
                   "GR/SES batches raised for this match are voided. The same file can be uploaded again later.")
        ok = st.checkbox("I understand and want to delete this invoice", key=f"dis_ok_{case_id}")
        if st.button("Dismiss invoice", key=f"dis_go_{case_id}", disabled=not ok):
            try:
                annie.delete_invoice_case(con, case_id, "Dismissed from the case page")
                ss.flash = [f"{case_id} dismissed and deleted."]
                ss.view_case = None
                st.rerun()
            except ValueError as e:
                st.error(str(e))


def po_page(record_id):
    """Preview of a PO/PR: the document, its line balances and the invoices charged to it."""
    d = annie.po_detail(con, record_id)
    top = st.columns([1, 6])
    if top[0].button("← Back"):
        ss.view_po = None
        st.rerun()
    p = d["doc"]["payload"]
    top[1].subheader(f"PO/PR {p.get('po_reference')}  |  {p.get('vendor_name')}  |  {d['lines'][0]['basis'] if d['lines'] else ''}")
    left, right = st.columns(2, gap="large")
    with left:
        summary(d["doc"])
        file_view(d["doc"], key="po_preview")
    with right:
        st.markdown("#### Lines and balances")
        if d["lines"] and d["lines"][0]["basis"].startswith("Services (single"):
            st.info("This PO/PR cites no amount, so it expects exactly one invoice.")
        st.dataframe(pd.DataFrame(d["lines"]), hide_index=True)
        st.markdown("#### Invoices charged to this PO/PR")
        if d["invoices"]:
            st.dataframe(pd.DataFrame(d["invoices"]), hide_index=True)
        else:
            st.caption("None yet.")


def do_post(case_id):
    """Runs when the Post download is clicked: archive the case."""
    try:
        annie.post_case(con, case_id)
        ss.flash = [f"{case_id} posted and moved to Complete Archive."]
    except Exception as e:
        ss.flash = [f"Post failed: {e}"]
    ss.view_case = None


def calloff_gr_form(case_id):
    """GR/SES batch for a call-off case: enter what was received for each line of the matched invoices."""
    d = annie.calloff_detail(con, case_id)
    po = d["po"]["payload"]
    lines = annie.calloff_lines(con, case_id)
    with st.form(f"co_gr_{case_id}"):
        st.caption("Enter what was actually received or performed in this batch, per invoice line. The case completes once receipts "
                   "cover every matched invoice.")
        a, b = st.columns(2)
        a.text_input("PO/PR reference", str(po.get("po_reference")), disabled=True)
        b.text_input("Vendor", po.get("vendor_name") or "", disabled=True)
        a.text_input("SES/GR number", annie.next_ses_number(con, po.get("category")) + " (auto-generated on submit)", disabled=True)
        b.text_input("Prepared by (name) - must match the PO/PR", po.get("prepared_by") or "", disabled=True)
        d1, d2 = st.columns(2)
        f_from = d1.date_input("Performance / service date - from", date.today(), key=f"co_f1_{case_id}")
        f_to = d2.date_input("Performance / service date - to", date.today(), key=f"co_f2_{case_id}")
        df = pd.DataFrame([dict(inv_case=l["inv_case_id"], invoice=l["invoice_reference"], inv_line=l["inv_line"], item=l["description"],
                                basis="Qty" if l["basis"] == "QTY" else "Amount", billed=l["billed"], received_so_far=l["received"],
                                receive_now=l["needed"]) for l in lines])
        edited = st.data_editor(df, hide_index=True, key=f"co_gr_ed_{case_id}",
                                disabled=[c for c in df.columns if c != "receive_now"],
                                column_config={"receive_now": st.column_config.NumberColumn("Receive now", min_value=0.0)})
        submitted = st.form_submit_button("Submit GR/SES", type="primary")
    if submitted:
        gr = dict(service_date_from=f_from.isoformat(), service_date_to=f_to.isoformat(), prepared_by=po.get("prepared_by"),
                  lines=[dict(inv_case_id=r["inv_case"], inv_line=int(r["inv_line"]), receive=float(r["receive_now"] or 0)) for r in edited.to_dict("records")])
        try:
            res = annie.submit_calloff_gr(con, case_id, gr)
            ss.flash = [f"GR/SES {res['ses_gr_number']} saved. Case status: {res['status']}."]
            st.rerun()
        except ValueError as e:
            st.error(str(e))


def _add_inv(case_id):
    ss[f"co_n_{case_id}"] = ss.get(f"co_n_{case_id}", 1) + 1


def _drop_inv(case_id):
    n = ss.get(f"co_n_{case_id}", 1)
    if n > 1:
        ss.pop(f"co_inv_{case_id}_{n - 1}", None)
        ss[f"co_n_{case_id}"] = n - 1


def calloff_page(case_id):
    """A PO/PR that cites no amount: pick the invoices that bill against it, then Confirm to close the case."""
    d = annie.calloff_detail(con, case_id)
    c, po = d["case"], d["po"]
    pp = po["payload"]
    top = st.columns([1, 6])
    if top[0].button("← Back to cases"):
        ss.view_case = None
        st.rerun()
    top[1].subheader(f"{case_id}  |  {c['status']}  |  PO/PR {pp.get('po_reference')}  |  {pp.get('vendor_name')}")
    for m in ss.pop("flash", []):
        st.success(m)
    is_open = c["status"] == annie.S_CALLOFF
    options = annie.calloff_open_invoices(con, case_id) if is_open else []
    by_id = {o["case_id"]: o for o in options}
    n = ss.setdefault(f"co_n_{case_id}", 1) if is_open else 0
    for i in range(n):   # forget a selection that is no longer an open invoice
        if ss.get(f"co_inv_{case_id}_{i}") not in [None] + list(by_id):
            ss[f"co_inv_{case_id}_{i}"] = None
    chosen = [ss.get(f"co_inv_{case_id}_{i}") for i in range(n)]
    chosen = [x for x in chosen if x]
    left, right = st.columns(2, gap="large")
    with left:
        st.markdown("#### Source documents")
        if is_open:
            docs = [("PO/PR " + str(pp.get("po_reference")), po)]
            for cid in dict.fromkeys(chosen):
                rid = con.execute("SELECT inv_record_id FROM cases WHERE case_id=?", (cid,)).fetchone()[0]
                docs.append(("Invoice " + str(by_id[cid]["invoice_reference"]), annie.get_doc(con, rid)))
        else:
            docs = [("PO/PR " + str(pp.get("po_reference")), po)] + [("Invoice " + str(i["doc"]["payload"].get("invoice_reference")), i["doc"]) for i in d["invoices"]]
        for tab, (label, doc), k in zip(st.tabs([x[0] for x in docs]), docs, range(len(docs))):
            with tab:
                summary(doc)
                file_view(doc, key=f"co{k}_{case_id}")
    with right:
        st.markdown("#### Match invoices to this Call-Off PO/PR")
        if is_open:
            st.info("This PO/PR cites no amount, so Annie does not match it automatically. Pick each invoice that bills against it "
                    "(use + Add Invoice for more), then press Confirm. The case moves to Stage 2 (Pending GR/SES).")
            if not options:
                st.warning("No open invoices are waiting in Stage 1.")
            for i in range(n):
                st.selectbox(f"Invoice {i + 1}", [None] + list(by_id), key=f"co_inv_{case_id}_{i}",
                             format_func=lambda v: "Select an open invoice..." if v is None else by_id[v]["label"])
            b1, b2, _ = st.columns([1.3, 1.3, 3])
            b1.button("+ Add Invoice", key=f"co_add_{case_id}", on_click=_add_inv, args=(case_id,))
            if n > 1:
                b2.button("Remove last", key=f"co_drop_{case_id}", on_click=_drop_inv, args=(case_id,))
            if len(chosen) != len(set(chosen)):
                st.error("The same invoice is selected more than once.")
            for cid in dict.fromkeys(chosen):
                if by_id[cid]["vendor_sim"] < annie.VENDOR_MIN:
                    st.warning(f"{cid}: vendor name looks different (similarity {by_id[cid]['vendor_sim']:.2f}).")
            reason = st.text_input("Reason for the match (required)", key=f"co_reason_{case_id}")
            if st.button("Confirm", type="primary", key=f"co_go_{case_id}", disabled=not chosen):
                try:
                    res = annie.confirm_calloff(con, case_id, chosen, reason)
                    for k in [k for k in list(ss.keys()) if str(k).startswith((f"co_inv_{case_id}_", f"co_n_{case_id}", f"co_reason_{case_id}"))]:
                        ss.pop(k, None)
                    ss.flash = [f"{case_id} matched with {len(res['invoices'])} invoice(s) and moved to Stage 2 (Pending GR/SES)."]
                    st.rerun()
                except ValueError as e:
                    st.error(str(e))
        else:
            st.success({annie.S_GR: "Matched - pending GR/SES.", annie.S_DONE: "Matched and received - complete.",
                        annie.S_ARCHIVED: "Matched and received - posted to the Complete Archive."}.get(c["status"], ""))
            st.dataframe(pd.DataFrame([dict(case=i["case_id"], invoice=i["doc"]["payload"].get("invoice_reference"),
                                            date=i["doc"]["payload"].get("date"), total=annie.doc_total(i["doc"]["payload"])) for i in d["invoices"]]),
                         hide_index=True)
            st.caption("Reason: " + (c["reason"] or "-"))
            if c["status"] == annie.S_GR:
                calloff_gr_form(case_id)
            grs = annie.case_gr_docs(con, case_id)
            if grs:
                st.markdown("**GR/SES batches for this call-off case**")
                for g in grs:
                    with st.expander(f"{g['payload']['ses_gr_number']} - {g['payload']['service_date_from']} - total {g['payload']['total']:,.2f}"):
                        st.json(g["payload"])
    if c["status"] == annie.S_DONE:
        st.divider()
        _, post_col = st.columns([5, 1])
        data, fname = annie.build_case_package(con, case_id)
        post_col.download_button("Post", data, file_name=fname, mime="application/zip", type="primary", key=f"post_{case_id}",
                                 on_click=do_post, args=(case_id,),
                                 help="Downloads the PO/PR, matched invoices and GR/SES, and moves the case to Complete Archive.")


def case_page(case_id):
    if str(case_id).startswith("CALL-"):
        return calloff_page(case_id)
    c = con.execute("SELECT * FROM cases WHERE case_id=?", (case_id,)).fetchone()
    inv, pos = annie.get_doc(con, c["inv_record_id"]), annie.case_po_docs(con, case_id)
    lines, grs = annie.case_lines(con, case_id), annie.case_gr_docs(con, case_id)
    top = st.columns([1, 6])
    if top[0].button("← Back to cases"):
        ss.view_case = None
        st.rerun()
    top[1].subheader(f"{case_id}  |  {c['status']}  |  invoice {inv['payload'].get('invoice_reference')}  |  link score {c['link_score'] or '-'}")
    for m in ss.pop("flash", []):
        st.success(m)
    left, right = st.columns(2, gap="large")
    with left:
        st.markdown("#### Source documents")
        tabs_ = st.tabs(["Invoice"] + [f"PO/PR {d['payload'].get('po_reference')}" for d in pos])
        for tab, doc, name in zip(tabs_, [inv] + pos, ["invoice"] + [f"po{i}" for i in range(len(pos))]):
            with tab:
                summary(doc)
                file_view(doc, key=name)
    with right:
        st.markdown("#### Allocation and receipts")
        if lines:
            st.dataframe(pd.DataFrame([dict(po_pr=l["po_reference"], item=l["description"], basis=BASIS_LABEL[l["basis"]],
                                            ordered=l["ordered"], billed_this_invoice=l["billed_this"],
                                            received_for_this_invoice=l["received"], still_to_receive=l["needed"],
                                            matched=("manual: " + (l["a_reason"] or "")) if l["manual"] else "automatic") for l in lines]), hide_index=True)
        if c["status"] == annie.S_OPEN_PO:
            st.warning("Not allocated yet: " + (c["reason"] or "waiting for a matching PO/PR"))
            if st.button("Retry automatic allocation"):
                annie.allocate_invoice(con, case_id, force=True)
                st.rerun()
        elif c["status"] == annie.S_GR:
            gr_form(case_id, lines)
        else:
            st.success("Receipts cover everything billed on this invoice" + (" - posted to the Complete Archive." if c["status"] == annie.S_ARCHIVED else "."))
        if grs:
            st.markdown("**GR/SES batches for this invoice-PO match**")
            for g in grs:
                with st.expander(f"{g['payload']['ses_gr_number']} - {g['payload']['service_date_from']} - total {g['payload']['total']:,.2f}"):
                    st.json(g["payload"])
    if c["status"] == annie.S_OPEN_PO:
        st.divider()
        manual_panel(case_id)
        dismiss_panel(case_id)
    elif c["status"] == annie.S_GR:
        st.divider()
        unmatch_panel(case_id)
        dismiss_panel(case_id)
    if c["status"] == annie.S_DONE:
        st.divider()
        _, post_col = st.columns([5, 1])
        data, fname = annie.build_case_package(con, case_id)
        post_col.download_button("Post", data, file_name=fname, mime="application/zip", type="primary",
                                 key=f"post_{case_id}", on_click=do_post, args=(case_id,),
                                 help="Downloads the case folder (invoice, matched PO/PRs, GR/SES) and moves the case to Complete Archive.")


# =====================================================================================
# 2B. SIDE NAVIGATION (rectangular buttons; the selected one is pressed in) AND PROCESSING TABS
# =====================================================================================
SIDE = [("dashboard", "Dashboard"), ("submit", "Submit Document"), ("review", "Intake Review"), ("processing", "Processing"),
        ("docs", "Documents"), ("archive", "Complete Archive"), ("audit", "Audit Log"), ("controls", "Controls")]
PAGE_TITLES = dict(SIDE)
PTABS = [("s1", "Stage 1 - Open cases"), ("s2", "Stage 2 - Pending GR/SES"), ("s3", "Stage 3 - Complete"), ("calloff", "Call-Offs")]
ss.setdefault("side", "dashboard")
ss.setdefault("ptab", "s1")
pending = annie.counts(con)
side_badge = {"review": pending["review"], "processing": pending["s1"] + pending["s2"] + pending["s3"] + pending["calloff"]}

st.markdown("""
<style>
[data-testid="stSidebar"] {background:#16233a;}
[data-testid="stSidebar"] h3 {color:#ffffff; letter-spacing:.6px; padding:.2rem .4rem 1rem .4rem;}
[class*="st-key-side_"] button {width:100%; justify-content:flex-start; text-align:left; border-radius:3px; padding:.6rem 1rem;
    margin:0 0 2px 0; font-weight:500; background:transparent; color:#c9d4e3; border:1px solid transparent; box-shadow:none;}
[class*="st-key-side_"] button p {color:inherit; margin:0; width:100%; text-align:left;}
[class*="st-key-side_"] button[kind="secondary"]:hover {background:#223454; border-color:#31497a; color:#ffffff;}
[class*="st-key-side_"] button[kind="primary"] {background:#0c1626; color:#ffffff; border:1px solid #060d18;
    border-left:4px solid #4da3ff; box-shadow:inset 0 3px 8px rgba(0,0,0,.6);}
[class*="st-key-side_"] button[kind="primary"]:hover {background:#0c1626; color:#ffffff; border-color:#060d18; border-left-color:#4da3ff;}
[class*="st-key-ptab_"] button {width:100%; border-radius:3px; padding:.6rem .8rem; font-weight:500; background:#f1f4f8;
    color:#33445c; border:1px solid #c5cfdc; box-shadow:0 1px 0 #ffffff inset, 0 1px 2px rgba(0,0,0,.08);}
[class*="st-key-ptab_"] button p {color:inherit; margin:0;}
[class*="st-key-ptab_"] button[kind="secondary"]:hover {background:#e6ecf4; border-color:#9fb1c8; color:#1b2a41;}
[class*="st-key-ptab_"] button[kind="primary"] {background:#dbe6f5; color:#0b2a55; border:1px solid #8fa8c8;
    border-bottom:3px solid #1f6feb; font-weight:700; box-shadow:inset 0 2px 5px rgba(0,0,0,.18);}
[class*="st-key-ptab_"] button[kind="primary"]:hover {background:#dbe6f5; color:#0b2a55; border-color:#8fa8c8; border-bottom-color:#1f6feb;}
[class*="st-key-polink_"] button {background:transparent; border:none; padding:0; min-height:0; justify-content:flex-start; box-shadow:none; color:#1556c0;}
[class*="st-key-polink_"] button p {color:inherit; margin:0; text-align:left; text-decoration:underline;}
[class*="st-key-polink_"] button:hover {background:transparent; color:#0b3a8a;}
</style>
""", unsafe_allow_html=True)


def go_side(key):
    ss.side, ss.view_case, ss.view_po = key, None, None


def set_ptab(key):
    ss.ptab, ss.view_case, ss.view_po = key, None, None


with st.sidebar:
    st.markdown("### Annie.")
    for key, label in SIDE:
        st.button(label + bubble(side_badge.get(key, 0)), key=f"side_{key}", type="primary" if ss.side == key else "secondary",
                  on_click=go_side, args=(key,))
page = ss.side


def processing_tabs():
    """Stage 1-3 as a row of rectangular tab buttons with pending-count bubbles."""
    cols = st.columns([1.5, 1.75, 1.5, 1.1, 5.4])
    for col, (key, label) in zip(cols, PTABS):
        col.button(label + bubble(pending[key]), key=f"ptab_{key}", type="primary" if ss.ptab == key else "secondary",
                   on_click=set_ptab, args=(key,))
    st.write("")
    return ss.ptab


if ss.view_po:
    po_page(ss.view_po)
    st.stop()
if ss.view_case:
    case_page(ss.view_case)
    st.stop()


# =====================================================================================
# 4. CASE LIST (Stage 1 invoices / Stage 2 / Stage 3)
# =====================================================================================
CASE_SORTS = ["case_id", "vendor", "invoice_reference", "amount", "link_score", "status", "updated_at"]


def case_list(statuses, key, empty_msg, show_reason=False):
    rows = annie.case_rows(con, statuses)
    if not rows:
        st.info(empty_msg)
        return
    rows = controls(key, rows, CASE_SORTS)
    widths = [1.3, 2.0, 1.6, 2.0, 1.2, 0.9, 3.0, 0.9]
    for col, h in zip(st.columns(widths), ["Case", "Vendor", "Invoice", "PO/PR", "Amount", "Score", "Reason" if show_reason else "Status", ""]):
        col.markdown(f"**{h}**")
    for r in rows:
        cols = st.columns(widths)
        vals = [r["case_id"], r["vendor"], r["invoice_reference"], r["po_references"] or "-",
                f"{r['amount']:,.2f}" if r["amount"] is not None else "-", r["link_score"] or "-",
                (r["reason"] or r["status"]) if show_reason else r["status"]]
        for col, v in zip(cols, vals):
            col.write(v)
        if cols[7].button("Open", key=f"{key}_{r['case_id']}"):
            ss.view_case = r["case_id"]
            st.rerun()


# =====================================================================================
# 5. PAGE HEADER
# =====================================================================================
for m in ss.pop("flash", []):
    st.success(m)
st.header(PAGE_TITLES[page])


# =====================================================================================
# 6. PAGE: SUBMIT DOCUMENT (manual upload + email intake, both run Module 1 first)
# =====================================================================================
if page == "dashboard":
    pass   # intentionally empty for now

elif page == "submit":
    st.markdown("Every document goes through **Module 1** (duplicate check, extraction, field check) first. "
                "Only documents with every mandatory field (`READY_FOR_MODULE_2`) enter Annie; the rest wait in **Intake Review**. "
                "PO/PRs and invoices can arrive in any order and in any number: invoices are allocated to PO/PR lines as soon as the balances allow.")
    up_col, mail_col = st.columns(2, gap="large")
    with up_col:
        st.subheader("A. Upload documents")
        files = st.file_uploader("PDF, image, DOCX, TXT, CSV or XLSX", type=UPLOAD_TYPES, accept_multiple_files=True)
        if st.button("Run Module 1 and route", type="primary") and files:
            with st.spinner("Running Module 1..."):
                for f in sorted(files, key=lambda f: f.name):   # process in file-name order
                    try:
                        ss.results += annie.ingest_file(con, f.name, f.getvalue(), channel="UPLOAD")
                    except Exception as e:
                        st.error(f"{f.name}: {e}")
            st.rerun()
    with mail_col:
        st.subheader("B. Email inbox")
        em = annie.email_config_status()
        if em["ok"]:
            st.write("Mailbox configured" + ": ailittest2026@gmail.com" + (f" (subject filter: `{em['subject_filter']}`)" if em["subject_filter"] else "") + ".")
        else:
            st.warning("Add " + ", ".join(em["missing"]) + " to the `.env` file (see `.env.example`), then restart the app.")
        if st.button("Check inbox now", disabled=not em["ok"]):
            with st.spinner("Reading unread emails..."):
                try:
                    res, stats = annie.check_inbox(con)
                    ss.results += res
                    st.success(f"{stats['emails']} email(s), {stats['attachments']} attachment(s) processed.")
                except Exception as e:
                    st.error(f"Email check failed: {e}")
    st.divider()
    st.subheader("Results this session")
    render_results(ss.results)


# =====================================================================================
# 7. PAGE: INTAKE REVIEW (Re-extract / Dismiss, then the missing fields)
# =====================================================================================
elif page == "review":
    st.caption("Documents Module 1 could not pass to Module 2. Re-extract after the extraction code or reference samples "
               "improve, or dismiss. A dismissed document is removed completely and can be resubmitted later.")
    review = annie.intake_review_rows(con)
    if not review:
        st.info("Nothing waiting for review.")
    else:
        review = controls("review", review, ["source_file", "m1_status", "channel", "received_at"])
    for r in review:
        k = r["record_id"]
        with st.expander(f"{r['source_file']} - {r['m1_status']} - {r['channel']} - {r['received_at']}"):
            lcol, rcol = st.columns(2, gap="large")
            with lcol:
                file_view(r, height=520, key=f"rev_{k}")
            with rcol:
                b1, b2 = st.columns(2)
                if b1.button("Re-extract from file", key=f"rx_{k}", type="primary"):
                    try:
                        ss.results.append(annie.reextract_intake(con, k))
                        st.rerun()
                    except ValueError as e:
                        st.error(str(e))
                if b2.button("Dismiss", key=f"dis_{k}"):
                    annie.dismiss_intake(con, k, "Dismissed from Intake Review")
                    st.rerun()
                st.markdown("**Missing fields**")
                if r["missing"]:
                    st.markdown("\n".join(f"- {annie.LABELS.get(m, m)}" for m in r["missing"]))
                else:
                    st.write("None listed - " + str(r["note"]))


# =====================================================================================
# 8. PAGE: PROCESSING (Stage 1 / Stage 2 / Stage 3 as top tabs)
# =====================================================================================
elif page == "processing":
    ptab = processing_tabs()
    if ptab == "s1":
        st.subheader("Invoices awaiting PO/PR")
        st.caption("Invoices that could not be fully allocated yet: the PO/PR has not arrived, no line matches, or the bill exceeds what is left on the PO/PR.")
        case_list([annie.S_OPEN_PO], "s1a", "No invoices waiting for a PO/PR.", show_reason=True)
        st.divider()
        st.subheader("PO/PRs with open balance")
        st.caption("Ordered, billed and received per PO/PR. Partly billed PO/PRs stay here until fully billed. Services with no amount cited expect one invoice only. Click a PO/PR reference to preview it.")
        po = [r for r in annie.po_rows(con) if r["status"] != annie.S_BILLED]
        if not po:
            st.info("No PO/PR with an open balance.")
        else:
            po = controls("s1b", po, ["po_reference", "vendor", "mode", "status", "ordered", "billed", "received", "open_to_bill", "date"])
            st.download_button("Download CSV", pd.DataFrame(po).drop(columns=["record_id"]).to_csv(index=False),
                               file_name="po_pr_open_balance.csv", mime="text/csv", key="po_csv")
            widths = [1.5, 2.2, 2.0, 1.8, 1.0, 1.0, 1.0, 1.1, 1.2]
            heads = ["PO/PR (click to preview)", "Vendor", "Type", "Status", "Ordered", "Billed", "Received", "Open to bill", "Date"]
            for col, h in zip(st.columns(widths), heads):
                col.markdown(f"**{h}**")
            fmt = lambda v: "-" if v is None else f"{v:,.2f}"
            for r in po:
                cols = st.columns(widths)
                if cols[0].button(str(r["po_reference"]), key=f"polink_{r['record_id']}"):
                    ss.view_po = r["record_id"]
                    st.rerun()
                for col, v in zip(cols[1:], [r["vendor"], r["mode"], r["status"], fmt(r["ordered"]), fmt(r["billed"]),
                                             fmt(r["received"]), fmt(r["open_to_bill"]), r["date"]]):
                    col.write(v)

    elif ptab == "s2":
        st.caption("Invoice is allocated to PO/PR line(s), or a Call-Off case is matched, but receipts do not yet cover it. Open a case to record a GR/SES batch.")
        case_list([annie.S_GR], "s2", "No cases pending GR/SES.")

    elif ptab == "s3":
        st.caption("Receipts cover the invoice (or the matched invoices of a Call-Off case). Open a case and press Post to download its folder and archive it.")
        case_list([annie.S_DONE], "s3", "No complete cases yet.")

    elif ptab == "calloff":
        st.caption("PO/PRs that cite no amount are never matched automatically. Each opens its own case: open it, pick the invoices that bill "
                   "against it, and Confirm. It then moves to Stage 2 (Pending GR/SES) and reaches Stage 3 once receipts cover the invoices.")
        rows = annie.calloff_rows(con, [annie.S_CALLOFF])
        if not rows:
            st.info("No Call-Off PO/PR waiting.")
        else:
            rows = controls("calloffs", rows, ["case_id", "po_references", "vendor", "type", "date", "status"])
            widths = [1.4, 1.7, 2.6, 1.1, 1.3, 1.8, 0.9]
            for col, h in zip(st.columns(widths), ["Case", "PO/PR", "Vendor", "Type", "PO date", "Status", ""]):
                col.markdown(f"**{h}**")
            for r in rows:
                cols = st.columns(widths)
                for col, v in zip(cols, [r["case_id"], r["po_references"], r["vendor"], r["type"], r["date"], r["status"]]):
                    col.write(v)
                if cols[6].button("Open", key=f"co_open_{r['case_id']}"):
                    ss.view_case = r["case_id"]
                    st.rerun()


# =====================================================================================
# 9. PAGE: DOCUMENTS (Module 1 + Module 2 movements per record)
# =====================================================================================
elif page == "docs":
    st.caption("One row per record: how it arrived, what Module 1 decided, and what Module 2 did with it.")
    rows = controls("docs", annie.list_documents(con),
                    ["record_id", "doc_type", "channel", "source_file", "module1_status", "route", "module2_outcome", "case_id", "received_at"])
    st.dataframe(pd.DataFrame(rows))


# =====================================================================================
# 10. PAGE: COMPLETE ARCHIVE (posted cases)
# =====================================================================================
elif page == "archive":
    st.caption("Posted cases kept as history. Open a case to view it, or download its package again.")
    rows = annie.archive_rows(con)
    if not rows:
        st.info("No posted cases yet.")
    else:
        rows = controls("archive", rows, CASE_SORTS + ["posted_at"])
        widths = [1.3, 2.0, 1.6, 2.0, 1.2, 2.0, 0.9, 1.2]
        for col, h in zip(st.columns(widths), ["Case", "Vendor", "Invoice", "PO/PR", "Amount", "Posted", "", ""]):
            col.markdown(f"**{h}**")
        for r in rows:
            cols = st.columns(widths)
            vals = [r["case_id"], r["vendor"], r["invoice_reference"], r["po_references"] or "-",
                    f"{r['amount']:,.2f}" if r["amount"] is not None else "-", r["posted_at"]]
            for col, v in zip(cols, vals):
                col.write(v)
            if cols[6].button("Open", key=f"ar_{r['case_id']}"):
                ss.view_case = r["case_id"]
                st.rerun()
            pkg = annie.read_archive_package(con, r["case_id"])
            if pkg:
                cols[7].download_button("Package", pkg, file_name=f"{r['case_id']}.zip", mime="application/zip", key=f"arz_{r['case_id']}")


# =====================================================================================
# 11. PAGE: AUDIT LOG (Module 1 and Module 2 events)
# =====================================================================================
elif page == "audit":
    stages = st.multiselect("Stage", ["MODULE 1", "MODULE 2"], default=["MODULE 1", "MODULE 2"])
    rows = controls("audit", annie.list_log(con, stages), ["ts", "stage", "event", "record_id", "case_id"])
    st.dataframe(pd.DataFrame(rows))


# =====================================================================================
# 12. PAGE: CONTROLS (sample data, database reset, Module 1 status)
# =====================================================================================
elif page == "controls":
    st.subheader("Sample data")
    if st.button("Load sample documents") and SAMPLES.exists():
        ss.results += annie.load_samples(con, SAMPLES)
        st.rerun()
    st.caption("Optional: loads pre-extracted sample JSONs (if a samples folder with JSON files exists) in place of Module 1 text extraction.")
    st.divider()
    st.subheader("Database")
    if st.button("Reset database"):
        annie.reset_db(con)
        ss.results = []
        st.rerun()
    st.caption("Deletes every document, case, allocation, GR/SES, archive entry and log line.")
    st.divider()
    st.subheader("Module 1 status")
    ref = annie.get_reference_summary()
    st.write(f"OCR engine: **{ref['ocr_engine'] or 'not available'}**")
    st.write(f"Reference samples: {ref['purchase_orders']} PO/PR, {ref['supplier_invoices']} invoices")
    em = annie.email_config_status()
    st.write("Email: " + ("configured" if em["ok"] else "missing " + ", ".join(em["missing"])))
    st.caption(f"Line-match threshold >= {annie.LINK_THRESHOLD}   |   Vendor-name gate >= {annie.VENDOR_MIN}")
