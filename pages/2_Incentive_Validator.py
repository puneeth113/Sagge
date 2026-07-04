import os
import importlib.util
from datetime import date, datetime

import streamlit as st
import pandas as pd


def _load_utils():
    this_dir = os.path.dirname(os.path.abspath(__file__))
    root_dir = os.path.dirname(this_dir)
    for candidate in (
        os.path.join(this_dir, "utils.py"), os.path.join(this_dir, "Utils.py"),
        os.path.join(root_dir, "utils.py"), os.path.join(root_dir, "Utils.py"),
    ):
        if os.path.exists(candidate):
            spec = importlib.util.spec_from_file_location("hr_utils", candidate)
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            return mod
    raise FileNotFoundError("Could not find utils.py.")


def _load_module(module_filename, module_name):
    this_dir = os.path.dirname(os.path.abspath(__file__))
    root_dir = os.path.dirname(this_dir)
    for candidate in (os.path.join(this_dir, module_filename), os.path.join(root_dir, module_filename)):
        if os.path.exists(candidate):
            spec = importlib.util.spec_from_file_location(module_name, candidate)
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            return mod
    raise FileNotFoundError(f"Could not find {module_filename}.")


_u = _load_utils()
render_top_nav = _u.render_top_nav
to_excel_bytes = _u.to_excel_bytes
safe_error_message = _u.safe_error_message
read_any_table = _u.read_any_table

db = _load_module("db.py", "hr_db")
auth = _load_module("auth.py", "hr_auth")
auth.init_db(db)

st.set_page_config(page_title="Incentive Tracker", page_icon="💰", layout="wide")
render_top_nav("Incentive Validator")

st.title("💰 Incentive Tracker")

# ---------------------------------------------------------------------- #
# Login gate — everything below requires a signed-in user
# ---------------------------------------------------------------------- #
try:
    user = auth.require_login()
except RuntimeError as e:
    st.error(str(e))
    st.info("This page requires a MySQL connection configured in `.streamlit/secrets.toml`. See db.py for the required format.")
    st.stop()

auth.render_user_badge()
st.divider()

UPLOAD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "uploads", "incentive_approvals")
os.makedirs(UPLOAD_DIR, exist_ok=True)


def _save_approval_document(uploaded_file, input_id_hint: str) -> str:
    """Saves the uploaded approval document to local disk and returns its
    path. In production, swap this for secure object storage (e.g. S3 with
    private ACLs) rather than local disk — this is fine for a single-server
    internal deployment, not for anything public-facing."""
    if uploaded_file is None:
        return None
    safe_name = f"{input_id_hint}_{uploaded_file.name}".replace(" ", "_")
    path = os.path.join(UPLOAD_DIR, safe_name)
    with open(path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return path


# ============================================================================
# BRANCH POC VIEW — upload only (individual or bulk)
# ============================================================================
if user["role"] == "BranchManager":

    st.caption("Submit incentive requests for your branch. Each request goes to HR for approval before it's paid.")

    tab_single, tab_bulk, tab_history = st.tabs(["✍️ Individual Entry", "📤 Bulk Upload", "📋 My Submissions"])

    with tab_single:
        st.markdown("#### Step 1 — Look up the employee by ERP/OIS number")
        erp_input = st.text_input("ERP/OIS Number (11 digits)", max_chars=11, key="poc_erp")

        employee = None
        if erp_input:
            try:
                employee = db.get_employee_by_erp(erp_input.strip())
            except Exception as e:
                st.error(safe_error_message(e, context="looking up the employee"))

            if not employee:
                st.warning("No active employee found with that ERP/OIS number.")

        if employee:
            st.success(f"Found: **{employee['first_name']} {employee['last_name']}** — {employee['designation']} ({employee['branch_name']})")

            st.markdown("#### Step 2 — Incentive details")
            with st.form("single_incentive_form", clear_on_submit=True):
                c1, c2 = st.columns(2)
                with c1:
                    incentive_type = st.text_input("Incentive Type", placeholder="e.g. Performance Bonus")
                    amount = st.number_input("Amount (₹)", min_value=0.01, value=500.0)
                with c2:
                    applicable_month = st.date_input("Applicable Month (payroll month)", value=date.today().replace(day=1))
                    remarks = st.text_input("Remarks")

                approval_doc = st.file_uploader("Approval Document (email/screenshot/PDF)", type=["pdf", "png", "jpg", "jpeg", "eml", "msg"])

                submitted = st.form_submit_button("Submit for Approval")

                if submitted:
                    if not incentive_type.strip():
                        st.error("Incentive Type is required.")
                    else:
                        try:
                            doc_path = _save_approval_document(approval_doc, f"pending_{employee['erp_id']}")
                            input_id = db.submit_incentive_request(
                                employee_id=employee["employee_id"],
                                incentive_type=incentive_type.strip(),
                                amount=amount,
                                applicable_month=applicable_month.isoformat(),
                                remarks=remarks,
                                branch_id=employee["branch_id"],
                                submitted_by=user["user_id"],
                                approval_document_path=doc_path,
                            )
                            st.success(f"✅ Submitted as **{input_id}**. It's now pending HR approval.")
                        except Exception as e:
                            st.error(safe_error_message(e, context="submitting the incentive request"))

    with tab_bulk:
        st.caption(
            "Columns required: **ERP ID, Incentive Type, Amount, Applicable Month, Remarks**. "
            "Name/Designation are looked up automatically — don't include them in your sheet."
        )
        sample = pd.DataFrame([
            {"ERP ID": "12345678901", "Incentive Type": "Performance Bonus", "Amount": 1000,
             "Applicable Month": "2026-08-01", "Remarks": "Q2 target achieved"},
        ])
        st.download_button(
            "⬇️ Download sample template",
            data=to_excel_bytes({"Template": sample}),
            file_name="incentive_bulk_template.xlsx",
        )

        uploaded_bulk = st.file_uploader("Upload bulk incentive sheet (.xlsx or .csv)", type=["xlsx", "xls", "csv"], key="poc_bulk")
        bulk_approval_doc = st.file_uploader(
            "Approval Document covering this batch (applied to all rows)",
            type=["pdf", "png", "jpg", "jpeg", "eml", "msg"], key="poc_bulk_doc",
        )

        if uploaded_bulk:
            bulk_df = read_any_table(uploaded_bulk)
            st.dataframe(bulk_df, use_container_width=True)

            required = {"ERP ID", "Incentive Type", "Amount", "Applicable Month"}
            missing = required - set(bulk_df.columns)
            if missing:
                st.error(f"Missing required column(s): {', '.join(missing)}")
            elif st.button("Submit all rows for approval"):
                doc_path = _save_approval_document(bulk_approval_doc, "bulk_batch") if bulk_approval_doc else None
                success_count, errors = 0, []
                for idx, row in bulk_df.iterrows():
                    try:
                        emp = db.get_employee_by_erp(str(row["ERP ID"]).strip())
                        if not emp:
                            errors.append(f"Row {idx+2}: ERP {row['ERP ID']} not found")
                            continue
                        db.submit_incentive_request(
                            employee_id=emp["employee_id"],
                            incentive_type=str(row["Incentive Type"]),
                            amount=float(row["Amount"]),
                            applicable_month=pd.to_datetime(row["Applicable Month"]).date().isoformat(),
                            remarks=str(row.get("Remarks", "")),
                            branch_id=emp["branch_id"],
                            submitted_by=user["user_id"],
                            approval_document_path=doc_path,
                        )
                        success_count += 1
                    except Exception as e:
                        errors.append(f"Row {idx+2}: {safe_error_message(e)}")

                st.success(f"✅ Submitted {success_count} of {len(bulk_df)} rows.")
                if errors:
                    st.warning("Some rows had issues:")
                    for err in errors:
                        st.caption(f"- {err}")

    with tab_history:
        try:
            history = db.list_requests_for_branch(user["branch_id"])
            if history:
                st.dataframe(pd.DataFrame(history), use_container_width=True, height=420)
            else:
                st.info("No submissions yet.")
        except Exception as e:
            st.error(safe_error_message(e, context="loading your submission history"))


# ============================================================================
# HR APPROVER VIEW — review, approve/reject, download for payroll
# ============================================================================
elif user["role"] in ("HRAdmin", "SuperAdmin"):

    tab_review, tab_payroll = st.tabs(["✅ Pending Approvals", "📦 Approved — Ready for Payroll"])

    with tab_review:
        try:
            pending = db.list_pending_requests()
        except Exception as e:
            st.error(safe_error_message(e, context="loading pending requests"))
            pending = []

        if not pending:
            st.info("No pending incentive requests.")
        else:
            st.caption(f"{len(pending)} request(s) awaiting your review.")
            for req in pending:
                with st.container(border=True):
                    c1, c2 = st.columns([3, 1])
                    with c1:
                        st.markdown(
                            f"**{req['input_id']}** — {req['first_name']} {req['last_name']} "
                            f"({req['designation']}, {req['branch_name']})"
                        )
                        st.caption(
                            f"ERP: {req['erp_id']} | {req['incentive_type']} | ₹{req['amount']:,.2f} | "
                            f"Applicable: {req['applicable_month']} | Submitted: {req['submitted_at']}"
                        )
                        if req["remarks"]:
                            st.caption(f"Remarks: {req['remarks']}")
                        if req["approval_document_path"] and os.path.exists(req["approval_document_path"]):
                            with open(req["approval_document_path"], "rb") as f:
                                st.download_button(
                                    "📎 View approval document", data=f.read(),
                                    file_name=os.path.basename(req["approval_document_path"]),
                                    key=f"doc_{req['request_id']}",
                                )
                        elif req["approval_document_path"]:
                            st.caption("⚠️ Approval document was recorded but the file is no longer found on disk.")

                    with c2:
                        if st.button("✅ Approve", key=f"approve_{req['request_id']}"):
                            db.review_request(req["request_id"], approve=True, reviewed_by=user["user_id"])
                            st.rerun()
                        if st.button("❌ Reject", key=f"reject_{req['request_id']}"):
                            st.session_state[f"show_reject_{req['request_id']}"] = True

                        if st.session_state.get(f"show_reject_{req['request_id']}"):
                            reason = st.text_input("Rejection reason", key=f"reason_{req['request_id']}")
                            if st.button("Confirm Reject", key=f"confirm_reject_{req['request_id']}"):
                                db.review_request(req["request_id"], approve=False, reviewed_by=user["user_id"], rejection_reason=reason)
                                st.rerun()

    with tab_payroll:
        try:
            approved = db.list_approved_for_payroll()
        except Exception as e:
            st.error(safe_error_message(e, context="loading approved incentives"))
            approved = []

        if not approved:
            st.info("No approved incentives waiting to be processed.")
        else:
            approved_df = pd.DataFrame(approved)
            st.success(f"{len(approved_df)} approved incentive(s) ready for payroll — total ₹{approved_df['amount'].sum():,.2f}")
            st.dataframe(approved_df, use_container_width=True, height=400)

            c1, c2 = st.columns(2)
            with c1:
                st.download_button(
                    "⬇️ Download for payroll processing",
                    data=to_excel_bytes({"Approved Incentives": approved_df}),
                    file_name=f"approved_incentives_{datetime.now().strftime('%Y%m%d')}.xlsx",
                )
            with c2:
                if st.button("🔒 Mark all as processed in payroll"):
                    ids = [r["request_id"] for r in approved]
                    db.mark_paid(ids)
                    st.success("Marked as processed. They won't appear here again next cycle.")
                    st.rerun()

else:
    st.warning(f"Your role ({user['role']}) doesn't have a defined view on this page yet.")
