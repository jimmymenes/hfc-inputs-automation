"""
HFC Inputs Auto-Populator — Web App
=====================================
Run with:  streamlit run app.py
Share with your team by sending them:  http://<your-ip-address>:8501

Install dependencies first:
    pip install streamlit openpyxl google-api-python-client google-auth-httplib2 google-auth-oauthlib boxsdk
"""

import io
import os
import re
import tempfile

import streamlit as st

# ─────────────────────────────────────────────────────────────
# Page config
# ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="HFC Inputs Auto-Populator",
    page_icon="📋",
    layout="centered",
)

st.title("📋 HFC Inputs Auto-Populator")
st.caption("Innovations for Poverty Action · Data Management Tool")
st.markdown("---")
st.markdown(
    "Paste your SurveyCTO XLSForm and an HFC inputs template — "
    "the tool will automatically populate all six DMS sheets "
    "*(other specify, outliers, constraints, logic, enumstats, text audits)*."
)

# ─────────────────────────────────────────────────────────────
# Core population logic (shared by both tabs)
# ─────────────────────────────────────────────────────────────

def clean_label(s):
    if not s:
        return ""
    return re.sub(r"<[^>]+>", "", str(s)).strip()


def classify_survey_variables(survey_bytes):
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(survey_bytes), read_only=True, data_only=True)
    ws = wb["survey"]

    SKIP_TYPES = {
        "begin group", "end group", "begin repeat", "end repeat",
        "note", "calculate", "start", "end", "deviceid",
        "phonenumber", "username", "caseid", "enumerator",
    }

    groups, numerics, selects, texts = [], [], [], []

    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue
        def g(idx):
            return row[idx] if len(row) > idx and row[idx] else ""
        t, n, l = g(0), g(1), clean_label(g(2))
        if not n:
            continue
        base = t.split()[0] if t else ""
        if "begin" in str(t) and "group" in str(t):
            groups.append({"name": n, "label": l})
        elif base in SKIP_TYPES or not base:
            continue
        elif base in ("integer", "decimal"):
            numerics.append({"name": n, "type": t, "label": l})
        elif base in ("select_one", "select_multiple"):
            selects.append({"name": n, "type": t, "label": l})
        elif base == "text":
            texts.append({"name": n, "label": l})

    wb.close()
    return numerics, selects, texts, groups


def existing_vars(ws, col=0):
    return {str(row[col]).strip() for row in ws.iter_rows(min_row=2, values_only=True) if row[col]}


def populate_other_specify(ws, selects, texts):
    text_names = {v["name"] for v in texts}
    existing = existing_vars(ws)
    added = 0
    for v in selects:
        name = v["name"]
        for suffix in ("_oth", "_oth_r"):
            child = name + suffix
            if child in text_names and name not in existing:
                child_label = next((t["label"] for t in texts if t["name"] == child), "Other, specify")
                ws.append((name, v["label"], child, child_label, None, None, "id member_name village enumerator_id"))
                existing.add(name)
                added += 1
                break
    return added


def populate_outliers(ws, numerics):
    existing = existing_vars(ws)
    added = 0
    for v in numerics:
        if v["name"] not in existing:
            ws.append((v["name"], v["label"], "enum", "sd", 3, None, None, None, "id member_name enumerator_id"))
            added += 1
    return added


def populate_constraints(ws, numerics):
    existing = existing_vars(ws)

    def bounds(name):
        n = name.lower()
        if "age" in n: return (15, 18, 70, 100)
        if "nbr" in n or "count" in n or "hh_own" in n: return (0, 0, 20, 100)
        if "year" in n and "begin" in n: return (1970, 1990, 2025, 2026)
        if "area" in n or "plot" in n: return (0, 0, 20, 50)
        if "hrs" in n or "hour" in n or "time" in n: return (0, 0, 14, 24)
        if "week" in n: return (0, 0, 80, 168)
        if "last1m" in n or "last_mont" in n: return (0, 10000, 2000000, 15000000)
        if "last12m" in n or "last12" in n: return (0, 100000, 10000000, 50000000)
        if "borrowed" in n or "loan" in n: return (0, 50000, 5000000, 20000000)
        if "saved" in n or "saving" in n: return (0, 0, 5000000, 50000000)
        if any(k in n for k in ("val", "sales", "rev", "cost", "earn", "spend", "spent", "expense")):
            return (0, 0, 1000000, 10000000)
        if "accessed" in n: return (0, 0, 26, 26)
        return (0, 0, None, None)

    added = 0
    for v in numerics:
        if v["name"] not in existing:
            hard_min, soft_min, soft_max, hard_max = bounds(v["name"])
            ws.append((v["name"], v["label"], hard_min, soft_min, soft_max, hard_max, None, None, None))
            added += 1
    return added


def populate_logic(ws, numerics):
    existing = existing_vars(ws)
    num_names = {v["name"]: v for v in numerics}
    added = 0

    for v in numerics:
        n = v["name"]
        if "last12m" in n:
            monthly = n.replace("last12m", "last1m")
            if monthly in num_names and n not in existing:
                ws.append((n, v["label"], f"{n} >= {monthly}", f"!missing({monthly})", None, None, "id member_name enumerator_id"))
                existing.add(n)
                added += 1

    for var, assert_expr, cond in [
        ("s6_time_paid_work",  "s6_time_paid_work + s6_time_unpaid_act <= 24", "!missing(s6_time_unpaid_act)"),
        ("week_hrs_spend",     "week_hrs_spend <= 168",                         None),
        ("year_begin",         "year_begin >= 1970 & year_begin <= 2026",        None),
    ]:
        if var in num_names and var not in existing:
            ws.append((var, num_names[var]["label"], assert_expr, cond, None, None, None))
            existing.add(var)
            added += 1

    return added


def populate_enumstats(ws, numerics):
    existing = existing_vars(ws)
    added = 0
    for v in numerics:
        if v["name"] not in existing:
            combine = "yes" if re.search(r"_r\??$|_r\d", v["name"].lower()) else None
            ws.append((v["name"], v["label"], "yes", None, "number", "yes", "yes", "yes", combine, None))
            added += 1
    return added


def populate_text_audits(ws, groups):
    existing = existing_vars(ws)
    added = 0
    for g in groups:
        if g["name"] not in existing and g["name"].startswith("section"):
            ws.append((g["name"], None, None, g["label"]))
            added += 1
    return added


def run_population(survey_bytes: bytes, template_bytes: bytes) -> bytes:
    from openpyxl import load_workbook

    numerics, selects, texts, groups = classify_survey_variables(survey_bytes)

    wb = load_workbook(io.BytesIO(template_bytes), keep_vba=True)

    results = {
        "other specify": populate_other_specify(wb["other specify"], selects, texts),
        "outliers":      populate_outliers(wb["outliers"], numerics),
        "constraints":   populate_constraints(wb["constraints"], numerics),
        "logic":         populate_logic(wb["logic"], numerics),
        "enumstats":     populate_enumstats(wb["enumstats"], numerics),
        "text audits":   populate_text_audits(wb["text audits"], groups),
    }

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read(), results, len(numerics), len(selects), len(texts)


def show_results(output_bytes, results, n_num, n_sel, n_txt, filename):
    st.success("Populated successfully!")

    col1, col2, col3 = st.columns(3)
    col1.metric("Numeric variables", n_num)
    col2.metric("Select variables", n_sel)
    col3.metric("Text variables", n_txt)

    st.markdown("**Rows added per sheet:**")
    for sheet, count in results.items():
        icon = "✅" if count > 0 else "➖"
        st.markdown(f"{icon} `{sheet}` — **+{count}** rows")

    st.download_button(
        label="⬇️  Download populated HFC inputs file",
        data=output_bytes,
        file_name=filename,
        mime="application/vnd.ms-excel.sheet.macroEnabled.12",
        use_container_width=True,
    )


# ─────────────────────────────────────────────────────────────
# Tab 1: Upload files directly
# Tab 2: Use Google Drive + Box IDs
# ─────────────────────────────────────────────────────────────

tab1, tab2 = st.tabs(["📁  Upload Files", "☁️  Google Drive & Box"])

# ── TAB 1: UPLOAD ────────────────────────────────────────────
with tab1:
    st.subheader("Upload your files")
    st.markdown("Upload the XLSForm survey and the HFC inputs template from your computer.")

    survey_file   = st.file_uploader("SurveyCTO XLSForm (.xlsx)", type=["xlsx"], key="upload_survey")
    template_file = st.file_uploader("HFC Inputs Template (.xlsm)", type=["xlsm"], key="upload_template")

    if survey_file and template_file:
        out_name = template_file.name.replace(".xlsm", "_populated.xlsm")
        if st.button("Run Auto-Populate", key="btn_upload", use_container_width=True, type="primary"):
            with st.spinner("Populating HFC inputs..."):
                try:
                    output_bytes, results, n_num, n_sel, n_txt = run_population(
                        survey_file.read(), template_file.read()
                    )
                    show_results(output_bytes, results, n_num, n_sel, n_txt, out_name)
                except Exception as e:
                    st.error(f"Something went wrong: {e}")
    else:
        st.info("Upload both files above to enable the run button.")


# ── TAB 2: GOOGLE DRIVE LINK + BOX PATH ──────────────────────
with tab2:
    st.subheader("Google Drive link & Box folder path")
    st.markdown(
        "Paste the Google Drive share link for your survey and the Box folder path "
        "where your HFC template lives. No credentials needed."
    )

    # ── Helper: extract file ID from Google Drive URL ──────────
    def extract_gdrive_id(url: str):
        import re
        patterns = [
            r"/d/([a-zA-Z0-9_-]{20,})",
            r"id=([a-zA-Z0-9_-]{20,})",
        ]
        for p in patterns:
            m = re.search(p, url)
            if m:
                return m.group(1)
        return None

    def download_gdrive_file(file_id: str) -> bytes:
        import requests
        # Try spreadsheet export first, then generic Drive download
        for url in [
            f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx",
            f"https://drive.google.com/uc?export=download&id={file_id}",
        ]:
            r = requests.get(url, timeout=30)
            if r.status_code == 200 and len(r.content) > 1000:
                return r.content
        raise ValueError(
            "Could not download the file. Make sure the Google Drive link is set to "
            "'Anyone with the link can view'."
        )

    # ── Helper: scan Box folder for .xlsm files ────────────────
    def scan_box_folder(folder_path: str):
        from pathlib import Path
        p = Path(folder_path)
        if not p.exists():
            return None, []
        files = sorted(p.glob("*.xlsm"))
        return p, [f for f in files]

    # ── Google Drive link input ────────────────────────────────
    st.markdown("**1. Survey form — Google Drive share link**")
    gdrive_url = st.text_input(
        "Paste the Google Drive share link",
        placeholder="https://docs.google.com/spreadsheets/d/.../edit?usp=sharing",
        label_visibility="collapsed",
    )

    gdrive_file_id = None
    if gdrive_url:
        gdrive_file_id = extract_gdrive_id(gdrive_url)
        if gdrive_file_id:
            st.success(f"File ID detected: `{gdrive_file_id}`")
        else:
            st.error("Could not read a file ID from that URL. Make sure you copied the full share link.")

    # ── Box folder path input ──────────────────────────────────
    st.markdown("**2. HFC inputs template — Box folder path**")
    box_folder_path = st.text_input(
        "Paste the Box folder path",
        placeholder=r"C:\Users\yourname\Box\IPA_Project\...\1_inputs",
        label_visibility="collapsed",
    )

    box_folder = None
    template_files = []
    selected_template = None

    if box_folder_path:
        box_folder, template_files = scan_box_folder(box_folder_path)
        if box_folder is None:
            st.error("Folder not found. Make sure Box Drive is synced and the path is correct.")
        elif not template_files:
            st.warning("No .xlsm files found in that folder.")
        else:
            selected_template = st.selectbox(
                "Select HFC inputs template",
                options=template_files,
                format_func=lambda f: f.name,
            )

    # ── Output filename ────────────────────────────────────────
    if selected_template:
        default_out = selected_template.name.replace(".xlsm", "_populated.xlsm")
    else:
        default_out = "hfc_inputs_populated.xlsm"

    output_name = st.text_input("Output filename (saved to the same Box folder)", value=default_out)

    # ── Run button ─────────────────────────────────────────────
    st.markdown("")
    ready = gdrive_file_id and selected_template

    if st.button("Run Auto-Populate", key="btn_cloud", use_container_width=True, type="primary", disabled=not ready):
        with st.spinner("Downloading survey, populating and saving to Box..."):
            try:
                # Download survey from Google Drive (no auth)
                survey_bytes = download_gdrive_file(gdrive_file_id)

                # Read template from local Box folder
                template_bytes = selected_template.read_bytes()

                # Populate
                output_bytes, results, n_num, n_sel, n_txt = run_population(survey_bytes, template_bytes)

                # Save populated file back to the same Box folder
                out_path = box_folder / output_name
                out_path.write_bytes(output_bytes)

                show_results(output_bytes, results, n_num, n_sel, n_txt, output_name)
                st.info(f"Saved to Box folder: `{out_path}`")

            except Exception as e:
                st.error(f"Something went wrong: {e}")

    if not ready:
        st.info("Complete both fields above to enable the run button.")


# ─────────────────────────────────────────────────────────────
# Footer
# ─────────────────────────────────────────────────────────────
st.markdown("---")
st.caption("Built for IPA Tanzania · Questions? Contact jjairo@poverty-action.org")
