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
    st.subheader("Upload your survey form")
    st.markdown("Upload your SurveyCTO XLSForm. The HFC inputs template is pre-loaded.")

    survey_file = st.file_uploader("SurveyCTO XLSForm (.xlsx)", type=["xlsx"], key="upload_survey")

    # Load bundled template
    BUNDLED_TEMPLATE = os.path.join(os.path.dirname(__file__), "hfc_inputs.xlsm")
    has_bundled = os.path.exists(BUNDLED_TEMPLATE)

    if has_bundled:
        st.success("HFC inputs template: `hfc_inputs.xlsm` (pre-loaded)")
    else:
        st.warning("No bundled template found — please upload one below.")

    template_file = None
    if not has_bundled:
        template_file = st.file_uploader("HFC Inputs Template (.xlsm)", type=["xlsm"], key="upload_template")

    ready = survey_file and (has_bundled or template_file)

    if ready:
        out_name = "hfc_inputs_populated.xlsm"
        if st.button("Run Auto-Populate", key="btn_upload", use_container_width=True, type="primary"):
            with st.spinner("Populating HFC inputs..."):
                try:
                    survey_bytes = survey_file.read()
                    if has_bundled:
                        with open(BUNDLED_TEMPLATE, "rb") as f:
                            template_bytes = f.read()
                    else:
                        template_bytes = template_file.read()
                    output_bytes, results, n_num, n_sel, n_txt = run_population(
                        survey_bytes, template_bytes
                    )
                    show_results(output_bytes, results, n_num, n_sel, n_txt, out_name)
                except Exception as e:
                    st.error(f"Something went wrong: {e}")
    else:
        st.info("Upload the survey form above to enable the run button.")


# ── TAB 2: GOOGLE DRIVE LINK + BOX PATH ──────────────────────
with tab2:
    st.subheader("Google Drive link & Box folder path")
    st.markdown(
        "Paste the Google Drive share link for your survey. "
        "The HFC inputs template is pre-loaded — just choose where to save the output in Box."
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

    # ── Helper: clean and resolve any local path ──────────────
    def resolve_path(raw: str):
        from pathlib import Path
        clean = raw.encode("utf-8", errors="ignore").decode("utf-8")
        clean = clean.replace("\x00", "").replace("​", "").replace("﻿", "")
        clean = "".join(c for c in clean if c.isprintable() or c == "\\")
        clean = clean.strip().strip('"').strip("'").strip()
        clean = clean.replace("/", "\\")
        return Path(clean), clean

    # ── 1. Google Drive survey link ────────────────────────────
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

    # ── 2. HFC inputs template ─────────────────────────────────
    BUNDLED_TEMPLATE = os.path.join(os.path.dirname(__file__), "hfc_inputs.xlsm")
    has_bundled = os.path.exists(BUNDLED_TEMPLATE)

    st.markdown("**2. HFC inputs template**")
    template_source = st.radio(
        "Template source",
        options=["Use pre-loaded template", "Use template from folder path"],
        label_visibility="collapsed",
        horizontal=True,
    )

    template_bytes_cloud = None
    if template_source == "Use pre-loaded template":
        if has_bundled:
            st.success("Using pre-loaded `hfc_inputs.xlsm`")
            with open(BUNDLED_TEMPLATE, "rb") as f:
                template_bytes_cloud = f.read()
        else:
            st.warning("No pre-loaded template found. Switch to folder path option below.")
    else:
        template_path_input = st.text_input(
            "Paste the full file path to your HFC template (.xlsm)",
            placeholder=r"C:\Users\yourname\Box\...\hfc_inputs.xlsm",
            key="template_path_input",
        )
        if template_path_input:
            tp, tp_clean = resolve_path(template_path_input)
            st.caption(f"Looking for: `{tp_clean}`")
            if tp.is_file():
                st.success(f"Template found: `{tp.name}`")
                template_bytes_cloud = tp.read_bytes()
            elif tp.is_dir():
                xlsm_files = sorted(tp.glob("*.xlsm"))
                if xlsm_files:
                    chosen = st.selectbox("Select template", xlsm_files, format_func=lambda f: f.name)
                    template_bytes_cloud = chosen.read_bytes()
                else:
                    st.error("No .xlsm files found in that folder.")
            else:
                st.error("Path not found. Check the path is correct and the file exists.")

    # ── 3. Output location ─────────────────────────────────────
    st.markdown("**3. Where to save the output**")
    output_path_input = st.text_input(
        "Paste a folder path or full file path for the output",
        placeholder=r"C:\Users\yourname\Desktop\DMS App\3_checks\1_inputs",
        label_visibility="collapsed",
        key="output_path_input",
    )

    out_folder = None
    output_name = "hfc_inputs_populated.xlsm"

    if output_path_input:
        op, op_clean = resolve_path(output_path_input)
        st.caption(f"Looking in: `{op_clean}`")

        # If user gave a full .xlsm path, split into folder + filename
        if op_clean.lower().endswith(".xlsm"):
            out_folder = op.parent
            output_name = op.name
        else:
            out_folder = op

        if out_folder and out_folder.exists():
            st.success(f"Output folder found: `{out_folder}`")
        else:
            st.error(
                "Folder not found. Check the path is correct.\n\n"
                "**Tip:** Open the folder in Windows Explorer → click the address bar → copy the path."
            )
            out_folder = None

    output_name = st.text_input("Output filename", value=output_name, key="output_name_cloud")

    # ── Run button ─────────────────────────────────────────────
    st.markdown("")
    ready = gdrive_file_id and template_bytes_cloud and out_folder

    if st.button("Run Auto-Populate", key="btn_cloud", use_container_width=True, type="primary", disabled=not ready):
        with st.spinner("Downloading survey, populating and saving..."):
            try:
                survey_bytes = download_gdrive_file(gdrive_file_id)
                output_bytes, results, n_num, n_sel, n_txt = run_population(
                    survey_bytes, template_bytes_cloud
                )
                out_path = out_folder / output_name
                out_path.write_bytes(output_bytes)
                show_results(output_bytes, results, n_num, n_sel, n_txt, output_name)
                st.info(f"Saved to: `{out_path}`")
            except Exception as e:
                st.error(f"Something went wrong: {e}")

    if not ready:
        missing = []
        if not gdrive_file_id: missing.append("Google Drive survey link")
        if not template_bytes_cloud: missing.append("HFC inputs template")
        if not out_folder: missing.append("output folder path")
        if missing:
            st.info(f"Still needed: {', '.join(missing)}")


# ─────────────────────────────────────────────────────────────
# Footer
# ─────────────────────────────────────────────────────────────
st.markdown("---")
st.caption("Built for IPA Tanzania · Questions? Contact jjairo@poverty-action.org")
