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


# ── TAB 2: GOOGLE DRIVE + BOX ────────────────────────────────
with tab2:
    st.subheader("Connect to Google Drive & Box")
    st.markdown(
        "Paste the file IDs from Google Drive and Box. "
        "The tool will download, populate, and upload the result back to Box automatically."
    )

    with st.expander("ℹ️ How to find file IDs"):
        st.markdown("""
**Google Drive file ID** — open the file in Drive and copy from the URL:
```
https://drive.google.com/file/d/**<FILE_ID>**/view
```

**Box file ID** — open the file in Box and copy the number from the URL:
```
https://app.box.com/file/**<FILE_ID>**
```

**Box folder ID** — open the destination folder and copy from the URL:
```
https://app.box.com/folder/**<FOLDER_ID>**
```
        """)

    gdrive_id      = st.text_input("Google Drive file ID (XLSForm survey)")
    box_template_id = st.text_input("Box file ID (HFC inputs template)")
    box_folder_id  = st.text_input("Box folder ID (where to save the output)")
    output_name    = st.text_input("Output filename", value="hfc_inputs_populated.xlsm")

    st.markdown("---")
    st.markdown("**Credentials** — ask your IT admin or project lead to fill these in once.")

    with st.expander("Google Drive credentials"):
        gdrive_creds = st.text_area(
            "Paste your Google service account JSON key here",
            height=120,
            placeholder='{"type": "service_account", "project_id": "...", ...}',
        )

    with st.expander("Box credentials"):
        box_client_id     = st.text_input("Box Client ID",     type="password")
        box_client_secret = st.text_input("Box Client Secret", type="password")
        box_access_token  = st.text_input("Box Access Token",  type="password",
                                           help="Generate from the Box developer console or run box_first_time_auth() from the Python script.")

    all_filled = all([gdrive_id, box_template_id, box_folder_id, gdrive_creds, box_client_id, box_client_secret, box_access_token])

    if st.button("Run Auto-Populate", key="btn_cloud", use_container_width=True, type="primary", disabled=not all_filled):
        with st.spinner("Downloading files and populating HFC inputs..."):
            try:
                import json
                from googleapiclient.discovery import build
                from google.oauth2 import service_account
                from googleapiclient.http import MediaIoBaseDownload
                from boxsdk import OAuth2, Client

                # Google Drive download
                creds_info = json.loads(gdrive_creds)
                creds = service_account.Credentials.from_service_account_info(
                    creds_info, scopes=["https://www.googleapis.com/auth/drive.readonly"]
                )
                service = build("drive", "v3", credentials=creds)
                request = service.files().get_media(fileId=gdrive_id)
                survey_buf = io.BytesIO()
                downloader = MediaIoBaseDownload(survey_buf, request)
                done = False
                while not done:
                    _, done = downloader.next_chunk()
                survey_bytes = survey_buf.getvalue()

                # Box download
                oauth = OAuth2(
                    client_id=box_client_id,
                    client_secret=box_client_secret,
                    access_token=box_access_token,
                )
                client = Client(oauth)
                template_buf = io.BytesIO()
                client.file(box_template_id).download_to(template_buf)
                template_bytes = template_buf.getvalue()

                # Populate
                output_bytes, results, n_num, n_sel, n_txt = run_population(survey_bytes, template_bytes)

                # Upload to Box
                folder = client.folder(box_folder_id)
                existing_items = {item.name: item for item in folder.get_items()}
                upload_buf = io.BytesIO(output_bytes)
                if output_name in existing_items:
                    existing_items[output_name].update_contents_with_stream(upload_buf)
                else:
                    folder.upload_stream(upload_buf, output_name)

                show_results(output_bytes, results, n_num, n_sel, n_txt, output_name)
                st.info(f"File also saved to Box as **{output_name}**")

            except Exception as e:
                st.error(f"Something went wrong: {e}")

    if not all_filled:
        st.info("Fill in all fields above to enable the run button.")


# ─────────────────────────────────────────────────────────────
# Footer
# ─────────────────────────────────────────────────────────────
st.markdown("---")
st.caption("Built for IPA Tanzania · Questions? Contact jjairo@poverty-action.org")
