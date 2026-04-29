"""
HFC Inputs Auto-Populator
=========================
Pulls a SurveyCTO XLSForm from Google Drive and an HFC inputs template
from Box, populates all six DMS sheets (other specify, outliers,
constraints, logic, enumstats, text audits), and uploads the result
back to Box.

SETUP — run once before first use:
    pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib boxsdk openpyxl

HOW TO USE:
    1. Fill in the CONFIG section below with your file IDs and credentials.
    2. Run:  python populate_hfc_inputs.py
    3. The populated file will appear in the same Box folder as your template.

GETTING FILE IDs:
    Google Drive: open the file → copy the long ID from the URL
        https://drive.google.com/file/d/<FILE_ID>/view
    Box: open the file → copy the number at the end of the URL
        https://app.box.com/file/<FILE_ID>

CREDENTIALS:
    Google Drive: download credentials.json from Google Cloud Console
        (APIs & Services → Credentials → OAuth 2.0 Client → Download)
    Box: create an app at https://developer.box.com
        (Custom App → User Authentication OAuth 2.0 → copy Client ID & Secret)
"""

import io
import os
import re
import tempfile

# ============================================================
# CONFIG — edit these values
# ============================================================

# Google Drive: file ID of your XLSForm survey (.xlsx)
GDRIVE_SURVEY_FILE_ID = "YOUR_GOOGLE_DRIVE_FILE_ID_HERE"

# Box: file ID of your HFC inputs template (.xlsm)
BOX_TEMPLATE_FILE_ID = "YOUR_BOX_TEMPLATE_FILE_ID_HERE"

# Box: folder ID where the populated output file should be saved
# (find it in the Box URL when you open the folder)
BOX_OUTPUT_FOLDER_ID = "YOUR_BOX_OUTPUT_FOLDER_ID_HERE"

# Output filename saved to Box
OUTPUT_FILENAME = "hfc_inputs_populated.xlsm"

# Path to your Google OAuth credentials JSON file
GOOGLE_CREDENTIALS_FILE = "google_credentials.json"

# Box credentials
BOX_CLIENT_ID     = "YOUR_BOX_CLIENT_ID"
BOX_CLIENT_SECRET = "YOUR_BOX_CLIENT_SECRET"

# ============================================================
# STEP 1: Authenticate and download survey from Google Drive
# ============================================================

def get_google_drive_service():
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build

    SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
    token_file = "google_token.json"
    creds = None

    if os.path.exists(token_file):
        creds = Credentials.from_authorized_user_file(token_file, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(GOOGLE_CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_file, "w") as f:
            f.write(creds.to_json())

    return build("drive", "v3", credentials=creds)


def download_from_google_drive(file_id: str, dest_path: str):
    from googleapiclient.http import MediaIoBaseDownload

    service = get_google_drive_service()
    request = service.files().get_media(fileId=file_id)
    with open(dest_path, "wb") as fh:
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
    print(f"  Downloaded survey from Google Drive → {dest_path}")


# ============================================================
# STEP 2: Authenticate and download template from Box
# ============================================================

def get_box_client():
    from boxsdk import OAuth2, Client

    auth = OAuth2(
        client_id=BOX_CLIENT_ID,
        client_secret=BOX_CLIENT_SECRET,
        store_tokens=_store_box_tokens,
        access_token=_load_box_token("access"),
        refresh_token=_load_box_token("refresh"),
    )
    return Client(auth)


def _store_box_tokens(access_token, refresh_token):
    with open("box_token.txt", "w") as f:
        f.write(f"{access_token}\n{refresh_token}")


def _load_box_token(kind: str):
    if not os.path.exists("box_token.txt"):
        return None
    with open("box_token.txt") as f:
        lines = f.read().splitlines()
    return lines[0] if kind == "access" and lines else (lines[1] if kind == "refresh" and len(lines) > 1 else None)


def box_first_time_auth():
    """
    Run this once to get your first Box tokens.
    Opens a browser window for you to log in.
    """
    from boxsdk import OAuth2, Client

    auth_url_visited = {}

    def _open_browser(url):
        import webbrowser
        print(f"\n  Opening Box login in your browser...")
        webbrowser.open(url)

    oauth = OAuth2(
        client_id=BOX_CLIENT_ID,
        client_secret=BOX_CLIENT_SECRET,
        store_tokens=_store_box_tokens,
    )
    auth_url, csrf_token = oauth.get_authorization_url("http://localhost")
    _open_browser(auth_url)
    redirect_url = input("  Paste the full redirect URL from your browser here: ").strip()
    code = redirect_url.split("code=")[1].split("&")[0]
    access_token, refresh_token = oauth.authenticate(code)
    print("  Box authentication successful. Tokens saved to box_token.txt")
    return Client(oauth)


def download_from_box(file_id: str, dest_path: str):
    client = get_box_client()
    box_file = client.file(file_id)
    with open(dest_path, "wb") as f:
        box_file.download_to(f)
    print(f"  Downloaded template from Box → {dest_path}")


def upload_to_box(local_path: str, folder_id: str, filename: str):
    client = get_box_client()
    folder = client.folder(folder_id)

    # If file already exists in folder, update it; otherwise upload new
    existing = {item.name: item for item in folder.get_items()}
    if filename in existing:
        existing[filename].update_contents(local_path)
        print(f"  Updated existing file in Box: {filename}")
    else:
        folder.upload(local_path, filename)
        print(f"  Uploaded new file to Box: {filename}")


# ============================================================
# STEP 3: Populate the HFC inputs file
# ============================================================

def clean_label(s):
    if not s:
        return ""
    return re.sub(r"<[^>]+>", "", str(s)).strip()


def build_label_lookup(survey_path: str) -> dict:
    from openpyxl import load_workbook
    wb = load_workbook(survey_path, read_only=True, data_only=True)
    ws = wb["survey"]
    labels = {}
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue
        name = row[1] if len(row) > 1 else None
        label = row[2] if len(row) > 2 else None
        if name:
            labels[str(name)] = clean_label(label)
    wb.close()
    return labels


def classify_survey_variables(survey_path: str):
    from openpyxl import load_workbook
    wb = load_workbook(survey_path, read_only=True, data_only=True)
    ws = wb["survey"]

    SKIP_TYPES = {
        "begin group", "end group", "begin repeat", "end repeat",
        "note", "calculate", "start", "end", "deviceid",
        "phonenumber", "username", "caseid", "enumerator",
    }

    groups = []
    numerics, selects, texts = [], [], []

    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue
        def g(idx):
            return row[idx] if len(row) > idx and row[idx] else ""
        t, n, l = g(0), g(1), clean_label(g(2))
        if not n:
            continue
        base = t.split()[0] if t else ""
        if base in ("begin", "end") and "group" in t:
            if "begin" in t:
                groups.append({"name": n, "label": l})
        elif base in SKIP_TYPES:
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
    vals = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[col]:
            vals.add(str(row[col]).strip())
    return vals


def populate_other_specify(ws, selects, texts):
    text_names = {v["name"] for v in texts}
    select_map = {v["name"]: v for v in selects}
    existing = existing_vars(ws, col=0)

    added = 0
    for v in selects:
        name = v["name"]
        # Look for a matching child variable: name + "_oth" or name + "_oth_r"
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
    existing = existing_vars(ws, col=0)
    added = 0
    for v in numerics:
        if v["name"] not in existing:
            ws.append((v["name"], v["label"], "enum", "sd", 3, None, None, None, "id member_name enumerator_id"))
            added += 1
    return added


def populate_constraints(ws, numerics):
    """
    Adds sensible hard/soft min/max defaults by variable name pattern.
    You should review and adjust these for your study context.
    """
    existing = existing_vars(ws, col=0)

    def bounds(name):
        n = name.lower()
        if "age" in n:
            return (15, 18, 70, 100)
        if "nbr" in n or "_num" in n or "count" in n or "hh_own" in n:
            return (0, 0, 20, 100)
        if "year" in n and "begin" in n:
            return (1970, 1990, 2025, 2026)
        if "area" in n or "plot" in n:
            return (0, 0, 20, 50)
        if "hrs" in n or "hour" in n or "time" in n:
            return (0, 0, 14, 24)
        if "week" in n and "hrs" not in n:
            return (0, 0, 80, 168)
        if "last1m" in n or "lastmonth" in n or "last_mont" in n:
            return (0, 10000, 2000000, 15000000)
        if "last12m" in n or "last12" in n:
            return (0, 100000, 10000000, 50000000)
        if "borrowed" in n or "loan" in n or "borrow" in n:
            return (0, 50000, 5000000, 20000000)
        if "saved" in n or "saving" in n:
            return (0, 0, 5000000, 50000000)
        if "val" in n or "value" in n or "sales" in n or "rev" in n or "cost" in n or "earn" in n or "spend" in n or "spent" in n or "expense" in n:
            return (0, 0, 1000000, 10000000)
        if "accessed" in n:
            return (0, 0, 26, 26)
        # generic numeric
        return (0, 0, None, None)

    added = 0
    for v in numerics:
        if v["name"] not in existing:
            hard_min, soft_min, soft_max, hard_max = bounds(v["name"])
            ws.append((v["name"], v["label"], hard_min, soft_min, soft_max, hard_max, None, None, None))
            added += 1
    return added


def populate_logic(ws, numerics):
    """
    Adds cross-variable consistency checks for common patterns.
    """
    existing = existing_vars(ws, col=0)
    added = 0
    num_names = {v["name"]: v for v in numerics}

    # Annual >= monthly pairs
    for v in numerics:
        n = v["name"]
        if "last12m" in n:
            monthly = n.replace("last12m", "last1m")
            if monthly in num_names and n not in existing:
                ws.append((
                    n, v["label"],
                    f"{n} >= {monthly}",
                    f"!missing({monthly})",
                    None, None, "id member_name enumerator_id"
                ))
                existing.add(n)
                added += 1

    # Time use: paid + unpaid <= 24 hours
    if "s6_time_paid_work" in num_names and "s6_time_paid_work" not in existing:
        ws.append((
            "s6_time_paid_work",
            num_names["s6_time_paid_work"]["label"],
            "s6_time_paid_work + s6_time_unpaid_act <= 24",
            "!missing(s6_time_unpaid_act)",
            None, None, None
        ))
        existing.add("s6_time_paid_work")
        added += 1

    # Weekly hours cap
    if "week_hrs_spend" in num_names and "week_hrs_spend" not in existing:
        ws.append(("week_hrs_spend", num_names["week_hrs_spend"]["label"], "week_hrs_spend <= 168", None, None, None, None))
        existing.add("week_hrs_spend")
        added += 1

    # Business start year
    if "year_begin" in num_names and "year_begin" not in existing:
        ws.append(("year_begin", num_names["year_begin"]["label"], "year_begin >= 1970 & year_begin <= 2026", None, None, None, None))
        existing.add("year_begin")
        added += 1

    return added


def populate_enumstats(ws, numerics):
    existing = existing_vars(ws, col=0)
    added = 0
    for v in numerics:
        if v["name"] not in existing:
            n = v["name"].lower()
            # Use combine=yes for repeat variables (ending in _r or _r??)
            combine = "yes" if re.search(r"_r\??$|_r\d", n) else None
            ws.append((v["name"], v["label"], "yes", None, "number", "yes", "yes", "yes", combine, None))
            added += 1
    return added


def populate_text_audits(ws, groups):
    existing = existing_vars(ws, col=0)
    added = 0
    for g in groups:
        if g["name"] not in existing and g["name"].startswith("section"):
            ws.append((g["name"], None, None, g["label"]))
            added += 1
    return added


def populate_hfc_inputs(survey_path: str, template_path: str, output_path: str):
    from openpyxl import load_workbook

    print("\n  Classifying survey variables...")
    numerics, selects, texts, groups = classify_survey_variables(survey_path)
    print(f"    Found {len(numerics)} numeric, {len(selects)} select, {len(texts)} text variables")

    print("  Loading HFC template...")
    wb = load_workbook(template_path, keep_vba=True)

    results = {}
    results["other specify"] = populate_other_specify(wb["other specify"], selects, texts)
    results["outliers"]      = populate_outliers(wb["outliers"], numerics)
    results["constraints"]   = populate_constraints(wb["constraints"], numerics)
    results["logic"]         = populate_logic(wb["logic"], numerics)
    results["enumstats"]     = populate_enumstats(wb["enumstats"], numerics)
    results["text audits"]   = populate_text_audits(wb["text audits"], groups)

    wb.save(output_path)
    print(f"\n  Populated file saved → {output_path}")

    for sheet, count in results.items():
        print(f"    {sheet:20} +{count} rows added")

    return results


# ============================================================
# MAIN — orchestrates everything
# ============================================================

def main():
    print("=" * 60)
    print("HFC Inputs Auto-Populator")
    print("=" * 60)

    with tempfile.TemporaryDirectory() as tmpdir:
        survey_path   = os.path.join(tmpdir, "survey.xlsx")
        template_path = os.path.join(tmpdir, "hfc_template.xlsm")
        output_path   = os.path.join(tmpdir, OUTPUT_FILENAME)

        print("\nStep 1: Downloading survey from Google Drive...")
        download_from_google_drive(GDRIVE_SURVEY_FILE_ID, survey_path)

        print("\nStep 2: Downloading HFC template from Box...")
        download_from_box(BOX_TEMPLATE_FILE_ID, template_path)

        print("\nStep 3: Populating HFC inputs...")
        populate_hfc_inputs(survey_path, template_path, output_path)

        print("\nStep 4: Uploading populated file to Box...")
        upload_to_box(output_path, BOX_OUTPUT_FOLDER_ID, OUTPUT_FILENAME)

    print("\nDone! Your populated HFC inputs file is in Box.")
    print("=" * 60)


if __name__ == "__main__":
    # First time Box setup: uncomment the line below, run once, then comment it out again
    # box_first_time_auth(); exit()

    main()
