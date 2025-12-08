import streamlit as st
import msal
import requests
import re
import os
from openpyxl import Workbook, load_workbook
from io import BytesIO

# =========================================================
# CONFIG
# =========================================================
CLIENT_ID = st.secrets["client_id"]
CLIENT_SECRET = st.secrets["client_secret"]
TENANT_ID = st.secrets["tenant_id"]
REDIRECT_URI = st.secrets["redirect_uri"]

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["User.Read", "Mail.Read"]    # NO admin approval needed

GRAPH = "https://graph.microsoft.com/v1.0"


# =========================================================
# AUTO-CREATE EXCEL FILES
# =========================================================
def ensure_dsr_file():
    if not os.path.exists("DSR.xlsx"):
        wb = Workbook()
        ws = wb.active
        ws.append(["Project Name", "Resource Name"])  # base headers
        wb.save("DSR.xlsx")


def ensure_wsr_file():
    if not os.path.exists("WSR.xlsx"):
        wb = Workbook()
        ws = wb.active
        ws.append(["Project Name"])  # base header
        wb.save("WSR.xlsx")


# Ensure files exist on startup
ensure_dsr_file()
ensure_wsr_file()


# =========================================================
# TOKEN CACHE HELPERS
# =========================================================
def build_app(cache=None):
    return msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET,
        token_cache=cache,
    )


def load_cache():
    cache = msal.SerializableTokenCache()
    if "token_cache" in st.session_state:
        cache.deserialize(st.session_state["token_cache"])
    return cache


def save_cache(cache):
    if cache.has_state_changed:
        st.session_state["token_cache"] = cache.serialize()


def get_token_silent():
    cache = load_cache()
    app = build_app(cache)
    accts = app.get_accounts()
    if accts:
        result = app.acquire_token_silent(SCOPES, account=accts[0])
        save_cache(cache)
        return result
    return None


# =========================================================
# GRAPH HELPERS
# =========================================================
def graph_get(url, token, params=None):
    headers = {"Authorization": f"Bearer {token}"}
    r = requests.get(url, headers=headers, params=params)
    r.raise_for_status()
    return r.json()


def find_folder_anywhere(token, folder_name):
    """Find folder with this name anywhere in mailbox."""
    url = f"{GRAPH}/me/mailFolders?$top=999"
    headers = {"Authorization": f"Bearer {token}"}

    r = requests.get(url, headers=headers)
    r.raise_for_status()

    for f in r.json().get("value", []):
        if f["displayName"].lower().strip() == folder_name.lower().strip():
            return f["id"]

    return None


def get_messages(token, folder_id, top=50):
    url = f"{GRAPH}/me/mailFolders/{folder_id}/messages"
    params = {
        "$top": top,
        "$select": "subject,body,bodyPreview,receivedDateTime"
    }
    return graph_get(url, token, params).get("value", [])


def get_two_inbox_emails(token):
    url = f"{GRAPH}/me/mailFolders/inbox/messages"
    params = {"$top": 2, "$select": "subject,body,receivedDateTime,from"}
    return graph_get(url, token, params).get("value", [])


# =========================================================
# PARSING
# =========================================================
DSR_RE = re.compile(r"^(?:DSR|Daily Status Report)\s*-\s*(?P<resource>.+?)\s*-\s*(?P<project>.+)$", re.I)
WSR_RE = re.compile(r"^WSR\s*-\s*(?P<project>.+)$", re.I)


def clean_body(body):
    content = body.get("content", "") or ""
    content = re.sub(r"<[^>]+>", "", content)
    return content.strip()


def parse_date(text):
    m = re.search(r"(\d{1,2}\s+[A-Za-z]{3,}\s+\d{4})", text)
    if m:
        return m.group(1)
    return "Unknown"


# =========================================================
# EXCEL UPDATE
# =========================================================
def update_dsr_excel(records):
    wb = load_workbook("DSR.xlsx")
    ws = wb.active

    headers = {ws.cell(row=1, column=i).value: i for i in range(1, ws.max_column + 1)}

    for rec in records:
        project = rec["project"]
        resource = rec["resource"]
        date_label = rec["date"]

        # Create date column if missing
        if date_label not in headers:
            col = ws.max_column + 1
            ws.cell(row=1, column=col, value=date_label)
            headers[date_label] = col

        # Find or create row
        row_idx = None
        for r in range(2, ws.max_row + 1):
            if ws.cell(row=r, column=1).value == project and ws.cell(row=r, column=2).value == resource:
                row_idx = r
                break

        if not row_idx:
            row_idx = ws.max_row + 1
            ws.cell(row=row_idx, column=1, value=project)
            ws.cell(row=row_idx, column=2, value=resource)

        ws.cell(row=row_idx, column=headers[date_label], value="Y")

    wb.save("DSR.xlsx")


def update_wsr_excel(records):
    wb = load_workbook("WSR.xlsx")
    ws = wb.active

    headers = {ws.cell(row=1, column=i).value: i for i in range(1, ws.max_column + 1)}

    for rec in records:
        project = rec["project"]
        week = rec["week"]

        if week not in headers:
            col = ws.max_column + 1
            ws.cell(row=1, column=col, value=week)
            headers[week] = col

        row_idx = None
        for r in range(2, ws.max_row + 1):
            if ws.cell(row=r, column=1).value == project:
                row_idx = r
                break

        if not row_idx:
            row_idx = ws.max_row + 1
            ws.cell(row=row_idx, column=1, value=project)

        ws.cell(row=row_idx, column=headers[week], value="Y")

    wb.save("WSR.xlsx")


# =========================================================
# STREAMLIT UI
# =========================================================
st.title("📩 DSR / WSR Automation — Auto-create Excel + Local Update")

# --- OAuth redirect handling ---
params = st.query_params

if "code" in params:
    code = params["code"]
    cache = load_cache()
    app = build_app(cache)

    result = app.acquire_token_by_authorization_code(
        code, SCOPES, redirect_uri=REDIRECT_URI
    )
    save_cache(cache)

    st.query_params.clear()

    if "access_token" not in result:
        st.error("Authentication Failed.")
    else:
        st.success("Authenticated Successfully!")


# --- Silent login ---
token_result = get_token_silent()

if not token_result:
    st.warning("You are not signed in.")

    if st.button("Sign in with Microsoft"):
        cache = load_cache()
        app = build_app(cache)
        auth_url = app.get_authorization_request_url(SCOPES, redirect_uri=REDIRECT_URI)
        save_cache(cache)
        st.markdown(f"[Click here to sign in]({auth_url})")

    st.stop()

token = token_result["access_token"]
st.success("Logged in ✔️")


# =========================================================
# SHOW 2 INBOX EMAILS
# =========================================================
st.header("🔍 View sample inbox emails")

if st.button("Show 2 Inbox Emails"):
    emails = get_two_inbox_emails(token)
    for e in emails:
        st.write(f"**Subject:** {e['subject']}")
        st.write("Preview:", e["body"])
        st.write("From:", e["from"]["emailAddress"]["address"])
        st.write("Received:", e["receivedDateTime"])
        st.write("---")


# =========================================================
# PROCESS DSR
# =========================================================
st.header("📘 Process DSR Folder")

if st.button("Process DSR"):
    folder_id = find_folder_anywhere(token, "DSR")

    if not folder_id:
        st.error("DSR folder was not found anywhere in mailbox.")
        st.stop()

    msgs = get_messages(token, folder_id, top=50)

    records = []
    for m in msgs:
        subj = m.get("subject", "")
        msub = DSR_RE.match(subj)
        if not msub:
            continue

        resource = msub.group("resource").strip()
        project = msub.group("project").strip()
        body = clean_body(m.get("body", {}))
        date_label = parse_date(body)

        records.append({
            "project": project,
            "resource": resource,
            "date": date_label
        })

    update_dsr_excel(records)
    st.success(f"Updated DSR.xlsx with {len(records)} entries.")
    st.write(records)


# =========================================================
# PROCESS WSR
# =========================================================
st.header("📗 Process WSR Folder")

if st.button("Process WSR"):
    folder_id = find_folder_anywhere(token, "WSR")

    if not folder_id:
        st.error("WSR folder was not found anywhere in mailbox.")
        st.stop()

    msgs = get_messages(token, folder_id, top=50)

    records = []
    for m in msgs:
        subj = m.get("subject", "")
        msub = WSR_RE.match(subj)
        if not msub:
            continue

        project = msub.group("project").strip()
        body = clean_body(m.get("body", {}))
        date_label = parse_date(body)

        records.append({
            "project": project,
            "week": date_label
        })

    update_wsr_excel(records)
    st.success(f"Updated WSR.xlsx with {len(records)} entries.")
    st.write(records)
