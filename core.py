"""
Core extraction and upload logic for Plankton Settlement Reports.
No GUI dependencies — can be used from Kivy, Tkinter, CLI, etc.
"""

import os
import re

import pdfplumber
import gspread
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2 import service_account

# ---------------------------------------------------------------------------
# Google API config
# ---------------------------------------------------------------------------
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# ---------------------------------------------------------------------------
# Amount regex
# ---------------------------------------------------------------------------
RE_AMOUNT = re.compile(
    r"R\s+"
    r"(\(?)"
    r"([\d,]+\.\d+)"
    r"\)?"
)
RE_AMOUNT_DASH = re.compile(r"R\s+-\s*\)?$")

RE_SHEET_ID = re.compile(r"/spreadsheets/d/([a-zA-Z0-9-_]+)")

# ---------------------------------------------------------------------------
# Auto-calc columns — never write to these
# ---------------------------------------------------------------------------
AUTO_CALC_HEADERS = {
    "PROFIT/LOSS (CURRENT & FUTURE)",
    "PROFIT / LOSS",
    "SUB-TOTAL",
    "SALES COMS",
    "AGENT",
}

# ---------------------------------------------------------------------------
# Keyword-based mapping: PDF line item → sheet header
# ---------------------------------------------------------------------------
KEYWORD_RULES = [
    (["EVENT DATE"],                    "Date",          1),
    (["EVENT NAME"],                    "Event Name",    1),
    (["SERVICE FEE", "PLANKTON"],       "Service Fee",   1),
    (["MERCHANT"],                      "Merchant Fee",  1),
    (["MANAGER FEE 2"],                 "Manager",       2),
    (["MANAGER"],                       "Manager",       1),
    (["LABOUR"],                        "Labour",        1),
    (["SCANNER"],                       "Scanner",       1),
    (["POS DEVICE"],                    "POS Device",    1),
    (["PRINTER"],                       "Printer",       1),
    (["TILL ROLL"],                     "Printer",       1),
    (["CARD TRANSACTION"],              "Card Machine",  1),
    (["DATA BU"],                       "Card Machine",  1),
    (["AA RATE"],                       "AA Rates",      1),
    (["BANKING FEE"],                   "Banking Fee",   1),
    (["TOTAL TICKET SALES INCOME"],     "Turnover",      1),
    (["TICKET SALES"],                  "Tickets",       1),
    (["S&T"],                           "Other",         1),
    (["FOOD"],                          "Other",         1),
    (["ACCOMMODATION"],                 "Other",         1),
    (["SOFTWARE"],                      "Other",         1),
    (["SMS"],                           "Other",         1),
    (["EMAILER"],                       "Other",         1),
    (["BANNER"],                        "Other",         1),
    (["DISCOUNT"],                      "Other",         1),
]

# ---------------------------------------------------------------------------
# PDF line-item labels
# ---------------------------------------------------------------------------
STARTSWITH_LABELS = [
    "PLANKTON SERVICE FEE",
    "BANKING FEE ON PRESALES ONLY",
    "MANAGER FEE",
    "LABOUR FEE",
    "TICKET SCANNER",
    "POS DEVICE AND TICKET PRINTER",
    "DATA BUDLE",
    "Till ROLLS - RECON AMOUNT",
    "AA RATES - 65KM ONE WAY",
    "S&Ts - FOOD FOR PLANKTON STAFF",
    "ACCOMMODATION",
    "SOFTWARE DEVELOPMENT",
    "SMS COSTS",
    "EMAILER COSTS",
    "MAIN FEATURED BANNER",
]

CONTAINS_LABELS = [
    "CARD TRANSACTIONS",
    "DISCOUNT",
    "TOTAL TICKET SALES INCOME",
]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def col_letter(idx: int) -> str:
    """Convert 0-based column index to sheet letter (0->A, 25->Z, 26->AA)."""
    result = ""
    while True:
        result = chr(65 + idx % 26) + result
        idx = idx // 26 - 1
        if idx < 0:
            break
    return result


def _parse_amount(text: str) -> float | None:
    if RE_AMOUNT_DASH.search(text):
        return 0.0
    matches = list(RE_AMOUNT.finditer(text))
    if not matches:
        return None
    m = matches[-1]
    value = float(m.group(2).replace(",", ""))
    if m.group(1) == "(":
        value = -value
    return value


def _match_header(pdf_key: str) -> tuple[str, int] | None:
    upper = pdf_key.upper()
    for keywords, header, section in KEYWORD_RULES:
        if all(kw.upper() in upper for kw in keywords):
            return header, section
    return None


# ---------------------------------------------------------------------------
# PDF extraction
# ---------------------------------------------------------------------------
def extract_from_pdf(pdf_path: str, log=print) -> dict:
    """Parse a Plankton Template 2 settlement PDF -> dict of label->value."""
    data = {}
    manager_fee_count = 0

    with pdfplumber.open(pdf_path) as pdf:
        if len(pdf.pages) < 2:
            raise ValueError("PDF has fewer than 2 pages - not a Template 2 report.")

        p1_text = pdf.pages[0].extract_text() or ""

        m = re.search(r"EVENT NAME:\s*(.+?)\s*EVENT DATE:", p1_text)
        if m:
            data["EVENT NAME"] = m.group(1).strip()
            log(f"  Event Name: {data['EVENT NAME']}")

        m = re.search(r"EVENT DATE:\s*(.+)$", p1_text, re.MULTILINE)
        if m:
            raw_date = m.group(1).strip()
            data["EVENT DATE"] = re.sub(r"\s+", "", raw_date)
            log(f"  Event Date: {data['EVENT DATE']}")

        p2_text = pdf.pages[1].extract_text() or ""
        if "FINANCIAL RECON" not in p2_text.upper():
            raise ValueError(
                "Page 2 does not contain 'FINANCIAL RECON' - "
                "this may not be a Template 2 settlement report."
            )

        lines = p2_text.splitlines()
        log(f"  Parsing {len(lines)} lines on page 2...")

        for line in lines:
            stripped = line.strip()
            if not stripped:
                continue

            for label in STARTSWITH_LABELS:
                if stripped.upper().startswith(label.upper()):
                    amount = _parse_amount(stripped)
                    if amount is None:
                        continue
                    key = label
                    if label == "MANAGER FEE":
                        manager_fee_count += 1
                        if manager_fee_count == 2:
                            key = "MANAGER FEE 2"
                    data[key] = amount
                    log(f"  {key}: {amount}")
                    break
            else:
                for label in CONTAINS_LABELS:
                    if label.upper() in stripped.upper():
                        amount = _parse_amount(stripped)
                        if amount is None:
                            continue
                        data[label] = amount
                        log(f"  {label}: {amount}")
                        break

    return data


# ---------------------------------------------------------------------------
# Row builder
# ---------------------------------------------------------------------------
def build_row(data: dict, headers: list[str], log=print) -> tuple[list, list, list]:
    """Map PDF data to sheet columns. Returns (row, matched, unmatched)."""
    row = [""] * len(headers)

    header_indices = {}
    for idx, h in enumerate(headers):
        key = h.strip().upper()
        if key and key not in AUTO_CALC_HEADERS:
            header_indices.setdefault(key, []).append(idx)

    matched = []
    unmatched = []

    for pdf_key, value in data.items():
        result = _match_header(pdf_key)
        if result is None:
            unmatched.append(f"{pdf_key} (no keyword match)")
            continue

        sheet_header, section = result
        lookup = sheet_header.strip().upper()

        if lookup in AUTO_CALC_HEADERS:
            log(f"  SKIP (auto-calc): {pdf_key} -> '{sheet_header}'")
            continue

        if lookup not in header_indices:
            unmatched.append(f"{pdf_key} -> '{sheet_header}' (header not found)")
            continue

        indices = header_indices[lookup]
        target_idx = section - 1

        if lookup == "OTHER":
            if target_idx < len(indices):
                col_idx = indices[target_idx]
                existing = row[col_idx]
                if existing == "":
                    row[col_idx] = value
                else:
                    row[col_idx] = existing + value
                matched.append(f"{pdf_key} -> '{sheet_header}' col {col_letter(col_idx)} (summed)")
            else:
                unmatched.append(f"{pdf_key} -> '{sheet_header}' (no section {section} column)")
        else:
            if target_idx < len(indices):
                col_idx = indices[target_idx]
                row[col_idx] = value
                matched.append(f"{pdf_key} -> '{sheet_header}' col {col_letter(col_idx)}")
            else:
                unmatched.append(f"{pdf_key} -> '{sheet_header}' (no section {section} column)")

    return row, matched, unmatched


# ---------------------------------------------------------------------------
# Google Sheets auth
# ---------------------------------------------------------------------------
def get_gspread_client(app_dir: str, log=print) -> gspread.Client:
    """
    Authenticate and return a gspread client.
    Tries service_account.json first (headless/iOS), falls back to OAuth.
    """
    sa_path = os.path.join(app_dir, "service_account.json")
    client_secret_path = os.path.join(app_dir, "client_secret.json")
    token_path = os.path.join(app_dir, "token.json")

    # Prefer service account (works on iOS without browser)
    if os.path.exists(sa_path):
        creds = service_account.Credentials.from_service_account_file(
            sa_path, scopes=SCOPES
        )
        log("Authenticated via service account.")
        return gspread.authorize(creds)

    # Fall back to OAuth (desktop only)
    creds = None
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            from google.auth.transport.requests import Request
            creds.refresh(Request())
            log("Refreshed existing token.")
        else:
            if not os.path.exists(client_secret_path):
                raise FileNotFoundError(
                    "Missing credentials.\n\n"
                    "Place one of these next to the app:\n"
                    "- service_account.json (recommended for iOS)\n"
                    "- client_secret.json (desktop OAuth)"
                )
            flow = InstalledAppFlow.from_client_secrets_file(
                client_secret_path, SCOPES
            )
            creds = flow.run_local_server(port=0)
            log("Authenticated via browser.")

        with open(token_path, "w") as f:
            f.write(creds.to_json())
        log("Token saved.")

    return gspread.authorize(creds)


# ---------------------------------------------------------------------------
# Upload
# ---------------------------------------------------------------------------
def upload_to_sheet(sheet_id: str, data: dict, app_dir: str, log=print) -> tuple[list, list]:
    """
    Upload extracted data to Google Sheet.
    Returns (matched, unmatched) lists.
    """
    gc = get_gspread_client(app_dir, log=log)

    spreadsheet = gc.open_by_key(sheet_id)
    sheet = spreadsheet.sheet1
    headers = sheet.row_values(7)

    if not headers:
        raise ValueError("Sheet has no headers in row 7.")

    log(f"Found {len(headers)} headers in row 7")

    row, matched, unmatched = build_row(data, headers, log=log)

    # Find next empty row
    all_values = sheet.col_values(1)
    next_row = max(len(all_values) + 1, 8)
    log(f"Writing to row {next_row}...")

    # Batch update non-empty cells
    batch = []
    for col_idx, val in enumerate(row):
        if val != "":
            cell_ref = f"{col_letter(col_idx)}{next_row}"
            batch.append({"range": cell_ref, "values": [[val]]})

    if batch:
        sheet.batch_update(batch, value_input_option="USER_ENTERED")

    return matched, unmatched
