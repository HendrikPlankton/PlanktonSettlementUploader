"""
Plankton Event Settlement Report → Google Sheets Uploader

Standalone Tkinter GUI that:
1. Parses Plankton event settlement report PDFs (Template 2)
2. Extracts event header info and financial recon line items
3. Uploads extracted data to an existing Google Sheet via OAuth2

Uses keyword-based fuzzy matching to map PDF line items to sheet headers.
"""

import os
import re
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

import pdfplumber
import gspread
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

# ---------------------------------------------------------------------------
# Google API config
# ---------------------------------------------------------------------------
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CLIENT_SECRET_PATH = os.path.join(SCRIPT_DIR, "client_secret.json")
TOKEN_PATH = os.path.join(SCRIPT_DIR, "token.json")

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

# Sheet URL → ID
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
#
# Each rule is (keywords_to_match, sheet_header, section).
# - keywords: if ALL keywords appear in the PDF label, it's a match
# - section: 1 = first occurrence of that header, 2 = second occurrence
#   (the sheet has two groups of Manager/Labour/Scanner/etc.)
# - Rules are checked top-to-bottom; first match wins.
# ---------------------------------------------------------------------------
KEYWORD_RULES = [
    # Header fields
    (["EVENT DATE"],                    "Date",          1),
    (["EVENT NAME"],                    "Event Name",    1),

    # Section 1: Service fee line items
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

    # Totals
    (["TOTAL TICKET SALES INCOME"],     "Turnover",      1),
    (["TICKET SALES"],                  "Tickets",       1),

    # Catch-all for miscellaneous items → "Other" (section 1, summed)
    (["S&T"],                           "Other",         1),
    (["FOOD"],                          "Other",         1),
    (["ACCOMMODATION"],                 "Other",         1),
    (["SOFTWARE"],                      "Other",         1),
    (["SMS"],                           "Other",         1),
    (["EMAILER"],                       "Other",         1),
    (["BANNER"],                        "Other",         1),
    (["DISCOUNT"],                      "Other",         1),
]


def _match_header(pdf_key: str) -> tuple[str, int] | None:
    """Return (sheet_header, section) for a PDF key, or None."""
    upper = pdf_key.upper()
    for keywords, header, section in KEYWORD_RULES:
        if all(kw.upper() in upper for kw in keywords):
            return header, section
    return None


# ---------------------------------------------------------------------------
# PDF line-item extraction
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


def _col_letter(idx: int) -> str:
    """Convert 0-based column index to sheet letter (0→A, 25→Z, 26→AA)."""
    result = ""
    while True:
        result = chr(65 + idx % 26) + result
        idx = idx // 26 - 1
        if idx < 0:
            break
    return result


def _parse_amount(text: str) -> float | None:
    """Return the *rightmost* R-amount on `text`, or 0.0 for 'R -'."""
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


def extract_from_pdf(pdf_path: str, log=print) -> dict:
    """Parse a Plankton Template 2 settlement PDF → dict of label→value."""
    data = {}
    manager_fee_count = 0

    with pdfplumber.open(pdf_path) as pdf:
        if len(pdf.pages) < 2:
            raise ValueError("PDF has fewer than 2 pages — not a Template 2 report.")

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
                "Page 2 does not contain 'FINANCIAL RECON' — "
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


def get_gspread_client(log=print) -> gspread.Client:
    """Authenticate via OAuth2 and return an authorized gspread client."""
    creds = None

    if os.path.exists(TOKEN_PATH):
        creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            from google.auth.transport.requests import Request
            creds.refresh(Request())
            log("Refreshed existing token.")
        else:
            if not os.path.exists(CLIENT_SECRET_PATH):
                raise FileNotFoundError(
                    f"Missing '{CLIENT_SECRET_PATH}'.\n\n"
                    "To fix this:\n"
                    "1. Go to https://console.cloud.google.com/apis/credentials\n"
                    "2. Create an OAuth 2.0 Client ID (Desktop app)\n"
                    "3. Download the JSON and save it as 'client_secret.json'\n"
                    "   next to this script."
                )
            flow = InstalledAppFlow.from_client_secrets_file(
                CLIENT_SECRET_PATH, SCOPES
            )
            creds = flow.run_local_server(port=0)
            log("Authenticated via browser.")

        with open(TOKEN_PATH, "w") as f:
            f.write(creds.to_json())
        log("Token saved for future use.")

    return gspread.authorize(creds)


# ---------------------------------------------------------------------------
# Map extracted data → row using keyword rules
# ---------------------------------------------------------------------------
def build_row(data: dict, headers: list[str], log=print) -> tuple[list, list, list]:
    """
    Map PDF data to sheet columns using keyword-based fuzzy matching.
    Returns (row_values, matched_info, unmatched_info).
    """
    row = [""] * len(headers)

    # Build header lookup: upper(name) → [col_indices...]
    header_indices = {}
    for idx, h in enumerate(headers):
        key = h.strip().upper()
        if key and key not in AUTO_CALC_HEADERS:
            header_indices.setdefault(key, []).append(idx)

    matched = []
    unmatched = []
    # Track usage per header to handle duplicates (section 1 vs 2)
    header_use = {}  # upper(header) → next_section_index (0-based)

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
        # section is 1-based; convert to 0-based index into indices list
        target_idx = section - 1

        if lookup == "OTHER":
            # "Other" accumulates: sum all values into the target column
            if target_idx < len(indices):
                col_idx = indices[target_idx]
                existing = row[col_idx]
                if existing == "":
                    row[col_idx] = value
                else:
                    row[col_idx] = existing + value
                matched.append(f"{pdf_key} -> '{sheet_header}' col {_col_letter(col_idx)} (summed)")
            else:
                unmatched.append(f"{pdf_key} -> '{sheet_header}' (no section {section} column)")
        else:
            if target_idx < len(indices):
                col_idx = indices[target_idx]
                row[col_idx] = value
                matched.append(f"{pdf_key} -> '{sheet_header}' col {_col_letter(col_idx)}")
            else:
                unmatched.append(f"{pdf_key} -> '{sheet_header}' (no section {section} column)")

    return row, matched, unmatched


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------
class SheetsUploaderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Plankton Settlement \u2192 Google Sheets Uploader")
        self.root.geometry("700x520")

        self.pdf_path = tk.StringVar()
        self.sheet_url = tk.StringVar(
            value="https://docs.google.com/spreadsheets/d/15nptM0xgTDyDxeO79L-AlfoB5G6Y7l50p7tTyoELveE/edit?gid=1844026567#gid=1844026567"
        )

        pad = {"padx": 10, "pady": 5}

        # --- Step 1: Select PDF ---
        tk.Label(root, text="Step 1: Select Settlement Report PDF").pack(
            anchor="w", **pad
        )
        frame1 = tk.Frame(root)
        frame1.pack(fill="x", padx=10)
        tk.Entry(frame1, textvariable=self.pdf_path).pack(
            side="left", fill="x", expand=True
        )
        tk.Button(frame1, text="Browse", command=self._browse_pdf).pack(
            side="right", padx=5
        )

        # --- Step 2: Google Sheet URL ---
        tk.Label(root, text="Step 2: Paste Google Sheet URL").pack(
            anchor="w", **pad
        )
        tk.Entry(root, textvariable=self.sheet_url).pack(fill="x", padx=10)

        # --- Upload button ---
        self.btn_upload = tk.Button(
            root,
            text="UPLOAD TO SHEET",
            bg="#2ecc71",
            fg="white",
            font=("Arial", 12, "bold"),
            command=self._start_upload,
        )
        self.btn_upload.pack(pady=15)

        # --- Log area ---
        self.log_widget = tk.Text(root, height=14, bg="#f0f0f0", state="disabled")
        self.log_widget.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    def _browse_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if path:
            self.pdf_path.set(path)

    def _safe_log(self, message):
        self.root.after(0, self._do_log, message)

    def _do_log(self, message):
        self.log_widget.config(state="normal")
        self.log_widget.insert("end", message + "\n")
        self.log_widget.see("end")
        self.log_widget.config(state="disabled")

    def _start_upload(self):
        pdf = self.pdf_path.get().strip()
        url = self.sheet_url.get().strip()

        if not pdf or not os.path.exists(pdf):
            messagebox.showerror("Error", "Please select a valid PDF file.")
            return

        m = RE_SHEET_ID.search(url)
        if not m:
            messagebox.showerror(
                "Error",
                "Invalid Google Sheet URL.\n"
                "Expected: https://docs.google.com/spreadsheets/d/<ID>/...",
            )
            return
        sheet_id = m.group(1)

        self.btn_upload.config(state="disabled")
        self._safe_log("=" * 50)
        self._safe_log("Starting upload...")

        def work():
            try:
                # 1. Parse PDF
                self._safe_log("\n[1/3] Parsing PDF...")
                data = extract_from_pdf(pdf, log=self._safe_log)
                if not data:
                    self._safe_log("ERROR: No data extracted from PDF.")
                    return

                self._safe_log(f"\nExtracted {len(data)} field(s).")

                # 2. Authenticate
                self._safe_log("\n[2/3] Authenticating with Google...")
                gc = get_gspread_client(log=self._safe_log)

                # 3. Upload
                self._safe_log("\n[3/3] Uploading to Google Sheet...")
                spreadsheet = gc.open_by_key(sheet_id)
                sheet = spreadsheet.sheet1
                headers = sheet.row_values(7)

                if not headers:
                    self._safe_log("ERROR: Sheet has no headers in row 7.")
                    return

                self._safe_log(f"  Found {len(headers)} headers in row 7")

                # Log all headers with column letters for debugging
                for i, h in enumerate(headers):
                    if h.strip():
                        self._safe_log(f"    {_col_letter(i)}: '{h}'")

                # Build row via keyword matching
                row, matched, unmatched = build_row(
                    data, headers, log=self._safe_log
                )

                # Find next empty row after row 7
                all_values = sheet.col_values(1)  # check column A
                next_row = max(len(all_values) + 1, 8)  # at least row 8
                self._safe_log(f"\n  Writing to row {next_row}...")

                # Write each non-empty cell individually
                cells_to_update = []
                for col_idx, val in enumerate(row):
                    if val != "":
                        cell_ref = f"{_col_letter(col_idx)}{next_row}"
                        cells_to_update.append((cell_ref, val))

                if cells_to_update:
                    # Batch update all cells at once
                    batch = []
                    for cell_ref, val in cells_to_update:
                        batch.append({
                            "range": cell_ref,
                            "values": [[val]],
                        })
                    sheet.batch_update(batch, value_input_option="USER_ENTERED")

                self._safe_log(f"\nMatched {len(matched)}/{len(data)} fields:")
                for info in matched:
                    self._safe_log(f"  \u2713 {info}")

                if unmatched:
                    self._safe_log(f"\nWARNING: {len(unmatched)} unmatched:")
                    for info in unmatched:
                        self._safe_log(f"  \u2717 {info}")

                self._safe_log("\nDone! Row appended successfully.")
                self.root.after(
                    0,
                    lambda: messagebox.showinfo(
                        "Success",
                        f"Uploaded {len(matched)} field(s) to Google Sheet.",
                    ),
                )

            except FileNotFoundError as e:
                msg = str(e)
                self._safe_log(f"\nERROR: {msg}")
                self.root.after(0, lambda m=msg: messagebox.showerror("Missing File", m))
            except Exception as e:
                import traceback
                msg = f"{type(e).__name__}: {e}"
                self._safe_log(f"\nERROR: {msg}")
                self._safe_log(traceback.format_exc())
                self.root.after(0, lambda m=msg: messagebox.showerror("Error", m))
            finally:
                self.root.after(
                    0, lambda: self.btn_upload.config(state="normal")
                )

        threading.Thread(target=work, daemon=True).start()


if __name__ == "__main__":
    root = tk.Tk()
    SheetsUploaderGUI(root)
    root.mainloop()
