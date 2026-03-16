"""
Tkinter desktop GUI wrapper (Windows/Mac/Linux).
For iOS, use main.py (Kivy) instead.
"""

import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

from core import extract_from_pdf, upload_to_sheet, RE_SHEET_ID

DEFAULT_SHEET_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "15nptM0xgTDyDxeO79L-AlfoB5G6Y7l50p7tTyoELveE/"
    "edit?gid=1844026567#gid=1844026567"
)
APP_DIR = os.path.dirname(os.path.abspath(__file__))


class SheetsUploaderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Plankton Settlement \u2192 Google Sheets Uploader")
        self.root.geometry("700x520")

        self.pdf_path = tk.StringVar()
        self.sheet_url = tk.StringVar(value=DEFAULT_SHEET_URL)

        pad = {"padx": 10, "pady": 5}

        tk.Label(root, text="Step 1: Select Settlement Report PDF").pack(anchor="w", **pad)
        frame1 = tk.Frame(root)
        frame1.pack(fill="x", padx=10)
        tk.Entry(frame1, textvariable=self.pdf_path).pack(side="left", fill="x", expand=True)
        tk.Button(frame1, text="Browse", command=self._browse_pdf).pack(side="right", padx=5)

        tk.Label(root, text="Step 2: Paste Google Sheet URL").pack(anchor="w", **pad)
        tk.Entry(root, textvariable=self.sheet_url).pack(fill="x", padx=10)

        self.btn_upload = tk.Button(
            root, text="UPLOAD TO SHEET",
            bg="#2ecc71", fg="white", font=("Arial", 12, "bold"),
            command=self._start_upload,
        )
        self.btn_upload.pack(pady=15)

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
            messagebox.showerror("Error", "Invalid Google Sheet URL.")
            return
        sheet_id = m.group(1)

        self.btn_upload.config(state="disabled")
        self._safe_log("=" * 50)
        self._safe_log("Starting upload...")

        def work():
            try:
                self._safe_log("\n[1/2] Parsing PDF...")
                data = extract_from_pdf(pdf, log=self._safe_log)
                if not data:
                    self._safe_log("ERROR: No data extracted.")
                    return

                self._safe_log(f"\nExtracted {len(data)} field(s).")
                self._safe_log("\n[2/2] Uploading to Google Sheet...")

                matched, unmatched = upload_to_sheet(
                    sheet_id, data, APP_DIR, log=self._safe_log
                )

                self._safe_log(f"\nMatched {len(matched)} fields:")
                for info in matched:
                    self._safe_log(f"  \u2713 {info}")

                if unmatched:
                    self._safe_log(f"\nUnmatched ({len(unmatched)}):")
                    for info in unmatched:
                        self._safe_log(f"  \u2717 {info}")

                self._safe_log("\nDone! Row written successfully.")
                self.root.after(0, lambda: messagebox.showinfo(
                    "Success", f"Uploaded {len(matched)} field(s)."))

            except Exception as e:
                import traceback
                msg = f"{type(e).__name__}: {e}"
                self._safe_log(f"\nERROR: {msg}")
                self._safe_log(traceback.format_exc())
                self.root.after(0, lambda m=msg: messagebox.showerror("Error", m))
            finally:
                self.root.after(0, lambda: self.btn_upload.config(state="normal"))

        threading.Thread(target=work, daemon=True).start()


if __name__ == "__main__":
    root = tk.Tk()
    SheetsUploaderGUI(root)
    root.mainloop()
