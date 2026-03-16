"""
Plankton Settlement Report -> Google Sheets Uploader
Kivy iOS/Android/Desktop app entry point.
"""

import os
import threading

from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.scrollview import ScrollView
from kivy.clock import Clock
from kivy.core.window import Window
from kivy.utils import platform

from core import extract_from_pdf, upload_to_sheet, RE_SHEET_ID

# Default sheet URL
DEFAULT_SHEET_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "15nptM0xgTDyDxeO79L-AlfoB5G6Y7l50p7tTyoELveE/"
    "edit?gid=1844026567#gid=1844026567"
)


class UploaderApp(App):
    def build(self):
        self.title = "Plankton Settlement Uploader"
        Window.clearcolor = (0.96, 0.96, 0.96, 1)

        root = BoxLayout(orientation="vertical", padding=15, spacing=10)

        # --- Step 1: PDF file ---
        root.add_widget(Label(
            text="Step 1: Select Settlement Report PDF",
            size_hint_y=None, height=30,
            color=(0.2, 0.2, 0.2, 1), halign="left",
        ))

        pdf_row = BoxLayout(size_hint_y=None, height=45, spacing=8)
        self.pdf_input = TextInput(
            hint_text="Tap Browse to select a PDF...",
            multiline=False, font_size=14,
        )
        pdf_row.add_widget(self.pdf_input)

        browse_btn = Button(
            text="Browse", size_hint_x=0.3,
            background_color=(0.3, 0.3, 0.3, 1),
            color=(1, 1, 1, 1),
        )
        browse_btn.bind(on_press=self._browse_pdf)
        pdf_row.add_widget(browse_btn)
        root.add_widget(pdf_row)

        # --- Step 2: Sheet URL ---
        root.add_widget(Label(
            text="Step 2: Google Sheet URL",
            size_hint_y=None, height=30,
            color=(0.2, 0.2, 0.2, 1), halign="left",
        ))

        self.url_input = TextInput(
            text=DEFAULT_SHEET_URL,
            multiline=False, font_size=14,
            size_hint_y=None, height=45,
        )
        root.add_widget(self.url_input)

        # --- Upload button ---
        self.upload_btn = Button(
            text="UPLOAD TO SHEET",
            size_hint_y=None, height=55,
            background_color=(0.18, 0.8, 0.44, 1),  # #2ecc71
            color=(1, 1, 1, 1),
            font_size=18, bold=True,
        )
        self.upload_btn.bind(on_press=self._start_upload)
        root.add_widget(self.upload_btn)

        # --- Log area ---
        root.add_widget(Label(
            text="Log",
            size_hint_y=None, height=25,
            color=(0.4, 0.4, 0.4, 1), halign="left",
        ))

        scroll = ScrollView(size_hint=(1, 1))
        self.log_output = TextInput(
            text="Ready.\n",
            readonly=True, multiline=True,
            font_size=13, background_color=(0.94, 0.94, 0.94, 1),
        )
        scroll.add_widget(self.log_output)
        root.add_widget(scroll)

        self.selected_pdf = None
        return root

    # -- File picker --
    def _browse_pdf(self, _instance):
        if platform == "ios":
            self._browse_ios()
        else:
            self._browse_desktop()

    def _browse_desktop(self):
        """Desktop file picker (Plyer or fallback to Tkinter)."""
        try:
            from plyer import filechooser
            paths = filechooser.open_file(
                title="Select PDF",
                filters=[["PDF files", "*.pdf"]],
            )
            if paths:
                self._set_pdf(paths[0])
        except Exception:
            # Fallback: Tkinter file dialog (works on Windows/Mac/Linux)
            try:
                import tkinter as tk
                from tkinter import filedialog
                tmp = tk.Tk()
                tmp.withdraw()
                path = filedialog.askopenfilename(
                    filetypes=[("PDF Files", "*.pdf")]
                )
                tmp.destroy()
                if path:
                    self._set_pdf(path)
            except Exception as e:
                self._log(f"File picker error: {e}")

    def _browse_ios(self):
        """iOS document picker via pyobjus."""
        try:
            from plyer import filechooser
            filechooser.open_file(
                on_selection=self._on_ios_file_selected,
                filters=["public.data"],
            )
        except Exception as e:
            self._log(f"iOS file picker error: {e}")
            self._log("You can also paste the file path directly.")

    def _on_ios_file_selected(self, selection):
        if selection:
            self._set_pdf(selection[0])

    def _set_pdf(self, path):
        self.selected_pdf = path
        self.pdf_input.text = os.path.basename(path)
        self._log(f"Selected: {os.path.basename(path)}")

    # -- Logging --
    def _log(self, message):
        Clock.schedule_once(lambda dt: self._do_log(message), 0)

    def _do_log(self, message):
        self.log_output.text += message + "\n"
        self.log_output.cursor = (0, len(self.log_output.text))

    # -- Upload --
    def _start_upload(self, _instance):
        pdf_path = self.selected_pdf or self.pdf_input.text.strip()
        url = self.url_input.text.strip()

        if not pdf_path or not os.path.exists(pdf_path):
            self._log("ERROR: Please select a valid PDF file.")
            return

        m = RE_SHEET_ID.search(url)
        if not m:
            self._log("ERROR: Invalid Google Sheet URL.")
            return
        sheet_id = m.group(1)

        self.upload_btn.disabled = True
        self._log("=" * 40)
        self._log("Starting upload...")

        # Resolve app directory for credentials
        app_dir = os.path.dirname(os.path.abspath(__file__))

        def work():
            try:
                self._log("\n[1/2] Parsing PDF...")
                data = extract_from_pdf(pdf_path, log=self._log)
                if not data:
                    self._log("ERROR: No data extracted.")
                    return

                self._log(f"\nExtracted {len(data)} field(s).")
                self._log("\n[2/2] Uploading to Google Sheet...")

                matched, unmatched = upload_to_sheet(
                    sheet_id, data, app_dir, log=self._log
                )

                self._log(f"\nMatched {len(matched)} fields:")
                for info in matched:
                    self._log(f"  OK: {info}")

                if unmatched:
                    self._log(f"\nUnmatched ({len(unmatched)}):")
                    for info in unmatched:
                        self._log(f"  -- {info}")

                self._log("\nDone! Row written successfully.")

            except Exception as e:
                self._log(f"\nERROR: {type(e).__name__}: {e}")
            finally:
                Clock.schedule_once(
                    lambda dt: setattr(self.upload_btn, "disabled", False), 0
                )

        threading.Thread(target=work, daemon=True).start()


if __name__ == "__main__":
    UploaderApp().run()
