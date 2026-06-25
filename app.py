"""Desktop app entry point. Opens a native window rendering the web/ UI and
exposes the core.Api methods to JavaScript via pywebview."""

import os
import sys

import webview

from core import Api


def _resource(*parts):
    """Path to bundled resources, working both from source and from a PyInstaller build."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, *parts)


class AppApi(Api):
    """core.Api plus UI-only helpers that need the window (native file dialog)."""

    def pick_excel(self):
        win = webview.windows[0]
        result = win.create_file_dialog(
            webview.OPEN_DIALOG, allow_multiple=False,
            file_types=("Excel files (*.xlsx;*.xlsm)", "All files (*.*)"))
        if not result:
            return {"ok": False, "cancelled": True}
        return self.load_excel(result[0], sheet=None) | {"path": result[0]}

    def load_excel_sheet(self, path, sheet):
        return self.load_excel(path, sheet=sheet)


def main():
    api = AppApi()
    webview.create_window(
        "Lender Importer", _resource("web", "index.html"),
        js_api=api, width=760, height=820, min_size=(680, 700))
    webview.start()


if __name__ == "__main__":
    main()
