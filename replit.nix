{ pkgs }:
{
  # System packages for the Streamlit app.
  # tesseract = optional OCR for scanned tax bills (the app degrades gracefully
  # if it's ever missing). pymupdf/pdfplumber/openpyxl need no system libs.
  deps = [
    pkgs.tesseract
  ];
}
