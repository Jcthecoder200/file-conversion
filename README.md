# File Converter — Setup & Usage

## 1. Install Python
You need Python 3.9+ from python.org (tick "Add to PATH" during install).
Tkinter — used for the window itself — ships with Python on Windows, so
there's nothing extra to install for the interface.

## 2. Install the conversion libraries
Open Command Prompt in the folder with these files and run:

pip install -r requirements.txt


This installs:
- **Pillow** – image format conversions (png/jpg/bmp/gif/tiff/webp) and image→pdf
- **PyMuPDF** – pdf→image and pdf→text
- **python-docx** – docx→txt and txt→docx
- **pdf2docx** – pdf→docx
- **docx2pdf** – docx→pdf (drives Microsoft Word in the background — Word
  must be installed for this one conversion only; everything else works
  without Word)
- **fpdf2** – txt→pdf

## 3. Run it

python file_conversion.py


## How to use

### Single file conversion
1. **Select File...** — pick the file you want to convert. Its name and detected type appear next to the button.
2. Middle row shows the detected source type on the left and a **drop‑down of valid target types** on the right — only conversions that actually make sense for that file are listed (e.g. a `.png` won't offer `.docx`, since that would need OCR, not a format conversion).
3. Bottom row: type the output file name, choose the output folder, then either:
   - **Convert** (right button) — does a normal, whole‑file conversion, or
   - **Slice Image...** (middle button, images only) — opens a window where you drag the 4 corner dots onto the region you want to keep, then confirms to crop and convert just that region.

### Batch folder conversion
- Click the **Batch Convert...** button (leftmost in the bottom row).
- A new window opens:
  - Choose a **source folder** (all files inside will be scanned).
  - Choose an **output folder** (auto‑filled as `<source>_converted` – you can change it).
  - Pick a **target file type** from the dropdown.
- Click **Start Batch** – the tool will convert every file in the source folder that supports the chosen target type.
- Each output file keeps its original name but uses the new extension. If a file with that name already exists in the output folder, a number (`_1`, `_2`, …) is automatically added to avoid overwriting.
- A progress label shows which file is being processed, and a final summary tells you how many succeeded and lists any errors.

---

## Supported types
`.pdf`, `.docx`, `.txt`, `.png`, `.jpg`/`.jpeg`, `.bmp`, `.gif`, `.tiff`, `.webp`

---

## Notes
- Multi‑page PDFs converted to an image format produce one image per page (named `..._page1`, `..._page2`, …).
- If a library is missing, the app still opens — you'll only get an error message when you try the specific conversion that needs it, telling you exactly what to `pip install`.
- The `.docx` → `.pdf` conversion requires **Microsoft Word** to be installed on your Windows PC. If Word is not installed, you'll see a clear error message instead of a crash.
