# Ultimate PDF Toolkit Pro

A comprehensive, all-in-one desktop application built with Python and Tkinter for managing, manipulating, and converting PDF and Word documents. It features a modern, dark-themed, DPI-aware interface designed for maximum productivity.

## 🚀 Features

- **Organize PDF**: Add a PDF, view pages, remove unwanted pages, or reorder them.
- **Merge PDF**: Combine multiple PDF files into a single document.
- **Merge Word**: Combine multiple `.docx` or `.doc` files into a single Word document.
- **Convert PDF ↔ Word**: Bi-directional conversion between PDF and Word.
- **Convert Image ↔ PDF**: Convert JPG/PNG images to a single PDF, or extract all pages of a PDF into individual images.
- **PDF to PowerPoint (PPTX)**: Convert PDF pages directly into presentation slides.
- **Split PDF**: Extract a specific range of pages (e.g. Page 1 to 5) or split the entire document into single-page PDFs.
- **Lock & Unlock PDF**: Secure your PDFs with password encryption, or remove passwords from encrypted PDFs.
- **Rotate PDF**: Rotate PDF pages (90°, 180°, 270°).
- **Annotate & Extract**: Extract embedded text from a PDF, or search and highlight specific keywords across the entire document.
- **Live Preview**: Click the "Preview PDF" button on any tab to view the PDF right inside the application.

## 📋 Requirements

This application relies on the following Python packages:
- `pymupdf` (for PDF viewing, image extraction, and highlighting)
- `pypdf` (for merging, splitting, rotating, and encrypting)
- `Pillow` (for image processing)
- `pdf2docx` (for PDF to Word conversion)
- `python-pptx` (for PDF to PowerPoint conversion)
- `pywin32` (for Word to Word merging and Word to PDF conversion on Windows)

## 🛠️ Installation & Setup

1. **Clone the repository:**
   ```bash
   git clone https://github.com/yourusername/Ultimate-PDF-Toolkit-Pro.git
   cd Ultimate-PDF-Toolkit-Pro
   ```

2. **Install dependencies:**
   Make sure you have Python installed, then run:
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the script directly:**
   ```bash
   python "PDF, Word file merger.py"
   ```

## 📦 Building the Executable (.exe)

If you want to package the application into a standalone `.exe` file that can be shared and run without installing Python:

1. Double-click the provided `build_exe.bat` script.
2. Wait for the compilation to finish.
3. Your executable will be available in the `dist/` folder.

*(Note: Depending on your system, Windows Defender might temporarily flag newly created PyInstaller executables. This is a common false-positive).*

## ⚠️ Notes

- **Word automation features** (Merging Word documents and converting Word to PDF) require Microsoft Word to be installed on the host Windows machine, as it utilizes the `pywin32` COM interface.
- **DPI Awareness** is enabled by default to ensure fonts look sharp on high-resolution Windows displays.
