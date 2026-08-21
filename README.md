# Ultimate PDF Toolkit Pro - Code Installation

This folder (`Code_Installation`) contains all the necessary source code, scripts, and configuration files to build the Ultimate PDF Toolkit Pro from scratch. 

Once compiled using the provided batch file, the generated setup installer (`Ultimate_PDF_Toolkit_Pro_Setup.exe`) will be created **directly in this folder**, as well as copied into the `Executable_Packages` directory. After installing the package via the setup installer, the application will stay on the PC forever (in your AppData/Start Menu), and you won't need to work with the setup `.exe` file again.

---

## 📂 Purpose of Every Single File in this Folder

| File Name | Purpose | Do You Need to Keep It? |
| :--- | :--- | :--- |
| **`Ultimate_PDF_Toolkit_Pro.py`** | The main Python source code containing all features (PDF/Word merge, split, convert, encrypt, OCR/preview). | ✅ **YES** — The core code. |
| **`setup_wizard.py`** | The Python source code that designs the Graphical Setup Wizard window. This wraps the core app into a beautiful installer. | ✅ **YES** — The installer code. |
| **`build_installer.bat`** | 1-Click Windows batch script. When you make changes to the code and double-click this script, it automatically recompiles everything into `Ultimate_PDF_Toolkit_Pro_Setup.exe` right in this same folder. It also places a copy in the `Executable_Packages` folder. | ✅ **YES** — Build tool. |
| **`requirements.txt`** | Lists the required Python libraries needed to build and run the app from source. | ✅ **YES** — Dependency list. |
| **`README.md` & `README.txt`** | The complete user manual, feature breakdown, and installation guide. | ✅ **YES** — Documentation. |
| **`.gitignore`** | Tells Git to ignore temporary compiler caches and large binaries when pushing to GitHub or other repositories. | ✅ **YES** — Git configuration. |

*(Note: During the build process, you will also see the generated `.exe` file appear in this folder).*

---

## 🚀 Features of the Application

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

---

## 🛠️ Developer Guide (Compiling & Building From Source)

If you modify the Python code and want to create a new installer:

1. **Test the app directly from code:**
   ```cmd
   pip install -r requirements.txt
   python Ultimate_PDF_Toolkit_Pro.py
   ```

2. **Rebuild the Setup Wizard Installer package (.exe):**
   - Double-click **`build_installer.bat`**.
   - It will compile the source code and automatically create the updated **`Ultimate_PDF_Toolkit_Pro_Setup.exe`** directly in **this Code_Installation directory**.
   - It will also automatically organize the portable app and installer into the `..\Executable_Packages` folder.

---

## 💻 How to Install (Sits on Your Computer Forever)

If you just want to use the application without dealing with code, go to the **`Executable_Packages\Installer_Version`** folder or run the `.exe` generated in this folder:

1. Double-click **`Ultimate_PDF_Toolkit_Pro_Setup.exe`**.
2. In the Setup Wizard, choose your options (Desktop Shortcut, Start Menu Search, App Registration) and click **"Install Now"**.
3. Click **"Finish"**.

**Once Installed (Just like WhatsApp or Word):**
- Launch from your **Desktop** icon.
- Press the **Windows Key**, type **"Ultimate PDF Toolkit Pro"**, and press **Enter**.
- It sits on your PC forever. You do not need to run the installer again.
- Uninstall anytime via **Windows Settings > Apps > Installed apps**.

---

## 🌐 GitHub / Repository Usage

If you plan to upload this to GitHub:
- Upload the `Code_Installation` directory.
- The included `.gitignore` will ensure that compiled binaries (like `.exe` files or `build/` folders) are excluded, keeping your repository clean and purely code-based.
