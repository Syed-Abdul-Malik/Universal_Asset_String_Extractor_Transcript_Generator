# 🛠️ Universal Asset String Extractor & Transcript Generator

A robust, team-agnostic utility designed to recursively scan, extract, and catalog printable text/ASCII strings from raw binary game archives (`.PAK`, `.IFF`, `.DAT`, `.BIN`), game configuration files, or raw text dumps. 

This tool was designed to replicate manual **Notepad++ string inspection** workflows programmatically—reducing days of tedious, error-prone manual copy-pasting into a single-click, verified extraction pipeline.

---

## ✨ Key Features

* **📦 Universal File Handling:** Parses any file format, whether they have extensions (`.xml`, `.bin`, `.pak`, `.dat`, `.txt`) or are extensionless numeric files (e.g., `0`, `1`, `242`).
* **🔄 Dynamic File Count & Naming:** Adaptable out of the box. No hardcoded file lists or expected counts—processes 1 to 10,000+ files of any name.
* **📂 Automated Google Drive Discovery:** When run in Google Colab, it recursively searches your entire Drive to automatically find target directories without requiring complex absolute paths.
* **🗂️ Subfolder Support (Recursive):** Traverses nested subdirectories (e.g., `Project/Missions/Act1/`) and logs the exact relative sub-path of each string in the output.
* **📊 Double-Sheet Excel Export:** 
  1. **Master Transcript:** Contains columns for *Root Folder*, *Sub-Path*, *File Name*, *Line ID*, *Extracted Text*, and *Character Length*.
  2. **Extraction Summary:** A verification sheet logging file counts, paths, and audit status.
* **🧼 Noise Filtering:** Features a minimum-length slider and alphanumeric checker to weed out binary noise and corrupt compiler data.

---

## 🚀 How to Use (Two Workflows)

### Option A: Run in Google Colab (Easiest / Non-Technical Teams)

No software installation or Python setup is required. The script automatically renders user-friendly form controls.

1. Create a new notebook in [Google Colab](https://colab.research.google.com/).
2. Copy and paste the code from [Universal_Extractor.py](./extract_strings.py) into a code cell.
3. Colab will display an interactive **Form UI** on the right side of the screen.
4. Fill in the fields:
   * **`PROJECT_NAME`**: Name of your project (e.g., `EA Games - Abdul Malik`).
   * **`TARGET_FOLDERS_INPUT`**: Comma-separated list of folder names to scan (e.g., `INTEL, CONV.PAK`).
   * **`INCLUDE_SUBFOLDERS`**: Check this if you have files organized inside nested folders.
   * **`MIN_STRING_LENGTH`**: Standard is `3` or `4` to capture game dialogue and filter out raw hex data.
5. Click **Run (▶️)** and authorize Google Drive. 
6. The fully processed `Master_Transcript` Excel workbook will automatically download straight to your browser.

---

### Option B: Run Locally (Command Line)

To run the script locally on your machine, ensure you have the required packages installed:

```bash
pip install pandas openpyxl
