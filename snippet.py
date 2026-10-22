# @title 🛠️ Universal Asset & String Extractor { display-mode: "form" }
# @markdown Run this cell to configure and extract strings from any Google Drive folder.
# @markdown Works across all teams, file formats, naming conventions, and file quantities.

# ==============================================================================
# 1. USER INPUT CONTROLS (Edit fields below or in the form panel)
# ==============================================================================
# @markdown **Project or Game Title** (Used to name the output Excel file):
PROJECT_NAME = "Asset_Extraction_Project" # @param {type:"string"}

# @markdown **Target Folder Names or Exact Paths** (Comma-separated, e.g. `INTEL, CONV.PAK` or full path `/content/drive/MyDrive/MyProject`):
TARGET_FOLDERS_INPUT = "INTEL, CONV.PAK" # @param {type:"string"}

# @markdown **Include Subfolders?** (If checked, recursively traverses all sub-directories):
INCLUDE_SUBFOLDERS = True # @param {type:"boolean"}

# @markdown **Minimum String Length** (Filter out binary noise; standard is 3 or 4):
MIN_STRING_LENGTH = 3 # @param {type:"slider", min:2, max:10, step:1}

# @markdown **Auto-Download Excel?**
AUTO_DOWNLOAD = True # @param {type:"boolean"}
# ==============================================================================

import os
import re
import pandas as pd
from google.colab import drive, files

print("⚡ Step 1: Connecting to Google Drive...")
drive.mount('/content/drive')

OUTPUT_FILE = f"{PROJECT_NAME.strip().replace(' ', '_')}_Master_Transcript.xlsx"
STRING_REGEX = re.compile(rb'[\x20-\x7E]{' + str(MIN_STRING_LENGTH).encode() + rb',}')

def find_target_paths(search_term):
    """
    Locates matching folders. Checks if input is an exact path first,
    otherwise searches recursively across Google Drive.
    """
    clean_term = search_term.strip()
    if not clean_term:
        return []

    # Case 1: Full path provided directly
    if os.path.exists(clean_term) and os.path.isdir(clean_term):
        return [clean_term]
    
    # Case 2: Relative MyDrive path
    direct_mydrive = os.path.join('/content/drive/MyDrive', clean_term)
    if os.path.exists(direct_mydrive) and os.path.isdir(direct_mydrive):
        return [direct_mydrive]

    # Case 3: Search Drive recursively for folder name
    matches = []
    print(f"🔍 Searching Google Drive for folder: '{clean_term}'...")
    for root, dirs, _ in os.walk('/content/drive/MyDrive'):
        for d in dirs:
            if d.lower() == clean_term.lower():
                matches.append(os.path.join(root, d))
    
    return matches

def extract_strings_from_file(file_path):
    """Reads any file format in binary mode and extracts clean ASCII strings."""
    lines = []
    try:
        with open(file_path, "rb") as f:
            content = f.read()
            raw_matches = STRING_REGEX.findall(content)
            for m in raw_matches:
                try:
                    decoded = m.decode('ascii', errors='ignore').strip()
                    # Keep strings containing at least one alphanumeric character
                    if len(decoded) >= MIN_STRING_LENGTH and any(c.isalnum() for c in decoded):
                        lines.append(decoded)
                except Exception:
                    continue
    except Exception as e:
        print(f"   ↳ ❌ Error reading '{os.path.basename(file_path)}': {e}")
    return lines

def execute_universal_extraction():
    # Parse comma-separated user inputs
    target_tokens = [t.strip() for t in TARGET_FOLDERS_INPUT.split(",") if t.strip()]
    if not target_tokens:
        print("❌ Error: Please specify at least one folder name in 'TARGET_FOLDERS_INPUT'.")
        return

    resolved_paths = []
    for token in target_tokens:
        found = find_target_paths(token)
        if found:
            resolved_paths.extend(found)
        else:
            print(f"⚠️ Warning: Could not locate any folder matching '{token}' in Google Drive.")

    # Remove duplicates while preserving order
    resolved_paths = list(dict.fromkeys(resolved_paths))

    if not resolved_paths:
        print("❌ No matching folders found. Please check your folder names in Google Drive.")
        return

    all_records = []
    folder_summary = []

    print("\n⚡ Step 2: Processing Files and Extracting Strings...")
    
    for folder_path in resolved_paths:
        folder_display_name = os.path.basename(folder_path)
        print(f"\n📂 Processing Folder: '{folder_display_name}' ({folder_path})")

        # Collect files (with or without subfolders)
        file_paths = []
        if INCLUDE_SUBFOLDERS:
            for root, _, files_in_dir in os.walk(folder_path):
                for f in files_in_dir:
                    if not f.startswith('.'):
                        file_paths.append(os.path.join(root, f))
        else:
            file_paths = [
                os.path.join(folder_path, f) for f in os.listdir(folder_path)
                if os.path.isfile(os.path.join(folder_path, f)) and not f.startswith('.')
            ]

        file_paths.sort()
        file_count = len(file_paths)
        strings_in_folder = 0

        print(f"   ↳ Found {file_count} files to inspect.")

        for fpath in file_paths:
            relative_subpath = os.path.relpath(fpath, folder_path)
            fname = os.path.basename(fpath)
            extracted_lines = extract_strings_from_file(fpath)
            strings_in_folder += len(extracted_lines)

            for line_no, text_content in enumerate(extracted_lines, start=1):
                all_records.append({
                    "Root Folder": folder_display_name,
                    "Sub-Path": relative_subpath,
                    "File Name": fname,
                    "Line ID": line_no,
                    "Extracted Text / String": text_content,
                    "Length": len(text_content)
                })

        folder_summary.append({
            "Folder Name": folder_display_name,
            "Full Directory Path": folder_path,
            "Total Files Discovered": file_count,
            "Total Strings Extracted": strings_in_folder,
            "Status": "COMPLETED" if file_count > 0 else "NO FILES FOUND"
        })

    if not all_records:
        print("\n❌ Extraction finished with 0 records. Files were empty or contained no printable strings.")
        return

    # Consolidate into Excel
    print("\n⚡ Step 3: Compiling Excel Master Transcript...")
    df_transcript = pd.DataFrame(all_records)
    df_summary = pd.DataFrame(folder_summary)

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        df_transcript.to_excel(writer, sheet_name="Master Transcript", index=False)
        df_summary.to_excel(writer, sheet_name="Extraction Summary", index=False)

    print(f"🎉 Success! Excel workbook generated: '{OUTPUT_FILE}'")

    # Display Verification Table
    print("\n" + "="*70)
    print("                     EXTRACTION AUDIT REPORT                       ")
    print("="*70)
    print(df_summary[["Folder Name", "Total Files Discovered", "Total Strings Extracted", "Status"]].to_string(index=False))
    print("="*70 + "\n")

    if AUTO_DOWNLOAD:
        print("📥 Downloading spreadsheet to your computer...")
        files.download(OUTPUT_FILE)

execute_universal_extraction()
