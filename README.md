# RehaCompareGUI

**RehaCompareGUI** is a PowerShell-based forensic utility that provides a graphical interface for comparing two directory trees using **path/name-based analysis** and **cryptographic hashing (SHA-256)**.

The tool was developed to support digital forensic examiners reviewing **large provider-returned datasets** where manual comparison is time-consuming, error-prone, and difficult to document defensibly.

---

## :floppy_disk: Key Features 

- GUI-based folder selection (no command-line interaction required)
- Recursive directory comparison
- Multiple comparison methods:
  - **Relative path comparison** (recommended)
  - **Filename-only comparison** (optional)
  - **SHA-256 content hashing**
- Identifies:
  - Files only in Folder A
  - Files only in Folder B
  - Files not common to both folders
  - Files with identical content but different names or paths
- Handles locked or in-use files gracefully and logs errors
- Captures run metadata:
  - Case number
  - Operator
  - Run date/time (local and UTC)
- Generates a detailed `Summary.txt` including:
  - Run metadata
  - File counts
  - Comparison results
  - Tool version and script hash for provenance

---

## :bookmark: Intended Use

RehaCompareGUI is intended for **internal forensic and investigative workflows**, including:

- Verifying completeness of provider returns
- Identifying newly disclosed files across multiple returns
- Reducing manual file comparison time
- Producing repeatable, explainable comparison results

The tool performs **read-only analysis** and does not modify files.

---

## :warning: Requirements

- Windows 10 or later
- PowerShell 5.1 or PowerShell 7+
- Read access to both directories being compared

No external PowerShell modules are required.

---

## :file_folder: Usage

1. Launch the script (PowerShell v5️⃣):
   ```powershell
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
   .\RehaCompareGUI.ps1
   
2. Launch the script (PowerShell v7️⃣):
   ```powershell
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
   .\RehaCompareGUI_PS7.ps1

3. In the GUI:
   - Select Folder A and Folder B
   - (Optional) Enter Case Number and Operator
   - Choose desired comparison options
   - Select an output directory
   - Click **Run Comparison**

4. Output Files
  -OnlyIn_A_ByPath.txt
  -OnlyIn_B_ByPath.txt
  -NotInBoth_ByPath.txt
  -OnlyIn_A_ByName.txt
  -OnlyIn_B_ByName.txt
  -NotInBoth_ByName.txt
  -OnlyIn_A_ByHash.csv
  -OnlyIn_B_ByHash.csv
  -SameHash_DifferentPath.csv
  -HashErrors_A.csv
  -HashErrors_B.csv
  -Summary.txt  

**Comparison Method Notes**

Relative Path Comparison (Recommended)
- Treats files as the same only if their relative paths match. This avoids filename collision issues.

Filename-Only Comparison
- Useful for quick checks, but may produce false matches if duplicate filenames exist in different directories.

Hash Comparison
- Uses SHA-256 to determine true content equality regardless of filename or location.

**Forensic Considerations**

The script records its own SHA-256 hash in Summary.txt for provenance.

Results are deterministic based on inputs and selected options.

If an executable launcher is used, the PowerShell script remains the authoritative implementation.

**License**

This project is provided as-is for professional and community use.
No warranty is expressed or implied.

**Author**

Curtis Reha
Digital Forensic Examiner

Community feedback and contributions are welcome.
