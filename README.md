# 📄 Data Point Ingestion Automation Tool

This Python-based tool automates the ingestion configuration process by mapping field names from a Data Dictionary file to Data Points. It supports both `.xls` and `.xlsx` Excel formats and is designed to be run easily via a batch file.

---

## ✅ Features

- Supports `.xls` (via `xlrd`) and `.xlsx` (via `openpyxl`)
- Reads from the **"Data_Dictionary"** sheet and **"CMDD_file"**
- Maps `Field Name (with ID)` to corresponding `Data Point Name`
- Easy to use — just double-click `ingestion.bat`

---

## 📦 Folder Structure
IngestionTool/
├── demo1.py

├── ingestion.py

├── newdtest1.py

├── ingestion.bat

├── requirements.txt

└── README.md

---

## 🖥️ How to Run the Tool (for Users)

### 🔹 1. **Install Python (if not already installed)**
Download Python 3.7 or above:  
👉 https://www.python.org/downloads/

✅ During installation, make sure to check **"Add Python to PATH"**

---

### 🔹 2. **Unzip the Folder**

- Right-click `IngestionTool.zip` → Extract All
- Open the extracted folder
- set path for **"Data_Dictionary"** sheet and **"CMDD_file"** in .py files

---

### 🔹 3. **Double-click `ingestion.bat`**

This will:
- Create a Python virtual environment (if not already created)
- Install required dependencies from `requirements.txt`
- Run the ingestion tool
- Prompt you to enter a **Field Name** and the **Document Name**

---

## 📝 How the Script Works

1. Prompts the user to enter a Field Name (with ID).
2. Asks for the full path to the .xls or .xlsx Data Dictionary file.
3. Searches for the matching row in the "Data_Dictionary" sheet.
4. Displays the corresponding Data Point Name.
5. Prompts the user to enter a keyword to search for the cmdd_path in the CMDD file, and select a matching entry.
6. If the data point is clonable, prompts the user to select a DataPoint to clone.
7. Asks the user to enter a Default Value and Transformation Rule.
8. Prompts the user to select a Collaboration Association.

---

## 📄 Excel File Format Requirements

- Sheet name must be: **`Data_Dictionary`**
- Must contain columns:
  - `Field Name (with ID)`
  - `Data Point Name`

---

## 🧠 Technical Requirements

- Python 3.7+
- The following Python packages (automatically installed via the batch file):
  - `pandas==2.0.1`
  - `xlrd==1.2.0`
  - `openpyxl==3.1.5`
  - `xlwt==1.3.0`
  - `xlutils==2.0.0`

---

## 💬 Sample Output

  Running Ingestion Tool...

Requirements already installed. Skipping installation.
✅ Running latest ingestion_template.py
🔹 Enter the Field ID (as in 'Field Name (with ID)'):
---

## 🆘 Troubleshooting

- ❌ `XLRDError: Excel xlsx file; not supported`:  
  You're using `.xlsx` with an outdated `xlrd`. This tool avoids that by auto-handling formats internally.
  
- ❌ `ModuleNotFoundError`:  
  Make sure to let `ingestion.bat` install required libraries.

- ❌ Python not found?  
  Reinstall Python and ensure **"Add to PATH"** was checked.

---
