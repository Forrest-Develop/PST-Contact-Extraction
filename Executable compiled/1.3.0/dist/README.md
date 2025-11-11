# PST Contacts Extractor GUI 1.0.0

A Windows PowerShell GUI tool for exporting contacts from Outlook PST files into Outlook 365–compatible CSV format.

---

## 📋 Overview
**PST Contacts Extractor GUI** provides a simple graphical interface that allows you to:
- Select a `.pst` file to extract contacts from.
- Choose an output location for a `.csv` file.
- Export all contacts from Outlook PST files using the Outlook COM interface.

This tool is particularly useful for migrations, archival, or rebuilding contact lists from offline Outlook data.

---

## ✨ Features
- **GUI-based workflow** – No command-line input required.
- **Full Outlook 365 CSV compatibility**.
- **Safe temporary PST mounting/unmounting**.
- **UTF-8 BOM encoding** for proper character imports.
- **Automatic COM cleanup** to prevent Outlook hangs.
- **No installation required** – standalone `.ps1` script.

---

## ⚙️ Requirements
- **Windows 10 or later**
- **Microsoft Outlook (desktop)** – must be installed and configured
- **PowerShell 5.1+** (Windows PowerShell, not PowerShell Core)

---

## 🚀 Usage
1. **Download** the script:
   ```bash
   PST_Contacts_Extractor_GUI_1.0.0.ps1
   ```

2. **Run** the script in PowerShell:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\PST_Contacts_Extractor_GUI_1.0.0.ps1
   ```

3. **Select** your PST file and a save location for the exported CSV.

4. **Click** “Extract Contacts”.

5. **Open Output Folder** when the export completes.

---

## 📄 Output Example
Example CSV header:
```csv
Title,First Name,Middle Name,Last Name,Suffix,Company,Department,Job Title,E-mail Address,E-mail 2 Address,...
```

Encoding: **UTF-8 with BOM** (compatible with Outlook import)

---

## ⚠️ Notes
- The script **does not modify** your PST or contacts.
- Outlook must be installed and initialized at least once.
- COM objects are released after use to prevent lockups.
- Password-protected PST files are not supported.

---

## 🧰 Version History
### v1.0.0 – October 2025
- Initial stable release
- COM-safe folder recursion (no type casting)
- Enhanced error logging and UI feedback
- Color-coded status messages
- UTF-8 BOM CSV output for Outlook 365

---

## 🪪 License
Copyright (c) 2025 **ForrestDev**. All rights reserved.

This software is the property of **ForrestDev** and is licensed for internal use by **Power Auto Group** only.

Redistribution, modification, copying, or alteration in any form, whole or in part, is strictly prohibited without express written permission from the author.

Usage is limited to the author and to Power Auto Group for authorized internal operations.

---

## 👤 Author
**ForrestDev**  
2025

