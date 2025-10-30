# PST_Contacts_Extractor_GUI_1.0.0

### Overview
**PST_Contacts_Extractor_GUI_1.0.0** is a standalone PowerShell script that provides a simple Windows GUI for extracting all contacts from an Outlook PST file into a CSV formatted for direct import into Outlook 365 or Outlook.com.

The script uses Outlook's COM interface to temporarily mount a PST file, traverse all Contact folders, and export the contents into a UTF-8 BOM CSV. It automatically unmounts the PST when finished.

---

### Features
- **Graphical Interface** – Select the PST file and the CSV save location using standard Windows dialogs.
- **Outlook 365 Compatible CSV** – Output file matches Outlook’s expected import format.
- **Automatic PST Mount/Unmount** – Safely attaches the PST in Outlook, then removes it after export.
- **UTF-8 Encoding with BOM** – Ensures proper import of non-ASCII characters.
- **Comprehensive Field Mapping** – Includes all standard Outlook contact fields (email, phone numbers, addresses, etc.).

---

### Requirements
- Windows 10 or later
- Microsoft Outlook (desktop version) installed
- PowerShell 5.1 or later (Windows PowerShell, not PowerShell Core)

---

### Usage
1. **Save** the script as:
   ```
   PST_Contacts_Extractor_GUI_1.0.0.ps1
   ```

2. **Run** it in PowerShell:
   ```powershell
   powershell -ExecutionPolicy Bypass -File "PST_Contacts_Extractor_GUI_1.0.0.ps1"
   ```

3. **Choose** your `.pst` file when prompted.

4. **Select** where to save your exported `contacts_outlook365.csv` file.

5. **Click** **Extract Contacts**.

6. When finished, click **Open Output Folder** to view the results.

---

### Output
- The script generates a CSV file formatted for Outlook import.
- Encoding: UTF-8 with BOM
- Sample header:
  ```csv
  Title,First Name,Middle Name,Last Name,Suffix,Company,Department,Job Title,E-mail Address,E-mail 2 Address,...
  ```

---

### Notes
- The script does **not** modify your PST or contacts.
- Outlook must be installed (it uses MAPI/COM from Outlook).
- If you encounter an error about `System.__ComObject`, ensure Outlook is fully installed and configured at least once.
- All COM objects are safely released to prevent Outlook lockups.
- The `$HOME` environment variable is never overwritten.

---

### Known Limitations
- Must run under a Windows account with Outlook profile access.
- Cannot process password-protected PSTs.
- Does not support non-contact items (only Contact folders are scanned).

---

### Version History
**1.0.0**  
Initial stable release (GUI-based, COM-safe build)
- Removed strict folder type casting to fix `System.__ComObject` conversion error
- Improved color-safe logging system
- Added BOM-safe CSV writer
- General error handling and status reporting

---

### License
Copyright (c) 2025 Benjamin Forrest. All rights reserved.

This software is the property of **Benjamin Forrest** and is licensed for internal use by **Power Auto Group** only. Redistribution, modification, copying, or alteration in any form, whole or in part, is strictly prohibited without the express written permission of the author.

Usage is limited to the author and to Power Auto Group for authorized internal operations. No part of this software may be incorporated into other tools, scripts, or commercial products without prior approval.

---

### Author
Benjamin Forrest  
IT Department, Power Auto Group  
October 2025

