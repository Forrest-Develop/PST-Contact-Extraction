# 📦 PST Contacts Extractor GUI — Version 1.3.8

**Author:** Benjamin Forrest  
**Copyright:** © 2025 Benjamin Forrest  
**Software Property of:** Benjamin Forrest  
**Usage:** Internal use only by Power Auto Group and the Author (Benjamin Forrest)  
**License:** Redistribution or alteration of this software in any form is strictly prohibited without written permission from the author.

---

## 🧭 Overview

**PST Contacts Extractor GUI** is a Windows-based utility that allows users to extract contacts directly from Microsoft Outlook **.pst** files and export them into an **Outlook-compatible CSV format**.

Version **1.3.8** introduces full compatibility with the **standalone packaged executable (.exe)** — eliminating the need to run PowerShell scripts or manage dependencies manually.

---

## 🚀 Features

- Fully standalone **Windows executable**
- Extracts all contacts from Outlook PST files
- Exports data in **Outlook CSV format** (ready for Outlook 365 or Outlook Desktop import)
- Optional **manual folder selection** using Outlook’s COM API
- Automatically detects and extracts from the **default Contacts folder** if unchecked
- Built-in **license/TOS agreement modal** before app launch
- Custom application icon (`app.ico`)
- Compatible with both **PowerShell 5** and **PowerShell 7**, packaged into `.exe`
- Internal logging and graceful error handling
- Requires only a local installation of **Microsoft Outlook** — no additional third-party libraries or modules

---

## 🧰 System Requirements

- **Operating System:** Windows 10 or Windows 11  
- **Outlook:** Microsoft Outlook (installed)  
- **Runtime:** None required — runs as a standalone `.exe`  
- **Permissions:** Must have access to read `.pst` files and write CSV output

---

## 🪄 How to Use

1. **Launch** `PST_Contacts_Extractor_GUI-1.3.8.exe`  
   The application will initialize and verify that Microsoft Outlook is installed.

2. **Read & Accept** the license agreement modal to continue.

3. **Load PST File**  
   - By default, the program connects automatically to Outlook’s default Contacts folder.  
   - ✅ **Optional:** Check **“Manually select PST contacts folder”** to browse and choose a specific folder inside your PST file before extraction.

4. **Click “Export Contacts”**  
   - The tool extracts all available contact fields (name, company, phone, email, etc.)  
   - Exports to a file named:
     ```
     contacts_outlook365.csv
     ```

5. **Review your CSV** and import it into Outlook 365 or Outlook Desktop.

---

## 📁 Output Format

Exports `contacts_outlook365.csv` with the following columns:

| Field | Description |
|-------|--------------|
| First Name | Contact’s given name |
| Last Name | Contact’s surname |
| Company Name | Organization or employer |
| Business Phone | Work number |
| Home Phone | Personal phone |
| Mobile Phone | Mobile number |
| Email Address | Primary email |
| Job Title | Role or position |
| Business Street | Work address |
| Business City | Work city |
| Business State | Work state |
| Business Postal Code | Work ZIP/postal code |
| Notes | Additional notes |

---

## ⚠️ Notes and Troubleshooting

- If the icon fails to appear, ensure the executable was built with the embedded `app.ico`.  
- The `.exe` version resolves previous icon and path binding issues seen in earlier PowerShell builds.  
- Do **not** rename or modify the executable or its internal resources.  
- Authorized for internal use only by **Power Auto Group** and **Benjamin Forrest**.

---

## 🧾 Version History

| Version | Date | Notes |
|----------|------|-------|
| **1.3.8** | Nov 2025 | Standalone `.exe` version, ps2exe path fix, icon binding patch |
| **1.3.7** | Oct 2025 | Manual folder selection feature added |
| **1.3.6** | Oct 2025 | License modal and footer added |
| **1.3.5** | Sep 2025 | GUI performance improvements and parsing fixes |
| **1.0.0** | Aug 2025 | Initial release |
