PST CONTACTS EXTRACTOR GUI - VERSION 1.3.8
------------------------------------------

Author: Benjamin Forrest
Copyright: © 2025 Benjamin Forrest
Software Property of: Benjamin Forrest
Usage: Internal use only by Power Auto Group and the Author (Benjamin Forrest)
License: Redistribution or alteration of this software in any form is strictly prohibited without written permission from the author.


OVERVIEW
---------
PST Contacts Extractor GUI is a Windows-based tool that extracts contacts directly from Microsoft Outlook .pst files and exports them into an Outlook-compatible CSV format.

Version 1.3.8 introduces full compatibility with the standalone packaged executable (.exe), removing the need for PowerShell scripts or external dependencies.


FEATURES
---------
- Fully standalone Windows executable
- Extracts contacts from Outlook PST files
- Exports data in Outlook CSV format
- Optional manual folder selection using Outlook’s COM API
- Automatically detects default Contacts folder if not manually selected
- Built-in license/TOS modal
- Custom application icon (app.ico)
- Internal logging and error handling
- Requires only Microsoft Outlook to be installed; no additional third-party libraries or modules


SYSTEM REQUIREMENTS
-------------------
- Windows 10 or Windows 11
- Microsoft Outlook (installed)
- No PowerShell runtime required (standalone .exe)
- Local file read/write permissions


HOW TO USE
-----------
1. Launch PST_Contacts_Extractor_GUI-1.3.8.exe
   - The app verifies that Microsoft Outlook is installed.

2. Read and accept the license agreement to proceed.

3. Load PST File:
   - Automatically connects to the default Contacts folder.
   - OPTIONAL: Check “Manually select PST contacts folder” to browse and select a specific folder within the PST before extraction.

4. Click “Export Contacts”
   - The app extracts all available contact fields and saves them to:
     contacts_outlook365.csv

5. Review your CSV file before importing into Outlook 365 or Outlook Desktop.


OUTPUT FORMAT
--------------
contacts_outlook365.csv includes:
First Name, Last Name, Company Name, Business Phone, Home Phone, Mobile Phone,
Email Address, Job Title, Business Street, City, State, Postal Code, Notes


NOTES AND TROUBLESHOOTING
--------------------------
- If the icon does not appear, ensure app.ico is embedded in the .exe.
- Do not rename or modify the executable or internal resources.
- Authorized for use by Power Auto Group and Benjamin Forrest only.


VERSION HISTORY
----------------
1.3.8 - Nov 2025  | Standalone .exe build, ps2exe path fix, icon binding patch
1.3.7 - Oct 2025  | Manual folder selection feature added
1.3.6 - Oct 2025  | License modal and footer added
1.3.5 - Sep 2025  | GUI and parsing improvements
1.0.0 - Aug 2025  | Initial release
