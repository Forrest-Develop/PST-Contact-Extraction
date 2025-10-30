# PST Contacts Extractor GUI (Standalone, COM-safe)
# -------------------------------------------------------------
# - Lets a user pick a .PST and a CSV path
# - Uses Outlook COM to mount the PST temporarily and export contacts to an Outlook-friendly CSV
# - No external scripts required
#
# Key fixes in this build:
# - Removed strict [Microsoft.Office.Interop.Outlook.MAPIFolder] type to avoid __ComObject cast error
# - Logger accepts color names via -Color 'Gray' and converts internally
# - Escaped 'Assistant''s Phone' correctly
# - Normalized CR/LF regex and CSV escaping
# - PowerShell backtick for line continuation
# -------------------------------------------------------------

# --- WinForms prerequisites ---
param([switch]$AcceptLicense)

# If you are on Windows PowerShell 5.1, -STA is default in the console.
# If you run this with PowerShell 7 (pwsh), WinForms can behave poorly in MTA.
# Enforce STA for reliability on WinForms/COM:
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    Write-Host "Re-launching in STA mode for WinForms compatibility..." -ForegroundColor Yellow
    $hostExe = (Get-Process -Id $PID).Path
    $args    = '-NoProfile -ExecutionPolicy Bypass -STA -File "' + $PSCommandPath + '"'
    if ($AcceptLicense) { $args += ' -AcceptLicense' }
    Start-Process -FilePath $hostExe -ArgumentList $args | Out-Null
    exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# --- Helper: always read files as UTF-8 correctly ---
function Get-FileUtf8 {
    param([string]$Path)
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    return [System.Text.Encoding]::UTF8.GetString($bytes)
}

function Get-LicenseText {
    $embedded = @"
SOFTWARE LICENSE AND TERMS OF USE
---------------------------------
Copyright (c) 2025 Benjamin Forrest
All Rights Reserved. Property of Benjamin Forrest.

Usage of this software is limited to the author (Benjamin Forrest)
and authorized personnel of Power Auto Group for internal use only.
Redistribution, modification, or alteration of this software in any form
is strictly prohibited without explicit written permission from the author.

By continuing, you acknowledge that:
  - You are an authorized user.
  - You agree not to distribute, modify, or reverse-engineer this software.
  - You understand this tool is provided as-is without warranty of any kind.
"@

    try {
        $here = Split-Path -Parent $PSCommandPath
        $licenseFile = Join-Path $here 'LICENSE.txt'
        if (Test-Path $licenseFile) {
            return [IO.File]::ReadAllText($licenseFile)
        }
    } catch { }
    return $embedded
}

function Show-LicenseWindow {
    param([string]$Text)

    # ---- Dialog ----
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "License Agreement"
    $dlg.StartPosition = 'CenterScreen'
    $dlg.Size = New-Object System.Drawing.Size(760, 560)
    $dlg.MinimizeBox = $false
    $dlg.MaximizeBox = $false
    $dlg.TopMost = $true
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.ShowInTaskbar = $true
    $dlg.KeyPreview = $true

    # ---- Header ----
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Please review and accept the license to continue."
    $lbl.AutoSize = $false
    $lbl.Dock = 'Top'
    $lbl.Height = 36
    $lbl.TextAlign = 'MiddleLeft'
    $lbl.Padding = '10,0,0,0'
    $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)

    # ---- Bottom panel (hosts checkbox + buttons) ----
    $bottom = New-Object System.Windows.Forms.Panel
    $bottom.Dock = 'Bottom'
    $bottom.Height = 76
    $bottom.Padding = '10,10,10,10'
    $bottom.BackColor = [System.Drawing.Color]::FromArgb(245,245,245)

    $chk = New-Object System.Windows.Forms.CheckBox
    $chk.Text = "I have read and agree to the License and Terms of Use."
    $chk.AutoSize = $true
    $chk.Location = New-Object System.Drawing.Point(10, 12)

    $btnAccept = New-Object System.Windows.Forms.Button
    $btnAccept.Text = "Accept"
    $btnAccept.Enabled = $false
    $btnAccept.Width = 100
    $btnAccept.Height = 30
    $btnAccept.Anchor = 'Top,Right'

    $btnDecline = New-Object System.Windows.Forms.Button
    $btnDecline.Text = "Decline"
    $btnDecline.Width = 100
    $btnDecline.Height = 30
    $btnDecline.Anchor = 'Top,Right'

    # Position buttons **relative to bottom panel**, not the dialog
    $buttonGap   = 10
    $rightMargin = 10
    $topMargin   = 35
    function Set-ButtonPositions {
        $declineX = $bottom.ClientSize.Width - $rightMargin - $btnDecline.Width
        $acceptX  = $declineX - $buttonGap - $btnAccept.Width
        $btnDecline.Location = New-Object System.Drawing.Point($declineX, $topMargin)
        $btnAccept.Location  = New-Object System.Drawing.Point($acceptX,  $topMargin)
    }
    $bottom.Add_Resize({ Set-ButtonPositions })

    # ---- Rich text (license body) ----
    $rtb = New-Object System.Windows.Forms.RichTextBox
    $rtb.ReadOnly = $true
    $rtb.Multiline = $true
    $rtb.ScrollBars = 'Vertical'
    $rtb.BorderStyle = 'FixedSingle'
    $rtb.DetectUrls = $false
    $rtb.Font = New-Object System.Drawing.Font("Consolas", 10)
    $rtb.Dock = 'Fill'
    $rtb.Text = $Text

    # ---- Wire events ----
    $chk.Add_CheckedChanged({ $btnAccept.Enabled = $chk.Checked })
    $btnAccept.Add_Click({ $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK; $dlg.Close() })
    $btnDecline.Add_Click({ $dlg.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $dlg.Close() })

    $dlg.Add_KeyDown({
        param($s,$e)
        if ($e.KeyCode -eq 'Escape') { $btnDecline.PerformClick() }
    })
    $dlg.AcceptButton = $btnAccept
    $dlg.CancelButton = $btnDecline

    # ---- Add controls in a safe dock order ----
    # Top, Bottom, then Fill to avoid overlap issues
    $dlg.SuspendLayout()
    $bottom.SuspendLayout()
    $bottom.Controls.Add($chk)
    $bottom.Controls.Add($btnAccept)
    $bottom.Controls.Add($btnDecline)

    $dlg.Controls.Add($rtb)     # Fill
    $dlg.Controls.Add($bottom)  # Bottom
    $dlg.Controls.Add($lbl)     # Top
    $dlg.ResumeLayout($false)
    $bottom.ResumeLayout($false)

    # Initial position for buttons
    Set-ButtonPositions

    # ---- Show modal ----
    $result = $dlg.ShowDialog()
    return ($result -eq [System.Windows.Forms.DialogResult]::OK)
}

# Gate execution unless bypassed for trusted automation
if (-not $AcceptLicense) {
    $ok = Show-LicenseWindow -Text (Get-LicenseText)
    if (-not $ok) { exit }
}
#endregion ===== END LICENSE / TERMS MODAL =====

# -------------------------------------------------------------
# Exporter: Export-PstContactsToOutlookCsv
# -------------------------------------------------------------
function Export-PstContactsToOutlookCsv {
    [CmdletBinding()] Param(
        [Parameter(Mandatory)] [string] $PstPath,
        [Parameter(Mandatory)] [string] $OutputCsv
    )

    # Validate paths
    if (-not (Test-Path -LiteralPath $PstPath)) {
        throw "PST not found: $PstPath"
    }
    $outDir = Split-Path -Parent $OutputCsv
    if (-not (Test-Path -LiteralPath $outDir)) {
        throw "Output folder does not exist: $outDir"
    }

    # Start Outlook COM
    try {
        $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
    } catch {
        $outlook = New-Object -ComObject Outlook.Application
    }
    if (-not $outlook) { throw 'Could not start or find Outlook.' }

    $session = $outlook.Session

    # Add the PST as a temporary store
    $storeRoot = $null
    $storeObj  = $null
    try {
        # AddStoreEx if available (safer) else fallback AddStore
        $addStoreEx = $session.GetType().GetMethod('AddStoreEx')
        if ($addStoreEx) { $session.AddStoreEx($PstPath, 1) } else { $session.AddStore($PstPath) }  # 1 = olStoreUnicode

        # Find the store we just added by matching FilePath (case-insensitive)
        foreach ($st in $session.Stores) {
            try {
                if ($st.FilePath -and ($st.FilePath -ieq $PstPath)) {
                    $storeObj  = $st
                    $storeRoot = $st.GetRootFolder()
                    break
                }
            } catch {}
        }
        if (-not $storeRoot) {
            # Fallback: pick the most recently added as the last in collection
            $storeObj  = $session.Stores | Select-Object -Last 1
            $storeRoot = $storeObj.GetRootFolder()
        }
        if (-not $storeRoot) { throw 'Unable to locate root folder for the added PST.' }

        # Traverse folders and collect contacts
        $contacts = New-Object System.Collections.Generic.List[Hashtable]

        # NOTE: No type annotation on $Folder to avoid __ComObject cast issues
        function Get-FoldersRecursive($Folder) {
            # DefaultItemType: 2 = olContactItem
            try {
                if ($Folder.DefaultItemType -eq 2) {
                    foreach ($item in $Folder.Items) {
                        if ($item -and $item.Class -eq 40) {  # 40 = olContact
                            # Map to Outlook-friendly CSV
                            $notes = ($item.Body -as [string])
                            if ($notes) { $notes = ($notes -replace '\r?\n',' ').Trim() }

                            $businessStreet = $item.BusinessAddressStreet
                            $businessCity   = $item.BusinessAddressCity
                            $businessState  = $item.BusinessAddressState
                            $businessPostal = $item.BusinessAddressPostalCode
                            $businessCountry= $item.BusinessAddressCountry

                            $homeStreet = $item.HomeAddressStreet
                            $homeCity   = $item.HomeAddressCity
                            $homeState  = $item.HomeAddressState
                            $homePostal = $item.HomeAddressPostalCode
                            $homeCountry= $item.HomeAddressCountry

                            $otherStreet = $item.OtherAddressStreet
                            $otherCity   = $item.OtherAddressCity
                            $otherState  = $item.OtherAddressState
                            $otherPostal = $item.OtherAddressPostalCode
                            $otherCountry= $item.OtherAddressCountry

                            # Avoid $HOME variable name collisions by using explicit names
                            $homeTel     = $item.HomeTelephoneNumber
                            $homeTel2    = $item.Home2TelephoneNumber
                            $businessTel = $item.BusinessTelephoneNumber
                            $businessTel2= $item.Business2TelephoneNumber
                            $mobileTel   = $item.MobileTelephoneNumber
                            $primaryTel  = $item.PrimaryTelephoneNumber
                            $assistantTel= $item.AssistantTelephoneNumber
                            $companyTel  = $item.CompanyMainTelephoneNumber
                            $carTel      = $item.CarTelephoneNumber
                            $otherTel    = $item.OtherTelephoneNumber
                            $pagerTel    = $item.PagerNumber
                            $isdnTel     = $item.ISDNNumber
                            $ttyTddTel   = $item.TTYTDDTelephoneNumber
                            $radioTel    = $item.RadioTelephoneNumber
                            $callbackTel = $item.CallbackTelephoneNumber
                            $telexTel    = $item.TelexNumber

                            $faxHome     = $item.HomeFaxNumber
                            $faxBusiness = $item.BusinessFaxNumber
                            $faxOther    = $item.OtherFaxNumber

                            # Prepare row with common Outlook CSV headers
                            $row = [ordered]@{
                                'Title'                       = $item.Title
                                'First Name'                  = $item.FirstName
                                'Middle Name'                 = $item.MiddleName
                                'Last Name'                   = $item.LastName
                                'Suffix'                      = $item.Suffix
                                'Company'                     = $item.CompanyName
                                'Department'                  = $item.Department
                                'Job Title'                   = $item.JobTitle
                                'E-mail Address'              = $item.Email1Address
                                'E-mail 2 Address'            = $item.Email2Address
                                'E-mail 3 Address'            = $item.Email3Address
                                'Business Street'             = $businessStreet
                                'Business City'               = $businessCity
                                'Business State'              = $businessState
                                'Business Postal Code'        = $businessPostal
                                'Business Country/Region'     = $businessCountry
                                'Home Street'                 = $homeStreet
                                'Home City'                   = $homeCity
                                'Home State'                  = $homeState
                                'Home Postal Code'            = $homePostal
                                'Home Country/Region'         = $homeCountry
                                'Other Street'                = $otherStreet
                                'Other City'                  = $otherCity
                                'Other State'                 = $otherState
                                'Other Postal Code'           = $otherPostal
                                'Other Country/Region'        = $otherCountry
                                'Business Phone'              = $businessTel
                                'Business Phone 2'            = $businessTel2
                                'Home Phone'                  = $homeTel
                                'Home Phone 2'                = $homeTel2
                                'Mobile Phone'                = $mobileTel
                                'Primary Phone'               = $primaryTel
                                'Assistant''s Phone'          = $assistantTel
                                'Company Main Phone'          = $companyTel
                                'Car Phone'                   = $carTel
                                'Other Phone'                 = $otherTel
                                'Pager'                        = $pagerTel
                                'ISDN'                         = $isdnTel
                                'TTY/TDD Phone'               = $ttyTddTel
                                'Radio Phone'                 = $radioTel
                                'Callback'                    = $callbackTel
                                'Telex'                       = $telexTel
                                'Home Fax'                    = $faxHome
                                'Business Fax'                = $faxBusiness
                                'Other Fax'                   = $faxOther
                                'Web Page'                    = $item.WebPage
                                'Birthday'                    = if ($item.Birthday) { $item.Birthday.ToString('yyyy-MM-dd') } else { $null }
                                'Categories'                  = $item.Categories
                                'Notes'                       = $notes
                            }
                            $contacts.Add($row) | Out-Null
                        }
                    }
                }
            } catch {
                Write-Output "[WARN] Skipping folder '$($Folder.Name)': $($_.Exception.Message)"
            }
            foreach ($sub in $Folder.Folders) { Get-FoldersRecursive -Folder $sub }
        }

        Get-FoldersRecursive -Folder $storeRoot

        # Write CSV (UTF-8 with BOM)
        $headers = @(
            'Title','First Name','Middle Name','Last Name','Suffix','Company','Department','Job Title',
            'E-mail Address','E-mail 2 Address','E-mail 3 Address',
            'Business Street','Business City','Business State','Business Postal Code','Business Country/Region',
            'Home Street','Home City','Home State','Home Postal Code','Home Country/Region',
            'Other Street','Other City','Other State','Other Postal Code','Other Country/Region',
            'Business Phone','Business Phone 2','Home Phone','Home Phone 2','Mobile Phone','Primary Phone',
            "Assistant's Phone",'Company Main Phone','Car Phone','Other Phone','Pager','ISDN','TTY/TDD Phone','Radio Phone','Callback','Telex',
            'Home Fax','Business Fax','Other Fax',
            'Web Page','Birthday','Categories','Notes'
        )

        $csvLines = New-Object System.Collections.Generic.List[string]
        # Header
        $csvLines.Add(($headers -join ',')) | Out-Null

        # Escape function for CSV fields
        function _esc([string]$s) {
            if ($null -eq $s) { return '' }
            $s = $s -replace '"','""'
            if ($s -match '[,\r\n"]') { return '"' + $s + '"' } else { return $s }
        }

        foreach ($c in $contacts) {
            $rowVals = foreach ($h in $headers) { _esc ($c[$h]) }
            $csvLines.Add(($rowVals -join ',')) | Out-Null
        }

        $utf8bom = New-Object System.Text.UTF8Encoding($true)
        [System.IO.File]::WriteAllLines($OutputCsv, $csvLines, $utf8bom)

        Write-Output "Exported $($contacts.Count) contact(s) to $OutputCsv"
    }
    finally {
        # Remove the store if we successfully added it
        try {
            if ($session -and $storeRoot) { $session.RemoveStore($storeRoot) }
        } catch {}
        # Release COM
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($storeRoot)  | Out-Null } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($storeObj)    | Out-Null } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($session)     | Out-Null } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook)     | Out-Null } catch {}
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

# -------------------------------------------------------------
# GUI
# -------------------------------------------------------------
# --- Helper: logging to the UI ---
function Write-Log {
    param(
        [string]$Message,
        [Parameter()] [object]$Color = 'Black'
    )
    # Convert Color if needed (accepts names like 'Gray', 'DodgerBlue', etc.)
    if ($Color -isnot [System.Drawing.Color]) {
        try { $Color = [System.Drawing.Color]::FromName([string]$Color) } catch { $Color = [System.Drawing.Color]::Black }
    }
    $timestamp = (Get-Date).ToString('HH:mm:ss')
    $line = "[$timestamp] $Message"
    $logBox.SelectionStart = $logBox.TextLength
    $logBox.SelectionLength = 0
    $logBox.SelectionColor = $Color
    $logBox.AppendText("$line`r`n")
    $logBox.SelectionColor = $logBox.ForeColor
    $logBox.ScrollToCaret()
}

$form = New-Object System.Windows.Forms.Form
$Form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon((Resolve-Path ".\app.ico"))
$form.Text = 'PST Contacts Extractor (Standalone)'
$form.StartPosition = 'CenterScreen'
$form.MinimumSize = New-Object System.Drawing.Size(720, 480)
$form.MaximizeBox = $true
$form.AutoSize = $true
$form.AutoSizeMode = 'GrowAndShrink'


$lblPst = New-Object System.Windows.Forms.Label
$lblPst.Text = 'PST file:'
$lblPst.Location = New-Object System.Drawing.Point(15,20)
$lblPst.AutoSize = $true

$txtPst = New-Object System.Windows.Forms.TextBox
$txtPst.Location = New-Object System.Drawing.Point(90,16)
$txtPst.Width = 480
$txtPst.ReadOnly = $true

$btnPst = New-Object System.Windows.Forms.Button
$btnPst.Text = 'Browse...'
$btnPst.Location = New-Object System.Drawing.Point(590,14)
$btnPst.Width = 90

$lblCsv = New-Object System.Windows.Forms.Label
$lblCsv.Text = 'Save CSV as:'
$lblCsv.Location = New-Object System.Drawing.Point(15,60)
$lblCsv.AutoSize = $true

$txtCsv = New-Object System.Windows.Forms.TextBox
$txtCsv.Location = New-Object System.Drawing.Point(90,56)
$txtCsv.Width = 480
$txtCsv.ReadOnly = $true

$btnCsv = New-Object System.Windows.Forms.Button
$btnCsv.Text = 'Choose...'
$btnCsv.Location = New-Object System.Drawing.Point(590,54)
$btnCsv.Width = 90

$logBox = New-Object System.Windows.Forms.RichTextBox
$logBox.Location = New-Object System.Drawing.Point(18,100)
$logBox.Size = New-Object System.Drawing.Size(662,220)
$logBox.ReadOnly = $true
$logBox.BackColor = [System.Drawing.Color]::White
$logBox.BorderStyle = 'FixedSingle'
$logBox.WordWrap = $true
$logBox.ScrollBars = 'Vertical'

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = 'Extract Contacts'
$btnRun.Location = New-Object System.Drawing.Point(18,335)
$btnRun.Width = 170
$btnRun.Enabled = $false

$btnOpenOut = New-Object System.Windows.Forms.Button
$btnOpenOut.Text = 'Open Output Folder'
$btnOpenOut.Location = New-Object System.Drawing.Point(200,335)
$btnOpenOut.Width = 170
$btnOpenOut.Enabled = $false

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(592,335)
$btnClose.Width = 88

$pstDialog = New-Object System.Windows.Forms.OpenFileDialog
$pstDialog.Filter = 'Outlook PST (*.pst)|*.pst|All files (*.*)|*.*'
$pstDialog.Title  = 'Select a PST file'

$csvDialog = New-Object System.Windows.Forms.SaveFileDialog
$csvDialog.Filter = 'Outlook CSV (*.csv)|*.csv'
$csvDialog.Title  = 'Choose where to save the CSV'
$csvDialog.AddExtension = $true
$csvDialog.DefaultExt   = 'csv'
$csvDialog.FileName = 'contacts_outlook365.csv'

# ===== Footer Label =====
$COPY = [char]0x00A9
$footerLabel = New-Object System.Windows.Forms.Label
$footerLabel.AutoSize = $false
$footerLabel.Dock = [System.Windows.Forms.DockStyle]::Bottom
$footerLabel.Height = 40
$footerLabel.TextAlign = 'MiddleCenter'
$footerLabel.BackColor = [System.Drawing.Color]::FromArgb(245,245,245)
$footerLabel.ForeColor = [System.Drawing.Color]::Gray
$footerLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
$footerLabel.Text = "$COPY 2025 Benjamin Forrest | All rights reserved | Internal Use Only | Redistribution or modification prohibited."

$form.Controls.Add($footerLabel)

function Update-RunEnabled {
    $btnRun.Enabled = ([string]::IsNullOrWhiteSpace($txtPst.Text) -eq $false) -and `
                      ([string]::IsNullOrWhiteSpace($txtCsv.Text) -eq $false)
}

$btnPst.Add_Click({
    if ($pstDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtPst.Text = $pstDialog.FileName
        Write-Log -Message "Selected PST: $($txtPst.Text)" -Color 'Gray'
        Update-RunEnabled
    }
})

$btnCsv.Add_Click({
    if ($csvDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtCsv.Text = $csvDialog.FileName
        Write-Log -Message "CSV will be saved to: $($txtCsv.Text)" -Color 'Gray'
        $btnOpenOut.Enabled = Test-Path (Split-Path -Parent $txtCsv.Text)
        Update-RunEnabled
    }
})

$btnOpenOut.Add_Click({
    if (-not [string]::IsNullOrWhiteSpace($txtCsv.Text)) {
        $outDir = Split-Path -Parent $txtCsv.Text
        if (Test-Path $outDir) { Start-Process explorer.exe $outDir }
    }
})

$btnClose.Add_Click({ $form.Close() })

$btnRun.Add_Click({
    try {
        if ([string]::IsNullOrWhiteSpace($txtPst.Text)) { throw 'Please select a PST file.' }
        if ([string]::IsNullOrWhiteSpace($txtCsv.Text)) { throw 'Please choose a CSV save location.' }

        $pstPath = $txtPst.Text
        $csvPath = $txtCsv.Text

        if (-not (Test-Path $pstPath)) { throw "PST not found: $pstPath" }
        $csvDir = Split-Path -Parent $csvPath
        if (-not (Test-Path $csvDir)) { throw "Output folder does not exist: $csvDir" }

        # UI state
        $form.UseWaitCursor = $true
        $btnRun.Enabled = $false
        Write-Log -Message "Starting export..." -Color 'DodgerBlue'
        Write-Log -Message " PST: $pstPath"
        Write-Log -Message " CSV: $csvPath"

        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $err = $null
        try {
            $output = & Export-PstContactsToOutlookCsv -PstPath $pstPath -OutputCsv $csvPath -ErrorAction Stop *>&1
            foreach ($o in $output) {
                if ($o -is [System.Management.Automation.ErrorRecord]) {
                    Write-Log -Message ($o.ToString()) -Color 'Crimson'
                } elseif ($o -is [System.Management.Automation.VerboseRecord]) {
                    Write-Log -Message ($o.ToString()) -Color 'Gray'
                } else {
                    Write-Log -Message (($o | Out-String).Trim())
                }
            }
        } catch {
            $err = $_
        }
        $sw.Stop()

        if ($err) {
            Write-Log -Message "Export failed: $($err.Exception.Message)" -Color 'Crimson'
        } else {
            Write-Log -Message ("Export completed in {0}s" -f [math]::Round($sw.Elapsed.TotalSeconds,2)) -Color 'ForestGreen'
            if (Test-Path $csvPath) {
                $btnOpenOut.Enabled = $true
                Write-Log -Message "Saved: $csvPath" -Color 'ForestGreen'
            }
        }
    }
    finally {
        $form.UseWaitCursor = $false
        Update-RunEnabled
    }
})

$form.Controls.AddRange(@(
    $lblPst, $txtPst, $btnPst,
    $lblCsv, $txtCsv, $btnCsv,
    $logBox,
    $btnRun, $btnOpenOut, $btnClose
))

Write-Log -Message 'Ready. Select a PST and choose where to save the CSV.' -Color 'Gray'
Update-RunEnabled
[void]$form.ShowDialog()
