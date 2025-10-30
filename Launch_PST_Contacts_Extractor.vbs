' ---------------------------------------------------------------
' Launch_PST_Contacts_Extractor.vbs
' Silent launcher for the PST Contacts Extractor GUI
' Copyright (c) 2025 Benjamin Forrest - All Rights Reserved
' ---------------------------------------------------------------

Option Explicit

Dim shell, scriptDir, ps1Path, cmd
Set shell = CreateObject("WScript.Shell")

' Get this VBS file’s directory
scriptDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

' Path to your PowerShell GUI script
ps1Path = """" & scriptDir & "\PST_Contacts_Extractor_GUI-1.0.0.ps1" & """"

' Build command to run PowerShell silently
cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File " & ps1Path

' Run it with windowstyle=0 (hidden), do not wait for completion (False)
shell.Run cmd, 0, False
