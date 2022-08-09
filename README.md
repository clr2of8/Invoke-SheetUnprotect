# Invoke-SheetUnprotect

This PowerShell script will change the password for any protected sheets in a Microsoft Excel file to the simple string **xyz** for easy recovery.

```powershell
Import-module .\Invoke-SheetUnprotect.ps1 #run this from the directory where you download the Invoke-SheetUnprotect.ps1 file
Invoke-SheetUnprotect C:\Users\art\Protected.xlsx # look for another copy of the file in the same directory with **-xyz** appended to it
```
