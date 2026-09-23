Get Win Prod. Key

```bash
Delay 2000
GuiR
WinPrintLine powershell
Delay 1000

WinPrintLine (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform").BackupProductKeyDefault
```