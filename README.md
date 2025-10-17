# WinAttend ISO Packer

A simple and lightweight PowerShell GUI app to repack Windows ISOs with your `autounattend.xml` (and optional `$OEM$`). 

Features Dark/light UI, progressreporting, and generates a bootable ISO via `oscdimg.exe`.

> **Requires Windows, PowerShell 5.1+ and Administrator rights.**.

---

## Features
- WPF GUI (dark/light), uses windows native file picker  
- Injects `autounattend.xml`, optionally `$OEM$`  
- Shows progress + saves logs in `%TEMP%\WinRepacker\logs`  
- Output: `<Original>.custom.iso` next to the app  
- Finds `oscdimg.exe` (embedded/temp, side-by-side, PATH, or ADK)


## Quick Start (GUI)
1. Run `Windows ISO Re-Packer.exe`
2. Pick **Windows ISO** and **autounattend.xml**
3. (Optional) Check **Include `$OEM$`**
4. Click **Start**

> If the `answerfiles` directory exists next to the app, the picker auto-opens there.


## CLI
> **Requires you to adjust ExecutionPolicy**
```powershell
# Basic
.\Windows_ISO-Gen.ps1 -IsoPath "C:\ISOs\Win11.iso" -UnattendXmlPath ".\answerfiles\autounattend.xml"

# Include $OEM$
.\Windows_ISO-Gen.ps1 -IsoPath "C:\ISOs\Win11.iso" -UnattendXmlPath ".\answerfiles\autounattend.xml" -IncludeOEM

# Force theme
.\Windows_ISO-Gen.ps1 -ForceDark
.\Windows_ISO-Gen.ps1 -ForceLight
