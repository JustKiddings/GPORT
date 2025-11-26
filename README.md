# GPORT
GPORT (GPO + Import) is a PowerShell script that automates GPO importing from backup.

## Usage
```powershell
.\hardening.ps1 -h

Usage: .\hardening.ps1 [-Msg] [-Rename] [-OU] [-Help]

Parameters:
  -Msg
     Prompts for Legal Notice title and text; adds/updates them in GptTmpl.inf under [Registry Values].
  -Rename
     Prompts for new Administrator and Guest account names; adds/updates them under [System Access].
  -OU
     Prompts for OU name, creates it in AD, and links the imported GPO.
  -Help / -h
     Displays this help screen.
```

## Important


### On GPOs on Windows Server
