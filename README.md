# GPORT

**GPORT** automates Windows machine hardening by importing Group Policy Object (GPO) backup derived from hardening suggestions in the **[CIS Microsoft Windows 10 Enterprise Benchmark (v.4.0.0)](https://learn.cisecurity.org/benchmarks) document**.

## Quick Start

<<<<<<< HEAD
### Prerequisites
* **OS:** At least one domain joined Windows 10/11 PC and Windows Server running as DC.
* **Permissions:** Domain Administrator or Delegated GPO/OU access.
* **Files:** A valid GPO backup folder must be in the script directory.

### Usage
Run the script to import hardening settings.
=======
## Coverage & Compatibility

### Windows 10
* **Without Parameters:** The script covers **99%** of the CIS Microsoft Windows 10 Enterprise Benchmark (v.4.0.0) L1 level.
* **With Parameters:** When run with `-Msg` and `-Rename`, the script achieves **100%** coverage.

### Windows 11
* **Note:** Windows 11 support is **not proper** at this time.
* The current GPO Backup covers approximately **86%** of the **CIS Microsoft Windows 11 Enterprise Benchmark (v4.0.0) L1 level**.

## Reports
You can find the detailed compliance and validation reports here:
[to fill]

## Prerequisites

* **OS:** Domain joined Windows 10/11 Workstation and Windows Server.
* **Permissions:** Domain Administrator or Delegated GPO/OU permissions.
* **Files:** A valid GPO backup folder (containing `GptTmpl.inf`) must exist in the same directory as the script.

## Usage
>>>>>>> 5ea2229429c9ddee831e74de88b8914ec44fa8fb

```powershell
.\hardening.ps1 -h

Usage: .\hardening.ps1 [-Msg] [-Rename] [-OU] [-Help/-h]

Parameters:
  -Msg          Prompts for title and text shown before log in.
  -Rename       Prompts for new Administrator and Guest account names.
  -OU           Creates OU and links the GPO to OU.
  -Help / -h    Shows this help section.
````

## Coverage of CIS Benchmarks

  * **Windows 10 Enterprise L1:** 99% base coverage. **100%** coverage when using `-Msg` and `-Rename`.
  * **Windows 11 Enterprise L1:** \~86% coverage (Experimental).

[View Compliance Reports](to fill)

## How It Works

1.  **GPO Initialization:** Prompts for a GPO name. If it exists, requests confirmation to overwrite; otherwise, creates a new GPO.
2.  **Atomic Modification:**
    * **Safety Net:** Detects previous crashes and restores the environment, then creates a temporary backup of `GptTmpl.inf`.
    * **Injection:** Directly modifies the INF file to inject parameters (Legal Notice, Account Renames) if flags are set.
3.  **Import:** Pushes the (potentially modified) configuration into Active Directory.
4.  **Auto-Restoration:** Uses a `try/finally` block to **always** restore the original `GptTmpl.inf` from the backup, ensuring the source files remain clean for future runs.
5.  **Deployment (Optional):** If `-OU` is selected:
    * **New OU:** Creates the OU and automatically links the imported GPO.
    * **Existing OU:** Detects the OU and prompts for confirmation before linking.

## Troubleshooting

If some policies are not visible in **Group Policy Management Editor**, you are missing the latest ADMX/ADML templates in the Central Store.

**Fix:**

1.  **Download:**
      * [Administrative Templates (Windows 11)](https://www.microsoft.com/en-us/download/details.aspx?id=108394)
      * [Security Baselines](https://www.microsoft.com/en-us/download/details.aspx?id=55319)
2.  **Create Central Store:** Ensure the following folder exists on your DC (if PolicyDefinitions doesn't exist create it):
    `\\yourdomain.local\SYSVOL\yourdomain.local\Policies\PolicyDefinitions\`
3.  **Install:** Copy `.admx` files and `en-US` folders from the downloads into `PolicyDefinitions`.
4.  **Verify:** Open GPMC; you should now see "MS Security Guide" and "MSS (legacy)" under Administrative Templates.

## Limitations

* **Single Backup Context:** If multiple GPO backup folders exist in the script directory, it blindly processes the first one found.
* **OS Support:** Fully validated for **Windows 10** only. Windows 11 support is currently experimental.
* **Environment:** Must be executed on a **Domain Controller**.
* **Object Management:** The script creates OUs and links GPOs but does **not** move or add computers/users to the new OU.
