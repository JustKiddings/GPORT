# GPORT

**GPORT** automates Windows machine hardening by importing Group Policy Object (GPO) backup derived from hardening suggestions in the **[CIS Microsoft Windows 10/11 Enterprise Benchmark (v.4.0.0)](https://learn.cisecurity.org/benchmarks) document**.


## Quick Start

### Prerequisites
* **OS:** At least one domain joined Windows 10/11 PC and Windows Server running as DC.
* **Permissions:** Domain Administrator or Delegated GPO/OU access.
* **Files:** A valid GPO backup folder must be in the script directory.

### To Get Started

1.  **Download:** Get the GPORT.zip from the [latest release](https://github.com/JustKiddings/GPORT/releases/) and extract its contents to a folder.
2.  **Open PowerShell:** Open a PowerShell window in the extracted folder. Ensure the GPO backup folder (e.g., `{C38F5F91-39F5-468B-8520-1EC31A282424}`) and `hardening.ps1` script are present in the working directory.

### Usage
Run the script to import hardening settings.

```powershell

.\hardening.ps1 -h

Usage: .\hardening.ps1 [-Msg] [-Rename] [-OU] [-Help/-h]

Parameters:

  -Msg          Prompts for title and text shown before log in.
  -Rename       Prompts for new Administrator and Guest account names.
  -OU           Creates OU and links the GPO to OU.
  -Help / -h    Shows this help section.
````
## Coverage & Compatibility

### Windows 10
* **GPO Backup: CIS-W10-L1/{C38F5F91-39F5-468B-8520-1EC31A282424}**
* **Without Parameters:** The script covers **99%** of the CIS Microsoft Windows 10 Enterprise Benchmark (v.4.0.0) L1 level.
* **With Parameters:** When run with `-Msg` and `-Rename`, the script achieves **100%** coverage.

### Windows 11
* **GPO Backup: CIS-W11-L1/{4F41DBA2-1D39-44FD-B09A-BC20565286A5}**
* **Without Parameters:** The script covers **99%** of the CIS Microsoft Windows 11 Enterprise Benchmark (v.4.0.0) L1 level.
* **With Parameters:** When run with `-Msg` and `-Rename`, the script achieves **100%** coverage. 

## Reports
[View All CIS-CAT Lite Audit Reports](https://justkiddings.github.io/GPORT/)

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

* **Environment:** Must be executed on a **Domain Controller**.
* **Object Management:** The script creates OUs and links GPOs but does **not** move or add computers/users to the new OU.
