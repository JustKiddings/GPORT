# GPORT

**GPORT** (GPO + Import) is a PowerShell script designed to automate the hardening of Windows machines by importing Group Policy Object (GPO) backups.

The hardening settings imported by this script are derived from the **CIS Microsoft Windows 10 Enterprise Benchmark (v.4.0.0) - Level 1**. You can obtain the official benchmark document from [CIS Security Benchmarks](https://learn.cisecurity.org/benchmarks).

## Prerequisites

* **OS:** Windows 10/11 or Windows Server.
* **Permissions:** Domain Administrator or Delegated GPO/OU permissions.
* **Files:** A valid GPO backup folder (containing `GptTmpl.inf`) must exist in the same directory as the script.

## Usage

```powershell
.\hardening.ps1 -h

Usage: .\hardening.ps1 [-Msg] [-Rename] [-OU] [-Help]

Parameters:
  -Msg         Prompts for title and text shown before log in.
  -Rename      Prompts for new Administrator and Guest account names.
  -OU          Creates OU and links the GPO to OU.
  -Help / -h   Displays this help screen.
```

## How It Works

1.  **Safety Check:** Checks if the GPO exists. If it does, it warns that importing will **overwrite all settings**.
2.  **Backup & Modify:** If parameters are used, it creates a temporary backup of `GptTmpl.inf`, modifies the original with your inputs, and prepares it for import.
3.  **Atomic Import:** The settings are imported into Active Directory.
4.  **Auto-Cleanup:** The script automatically restores the original clean `GptTmpl.inf` from the backup and removes any incomplete "zombie" GPOs if the script crashes.

---

## Important: Visibility & ADMX Templates

### Missing Group Policies in Editor
After running the script, you may notice that **not all Group Policies appear** when viewing the GPO in the **Group Policy Management Editor**.

According to the CIS Benchmark documentation, visibility of certain updated policies requires the latest versions of **ADMX/ADML templates** installed in a **Central Store**. Without these templates, the settings are applied to the machines, but you cannot view or edit them in the GPMC.

### Solution: Updating ADMX/ADML Files
To fix this, you must create a Central Store and populate it with the latest Administrative Templates and Security Baselines.

#### 1. Download Templates
* **Administrative Templates (Windows 11 25H2):** [Download here](https://www.microsoft.com/en-us/download/details.aspx?id=108394)
    * Run the MSI installer. By default, files are extracted to:
      `C:\Program Files (x86)\Microsoft Group Policy\Windows 11 Sep 2025 Update (25H2)\PolicyDefinitions`
* **Security Baselines:** [Download here](https://www.microsoft.com/en-us/download/details.aspx?id=55319)
    * Download `Windows 11 v25H2 Security Baseline.zip`.
    * Extract the zip and locate the `Templates` folder.

#### 2. Create the Central Store
The Central Store is not created by default. You must create it in the SYSVOL folder on your Domain Controller.

1.  Find your domain DNS root:
    ```powershell
    Get-ADDomain | Select-Object DNSRoot
    ```
2.  Navigate to the following path (replace `yourdomain.local` with your actual domain):
    `\\yourdomain.local\SYSVOL\yourdomain.local\Policies\`
3.  Create a new folder named **`PolicyDefinitions`**.

#### 3. Copy Files
You must copy files from the downloaded locations (Step 1) into the new `PolicyDefinitions` folder (Step 2).

* **ADMX Files:** Copy all `.admx` files directly into the root of `PolicyDefinitions`.
* **ADML Files:** Copy the `en-US` folder (or your specific region folder) containing `.adml` files into `PolicyDefinitions`.

**Structure Example:**
```text
\\yourdomain.local\SYSVOL\yourdomain.local\Policies\
└── PolicyDefinitions\
    ├── en-US\
    │   ├── *.adml
    │   └── ...
    ├── *.admx
    └── ...
```

#### 4. Verification
1.  Open **Group Policy Management**.
2.  Right-click the GPO and select **Edit**.
3.  Navigate to **Computer Configuration > Policies > Administrative Templates**.
4.  You should now see **"MS Security Guide"** and **"MSS (Legacy)"**, confirming the templates are loaded from the Central Store.
