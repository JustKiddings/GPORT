param(
    [switch]$Msg,
    [switch]$Rename,
    [switch]$OU,
    [switch]$Help,
    [switch]$h
)

# Show help if requested
if ($Help -or $h) {
@"
Usage: .\hardening.ps1 [-Msg] [-Rename] [-OU] [-Help]

Parameters:
  -Msg     Prompts for title and text shown before log in.
  -Rename  Prompts for new Administrator and Guest account names.
  -OU      Creates OU and links the GPO.
"@ | Write-Host
    exit
}

$currentDir = (Get-Location).ProviderPath
Write-Host "Working folder: $currentDir"

# ==========================================
# GPO NAME CHECK & CREATION
# ==========================================

$gpoName = $null
$proceedWithGPO = $false

do {
    Write-Host ""
    $gpoName = Read-Host 'Name the GPO'

    if ([string]::IsNullOrWhiteSpace($gpoName)) { continue }

    $existing = Get-GPO -Name $gpoName -ErrorAction SilentlyContinue

    if ($existing) {
        Write-Host "⚠️  WARNING: GPO '$gpoName' already exists." -ForegroundColor Yellow
        $ans = Read-Host "Are you sure you want to OVERWRITE it? (Y/N)"
        if ($ans -match '^[Yy]') {
            $proceedWithGPO = $true
            Write-Host "Selected existing GPO: $gpoName" -ForegroundColor Cyan
        }
    } else {
        try {
            New-GPO -Name $gpoName -Comment "Imported by GPORT script" | Out-Null
            Write-Host "✅ Created new GPO: $gpoName" -ForegroundColor Green
            $proceedWithGPO = $true
        } catch {
            Write-Error "Failed to create GPO: $_"
            exit 1
        }
    }
} until ($proceedWithGPO)

# ==========================================
# HELPER FUNCTIONS
# ==========================================

function Get-BackupRootAndIdFromFile {
    param([string]$fileFullPath)
    $dir = Split-Path -Parent $fileFullPath
    while ($dir -and ($dir -ne [IO.Path]::GetPathRoot($dir))) {
        $name = Split-Path -Leaf $dir
        if ($name -match '^\{?[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}\}?$') {
            return @{ BackupRoot = (Split-Path -Parent $dir); BackupId = $name.Trim('{}') }
        }
        $dir = Split-Path -Parent $dir
    }
    return $null
}

function Set-RegistryValueInINF {
    param([string]$File, [string]$KeyPattern, [string]$NewValue, [int]$DefaultType = 1)
    $lines = Get-Content -Path $File -Encoding Unicode
    $list = [System.Collections.Generic.List[string]]::new()
    $list.AddRange([string[]]$lines)

    $sectionIndex = -1
    for ($i = 0; $i -lt $list.Count; $i++) { if ($list[$i].Trim() -eq "[Registry Values]") { $sectionIndex = $i; break } }
    if ($sectionIndex -lt 0) { $list.Add("[Registry Values]"); $sectionIndex = $list.Count - 1 }

    $escapedKey = [regex]::Escape($KeyPattern)
    $keyRegex = "^$escapedKey\s*=\s*(\d+),.*$"
    $existingIndex = -1
    for ($i = $sectionIndex + 1; $i -lt $list.Count; $i++) {
        if ($list[$i].Trim() -match '^\[.*\]$') { break }
        if ($list[$i] -match $keyRegex) { $existingIndex = $i; break }
    }

    $newLine = "$KeyPattern=$DefaultType,`"$NewValue`""
    if ($existingIndex -ge 0) { $list[$existingIndex] = $newLine } else { $list.Insert($sectionIndex + 1, $newLine) }
    Set-Content -Path $File -Value $list -Encoding Unicode
}

function Ensure-SystemAccessLine {
    param([System.Collections.Generic.List[string]]$List, [int]$SysIndex, [string]$Key, [string]$Value)
    $pattern = "^$Key\s*="
    $idx = -1
    for ($i = $SysIndex + 1; $i -lt $List.Count; $i++) {
        if ($List[$i].Trim() -match '^\[.*\]$') { break }
        if ($List[$i] -match $pattern) { $idx = $i; break }
    }
    if ($idx -ge 0) { $List[$idx] = "$Key = `"$Value`"" } else { $List.Insert($SysIndex + 1, "$Key = `"$Value`"") }
}

# ==========================================
# LOCATE BACKUP & PREPARE SAFETY NET
# ==========================================

$gptFile = Get-ChildItem -Path $currentDir -Recurse -Filter "GptTmpl.inf" -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $gptFile) {
    Write-Host "❌ No GptTmpl.inf found. Cannot import." -ForegroundColor Red
    exit 1
}

$FilePath = $gptFile.FullName
$backupFile = "$FilePath.bak"
$info = Get-BackupRootAndIdFromFile $FilePath
$backupRoot = $info.BackupRoot
$backupId   = $info.BackupId

# [CRASH RECOVERY]
# If a .bak exists from a previous run that crashed, restore it NOW to ensure we start clean.
if (Test-Path $backupFile) {
    Write-Warning "Found leftover backup from a previous crashed run. Restoring clean INF..."
    Move-Item -Path $backupFile -Destination $FilePath -Force
}

# [CREATE SNAPSHOT]
# Create a fresh backup of the clean state before we touch anything.
Copy-Item -Path $FilePath -Destination $backupFile -Force

# ==========================================
# MODIFY & IMPORT (Wrapped in Try/Finally)
# ==========================================

try {
    # --- Modification Logic ---
    if ($Msg) {
        Write-Host "📝 Updating legal notice..."
        $cap = Read-Host "Legal Title"
        $txt = Read-Host "Legal Text"
        Set-RegistryValueInINF -File $FilePath -KeyPattern "MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System\LegalNoticeCaption" -NewValue $cap -DefaultType 1
        Set-RegistryValueInINF -File $FilePath -KeyPattern "MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System\LegalNoticeText" -NewValue $txt -DefaultType 7
    }

    if ($Rename) {
        Write-Host "📝 Updating account names..."
        $adm = Read-Host "New Admin Name"
        $gst = Read-Host "New Guest Name"

        $lines = Get-Content -Path $FilePath -Encoding Unicode
        $list = [System.Collections.Generic.List[string]]::new()
        $list.AddRange([string[]]$lines)
        $sysIndex = -1
        for ($i = 0; $i -lt $list.Count; $i++) { if ($list[$i].Trim() -eq "[System Access]") { $sysIndex = $i; break } }

        if ($sysIndex -ge 0) {
            Ensure-SystemAccessLine -List $list -SysIndex $sysIndex -Key "NewAdministratorName" -Value $adm
            Ensure-SystemAccessLine -List $list -SysIndex $sysIndex -Key "NewGuestName" -Value $gst
            Set-Content -Path $FilePath -Value $list -Encoding Unicode
        } else {
            Write-Warning "Cannot rename accounts: [System Access] section missing."
        }
    }

    # --- Import Logic ---
    Write-Host "Importing settings into '$gpoName'..."
    Import-GPO -BackupId ([guid]$backupId) -Path $backupRoot -TargetName $gpoName -CreateIfNeeded $true -ErrorAction Stop | Out-Null
    Write-Host "✅ Import successfully completed." -ForegroundColor Green

}
catch {
    Write-Error "An error occurred: $_"
    # We do NOT exit here, we let it fall through to 'Finally' to clean up.
}
finally {
    # ==========================================
    # CLEANUP / RESTORE
    # ==========================================
    # This runs whether the script Succeeded OR Failed.
    if (Test-Path $backupFile) {
        Write-Host "🧹 Restoring original GptTmpl.inf..." -ForegroundColor DarkGray
        Move-Item -Path $backupFile -Destination $FilePath -Force
    }
}

# ==========================================
# OU CREATION & LINKING
# ==========================================

if ($OU) {
    $ouName = Read-Host "Enter OU name to create/link"
    $domainDN = (Get-ADDomain).DistinguishedName
    $ouDN = "OU=$ouName,$domainDN"

    $ouExists = Get-ADOrganizationalUnit -LDAPFilter "(distinguishedName=$ouDN)" -ErrorAction SilentlyContinue

    if (-not $ouExists) {
        # --- NEW OU (Auto-Link) ---
        Write-Host "Creating new OU: $ouDN"
        New-ADOrganizationalUnit -Name $ouName -Path $domainDN
        Write-Host "✅ OU Created."

        Write-Host "Linking GPO to the new OU..."
        New-GPLink -Name $gpoName -Target $ouDN -Enforced No | Out-Null
        Write-Host "✅ GPO Linked automatically." -ForegroundColor Green
    }
    else {
        # --- EXISTING OU (Ask Confirmation) ---
        Write-Host "⚠️  OU already exists: $ouDN" -ForegroundColor Yellow
        $linkConfirm = Read-Host "Do you want to link the GPO to this EXISTING OU? (Y/N)"

        if ($linkConfirm -match '^[Yy]') {
            New-GPLink -Name $gpoName -Target $ouDN -Enforced No | Out-Null
            Write-Host "✅ GPO Linked." -ForegroundColor Green
        } else {
            Write-Host "⏭️  Skipping Link."
        }
    }
}

