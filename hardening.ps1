param(
    [switch]$Msg,
    [switch]$Rename,
    [switch]$OU,
    [switch]$Help,
    [switch]$h
)

# [ENCODING FIX]
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Show help if requested
if ($Help -or $h) {
@"
Usage: .\hardening.ps1 [-Msg] [-Rename] [-OU] [-Help]

Parameters:
  -Msg          Prompts for title and text shown before log in.
  -Rename       Prompts for new Administrator and Guest account names.
  -OU           Creates OU and links the GPO to OU.
  -Help / -h    Shows this help section.
"@ | Write-Host
    exit
}

$currentDir = (Get-Location).ProviderPath
Write-Host "Working folder: $currentDir"

# Define Emojis (ASCII-Safe)
$eCheck   = [char]0x2705                       # ✅
$eWarn    = [char]0x26A0                       # ⚠️
$eCross   = [char]0x274C                       # ❌
$eSkip    = [char]0x23ED                       # ⏭️
$eMemo    = [char]::ConvertFromUtf32(0x1F4DD)  # 📝
$eBroom   = [char]::ConvertFromUtf32(0x1F9F9)  # 🧹
$eSearch  = [char]::ConvertFromUtf32(0x1F50D)  # 🔍

# ==========================================
# HELPER FUNCTIONS
# ==========================================

function Get-GpoDisplayNameFromBackup {
    param([string]$backupXmlPath)
    try {
        [xml]$xml = Get-Content -Path $backupXmlPath -ErrorAction Stop
        $name = $xml.GroupPolicyBackupScheme.GroupPolicyObject.GroupPolicyCoreSettings.DisplayName
        if (-not [string]::IsNullOrWhiteSpace($name.'#cdata-section')) { return $name.'#cdata-section' }
        if (-not [string]::IsNullOrWhiteSpace($name)) { return $name }
    } catch {}
    return "Unknown-GPO"
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
# DISCOVER & SELECT GPO BACKUP
# ==========================================

Write-Host "`n$eSearch Scanning for GPO backups..."
$backupXmlFiles = Get-ChildItem -Path $currentDir -Recurse -Filter "Backup.xml" -ErrorAction SilentlyContinue

if (-not $backupXmlFiles) {
    Write-Host "$eCross No 'Backup.xml' files found in subdirectories." -ForegroundColor Red
    exit 1
}

$availableBackups = @()
foreach ($file in $backupXmlFiles) {
    $dir = $file.Directory
    if ($dir.Name -match '^\{?[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}\}?$') {
         $displayName = Get-GpoDisplayNameFromBackup -backupXmlPath $file.FullName
         $availableBackups += [PSCustomObject]@{
            Index = $availableBackups.Count + 1
            DisplayName = $displayName
            BackupId = $dir.Name.Trim('{}')
            BackupRoot = $dir.Parent.FullName
            Directory = $dir.FullName
         }
    }
}

if ($availableBackups.Count -eq 0) {
    Write-Host "$eCross No valid GPO backup folders (GUID-named) found." -ForegroundColor Red
    exit 1
}

Write-Host "Found the following GPO backups:" -ForegroundColor Cyan
$availableBackups | Format-Table -Property Index, DisplayName, BackupId -AutoSize

$selectedBackup = $null
do {
    $selectedIndex = Read-Host "Select ONE backup index to import"
    $selectedBackup = $availableBackups | Where-Object { $_.Index -eq $selectedIndex }
    if (-not $selectedBackup) {
        Write-Warning "Invalid index. Please try again."
    }
} while (-not $selectedBackup)

# ==========================================
# GPO NAME CHECK & CREATION
# ==========================================

$gpoName = $null
$proceedWithGPO = $false
do {
    Write-Host ""
    $defaultName = $selectedBackup.DisplayName
    $gpoName = Read-Host "Name the GPO (Press Enter for '$defaultName')"
    if ([string]::IsNullOrWhiteSpace($gpoName)) { $gpoName = $defaultName }

    $existing = Get-GPO -Name $gpoName -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "$eWarn WARNING: GPO '$gpoName' already exists." -ForegroundColor Yellow
        $ans = Read-Host "Are you sure you want to OVERWRITE it? (Y/N)"
        if ($ans -match '^[Yy]') {
            $proceedWithGPO = $true
            Write-Host "Selected existing GPO: $gpoName" -ForegroundColor Cyan
        }
    } else {
        try {
            New-GPO -Name $gpoName -Comment "Imported by GPORT from backup $($selectedBackup.BackupId)" | Out-Null
            Write-Host "$eCheck Created new GPO: $gpoName" -ForegroundColor Green
            $proceedWithGPO = $true
        } catch {
            Write-Error "Failed to create GPO: $_"
            exit 1
        }
    }
} while (-not $proceedWithGPO)


# ==========================================
# LOCATE BACKUP & PREPARE SAFETY NET
# ==========================================

$gptFile = Get-ChildItem -Path $selectedBackup.Directory -Recurse -Filter "GptTmpl.inf" -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $gptFile) {
    Write-Host "$eCross No GptTmpl.inf found in selected backup. Cannot import." -ForegroundColor Red
    exit 1
}

$FilePath = $gptFile.FullName
$backupFile = "$FilePath.bak"
$backupRoot = $selectedBackup.BackupRoot
$backupId   = $selectedBackup.BackupId

if (Test-Path $backupFile) {
    Write-Warning "Found leftover backup from a previous crashed run. Restoring clean INF..."
    Move-Item -Path $backupFile -Destination $FilePath -Force
}

Copy-Item -Path $FilePath -Destination $backupFile -Force

# ==========================================
# MODIFY & IMPORT (Wrapped in Try/Finally)
# ==========================================

try {
    if ($Msg) {
        Write-Host "$eMemo Updating legal notice..."
        $cap = Read-Host "Legal Title"
        $txt = Read-Host "Legal Text"
        Set-RegistryValueInINF -File $FilePath -KeyPattern "MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System\LegalNoticeCaption" -NewValue $cap -DefaultType 1
        Set-RegistryValueInINF -File $FilePath -KeyPattern "MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System\LegalNoticeText" -NewValue $txt -DefaultType 7
    }

    if ($Rename) {
        Write-Host "$eMemo Updating account names..."
        $adm = Read-Host "New Admin Name"
        $gst = Read-Host "New Guest Name"
        $lines = Get-Content -Path $FilePath -Encoding Unicode
        $list = [System.Collections.Generic.List[string]]::new(); $list.AddRange([string[]]$lines)
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

    Write-Host "Importing settings into '$gpoName'..."
    Import-GPO -BackupId ([guid]$backupId) -Path $backupRoot -TargetName $gpoName -ErrorAction Stop | Out-Null
    Write-Host "$eCheck Import successfully completed." -ForegroundColor Green
}
catch {
    Write-Error "An error occurred: $_"
}
finally {
    if (Test-Path $backupFile) {
        Write-Host "$eBroom Restoring original GptTmpl.inf..." -ForegroundColor DarkGray
        Move-Item -Path $backupFile -Destination $FilePath -Force
    }
}

# ==========================================
# OU CREATION & LINKING
# ==========================================

if ($OU) {
    $ouName = Read-Host "Enter OU name to create"
    $domainDN = (Get-ADDomain).DistinguishedName
    $ouDN = "OU=$ouName,$domainDN"

    $ouExists = Get-ADOrganizationalUnit -LDAPFilter "(distinguishedName=$ouDN)" -ErrorAction SilentlyContinue

    if (-not $ouExists) {
        Write-Host "Creating new OU: $ouDN"
        New-ADOrganizationalUnit -Name $ouName -Path $domainDN
        Write-Host "$eCheck OU Created."

        Write-Host "Linking GPO to the new OU..."
        New-GPLink -Name $gpoName -Target $ouDN -Enforced No | Out-Null
        Write-Host "$eCheck GPO Linked automatically." -ForegroundColor Green
    }
    else {
        Write-Host "$eWarn  OU already exists: $ouDN" -ForegroundColor Yellow
        $linkConfirm = Read-Host "Do you want to link the GPO to this EXISTING OU? (Y/N)"
        if ($linkConfirm -match '^[Yy]') {
            New-GPLink -Name $gpoName -Target $ouDN -Enforced No | Out-Null
            Write-Host "$eCheck GPO Linked." -ForegroundColor Green
        } else {
            Write-Host "$eSkip  Skipping Link."
        }
    }
}
