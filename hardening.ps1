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
  -Msg     Prompts for Legal Notice title and text.
  -Rename  Prompts for new Administrator and Guest account names.
  -OU      Creates OU and links the GPO.
"@ | Write-Host
    exit
}

# Current working folder
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

    # Check existence
    $existing = Get-GPO -Name $gpoName -ErrorAction SilentlyContinue

    if ($existing) {
        Write-Host "⚠️  WARNING: GPO '$gpoName' already exists." -ForegroundColor Yellow
        $ans = Read-Host "Are you sure you want to OVERWRITE it? (Y/N)"
        if ($ans -match '^[Yy]') {
            $proceedWithGPO = $true
            Write-Host "Selected existing GPO: $gpoName" -ForegroundColor Cyan
        }
        # If N, loop repeats
    } else {
        # Create immediately
        try {
            New-GPO -Name $gpoName -Comment "Imported by script" | Out-Null
            Write-Host "✅ Created new GPO: $gpoName" -ForegroundColor Green
            $proceedWithGPO = $true
        } catch {
            Write-Error "Failed to create GPO: $_"
            # If creation fails, we might want to exit or retry
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
# BACKUP FILE DISCOVERY
# ==========================================

$gptFile = Get-ChildItem -Path $currentDir -Recurse -Filter "GptTmpl.inf" -ErrorAction SilentlyContinue | Select-Object -First 1
$FilePath = $null

if ($gptFile) {
    $FilePath = $gptFile.FullName
    $info = Get-BackupRootAndIdFromFile $FilePath
    $backupRoot = $info.BackupRoot
    $backupId   = $info.BackupId
    Write-Host "Detected Backup Source: $backupId"
} else {
    Write-Host "❌ No GptTmpl.inf found in this folder." -ForegroundColor Red
    Write-Host "   Cannot import GPO settings without the backup files."
    exit 1
}

# ==========================================
# MODIFY BACKUP FILES (If requested)
# ==========================================

if ($Msg) {
    Write-Host "📝 Updating legal notice..."
    Set-RegistryValueInINF -File $FilePath -KeyPattern "MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System\LegalNoticeCaption" -NewValue (Read-Host "Legal Title") -DefaultType 1
    Set-RegistryValueInINF -File $FilePath -KeyPattern "MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System\LegalNoticeText" -NewValue (Read-Host "Legal Text") -DefaultType 7
}

if ($Rename) {
    Write-Host "📝 Updating account names..."
    $lines = Get-Content -Path $FilePath -Encoding Unicode
    $list = [System.Collections.Generic.List[string]]::new()
    $list.AddRange([string[]]$lines)
    $sysIndex = -1
    for ($i = 0; $i -lt $list.Count; $i++) { if ($list[$i].Trim() -eq "[System Access]") { $sysIndex = $i; break } }

    if ($sysIndex -ge 0) {
        Ensure-SystemAccessLine -List $list -SysIndex $sysIndex -Key "NewAdministratorName" -Value (Read-Host "New Admin Name")
        Ensure-SystemAccessLine -List $list -SysIndex $sysIndex -Key "NewGuestName" -Value (Read-Host "New Guest Name")
        Set-Content -Path $FilePath -Value $list -Encoding Unicode
    } else { Write-Error "❌ [System Access] section not found in INF file." }
}

# ==========================================
# IMPORT SETTINGS TO GPO
# ==========================================

Write-Host "Importing settings into '$gpoName'..."
try {
    # We use CreateIfNeeded $true just in case, but we handled creation in Step 1
    Import-GPO -BackupId ([guid]$backupId) -Path $backupRoot -TargetName $gpoName -CreateIfNeeded $true -ErrorAction Stop | Out-Null
    Write-Host "✅ Import successfully completed." -ForegroundColor Green
} catch {
    Write-Error "Import failed: $_"
    exit 1
}

# ==========================================
# OU CREATION & LINKING
# ==========================================

if ($OU) {
    # Check for AD module availability
    try { Get-ADDomain -ErrorAction SilentlyContinue | Out-Null }
    catch { Write-Error "Active Directory module or connection is missing."; exit 1 }

    $ouName = Read-Host "Enter OU name to create"
    $domainDN = (Get-ADDomain).DistinguishedName
    $ouDN = "OU=$ouName,$domainDN"

    if (-not (Get-ADOrganizationalUnit -LDAPFilter "(distinguishedName=$ouDN)" -ErrorAction SilentlyContinue)) {
        New-ADOrganizationalUnit -Name $ouName -Path $domainDN
        Write-Host "✅ Created OU: $ouDN"
    } else { Write-Host "OU already exists: $ouDN" }

    Write-Host "Ready to link GPO '$gpoName' to OU '$ouDN'."
    $linkConfirm = Read-Host "Link now? (Y/N)"

    if ($linkConfirm -match '^[Yy]') {
        New-GPLink -Name $gpoName -Target $ouDN -Enforced No | Out-Null
        Write-Host "✅ GPO Linked." -ForegroundColor Green
    } else {
        Write-Host "⏭️  Skipping Link."
    }
}
