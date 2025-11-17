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
Usage: .\gpt.ps1 [-Msg] [-Rename] [-OU] [-Help]

Parameters:
  -Msg
     Prompts for Legal Notice title and text; adds/updates them in GptTmpl.inf under [Registry Values].
  -Rename
     Prompts for new Administrator and Guest account names; adds/updates them under [System Access].
  -OU
     Prompts for OU name, creates it in AD, and links the imported GPO.
  -Help / -h
     Displays this help screen.
"@ | Write-Host
    exit
}

# Current working folder
$currentDir = (Get-Location).ProviderPath
Write-Host "Working folder: $currentDir"

# Ask for GPO name
$gpoName = Read-Host 'Name the GPO'

# Only search for GptTmpl.inf if -Msg or -Rename used
# --- Only attempt INF modifications if -Msg or -Rename ---
$FilePath = $null


# Function to find GPO backup root & GUID
function Get-BackupRootAndIdFromFile {
    param([string]$fileFullPath)
    $dir = Split-Path -Parent $fileFullPath
    while ($dir -and ($dir -ne [IO.Path]::GetPathRoot($dir))) {
        $name = Split-Path -Leaf $dir
        if ($name -match '^\{?[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}\}?$') {
            $guidRaw = $name.Trim('{}')
            $backupRoot = Split-Path -Parent $dir
            return @{ BackupRoot = $backupRoot; BackupId = $guidRaw }
        }
        $dir = Split-Path -Parent $dir
    }
    return $null
}


if ($Msg -or $Rename) {
    $gptFile = Get-ChildItem -Path $currentDir -Recurse -Filter "GptTmpl.inf" -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $gptFile) {
        if ($Msg) { Write-Host "❌ No GptTmpl.inf found under $currentDir. Cannot modify legal notice." }
        if ($Rename) { Write-Host "❌ No GptTmpl.inf found under $currentDir. Cannot rename accounts." }
        exit 1
    }
    Write-Host "✅ Found GptTmpl.inf at: $($gptFile.FullName)"
    $FilePath = $gptFile.FullName
    # Only call this function if $FilePath exists
    $info = Get-BackupRootAndIdFromFile $FilePath
    $backupRoot = $info.BackupRoot
    $backupId   = $info.BackupId
    Write-Host "Detected backup root: $backupRoot"
    Write-Host "Detected backup ID : $backupId"
} else {
    # No INF modifications, but we still need $backupRoot/$backupId for Import-GPO
    # In that case, search for GPO backup folder (GUID folder) in current directory
    $gptFile = Get-ChildItem -Path $currentDir -Recurse -Filter "GptTmpl.inf" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($gptFile) {
        $info = Get-BackupRootAndIdFromFile $gptFile.FullName
        Write-Host $info
        $backupRoot = $info.BackupRoot
        $backupId   = $info.BackupId
        Write-Host "Detected backup root: $backupRoot"
        Write-Host "Detected backup ID : $backupId"
    } else {
        Write-Host "⚠️ No GptTmpl.inf found, skipping INF modifications. Importing GPO requires a valid backup folder."
        exit 1
    }
}

# --- Functions to modify INF ---
function Set-RegistryValueInINF {
    param(
        [string]$File,
        [string]$KeyPattern,
        [string]$NewValue,
        [int]$DefaultType = 1
    )

    # Read file with UTF-16LE encoding
    $lines = Get-Content -Path $File -Encoding Unicode
    $list = [System.Collections.Generic.List[string]]::new()
    $list.AddRange([string[]]$lines)

    # Find or create [Registry Values] section
    $sectionIndex = -1
    for ($i = 0; $i -lt $list.Count; $i++) {
        if ($list[$i].Trim() -eq "[Registry Values]") {
            $sectionIndex = $i
            break
        }
    }

    if ($sectionIndex -lt 0) {
        $list.Add("[Registry Values]")
        $sectionIndex = $list.Count - 1
    }

    # Escape special regex characters in the key pattern
    $escapedKey = [regex]::Escape($KeyPattern)
    $keyRegex = "^$escapedKey\s*=\s*(\d+),.*$"

    # Find existing line with this key
    $existingIndex = -1
    for ($i = $sectionIndex + 1; $i -lt $list.Count; $i++) {
        # Stop if we hit another section
        if ($list[$i].Trim() -match '^\[.*\]$') {
            break
        }
        if ($list[$i] -match $keyRegex) {
            $existingIndex = $i
            break
        }
    }

    $newLine = "$KeyPattern=$DefaultType,`"$NewValue`""

    if ($existingIndex -ge 0) {
        # Update existing line
        Write-Host "  Updating existing line at index $existingIndex"
        $list[$existingIndex] = $newLine
    } else {
        # Insert new line after section header
        Write-Host "  Inserting new line after section header at index $($sectionIndex + 1)"
        $list.Insert($sectionIndex + 1, $newLine)
    }

    # Write back with UTF-16LE encoding
    Set-Content -Path $File -Value $list -Encoding Unicode
}

function Ensure-SystemAccessLine {
    param(
        [System.Collections.Generic.List[string]]$List,
        [int]$SysIndex,
        [string]$Key,
        [string]$Value
    )

    # Build pattern to match the key
    $pattern = "^$Key\s*="

    # Search for existing line in [System Access] section
    $idx = -1
    for ($i = $SysIndex + 1; $i -lt $List.Count; $i++) {
        # Stop if we hit another section
        if ($List[$i].Trim() -match '^\[.*\]$') {
            break
        }
        if ($List[$i] -match $pattern) {
            $idx = $i
            break
        }
    }

    $newLine = "$Key = `"$Value`""

    if ($idx -ge 0) {
        # Update existing line
        Write-Host "  Updating existing line: $Key"
        $List[$idx] = $newLine
    } else {
        # Insert after section header
        Write-Host "  Inserting new line: $Key"
        $List.Insert($SysIndex + 1, $newLine)
    }
}

# --- Update Legal Notice if requested ---
if ($Msg) {
    Write-Host "📝 Updating legal notice..."
    $caption = Read-Host "Enter LegalNoticeCaption text (Title)"
    $text    = Read-Host "Enter LegalNoticeText (Body)"

    Set-RegistryValueInINF -File $FilePath `
        -KeyPattern "MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System\LegalNoticeCaption" `
        -NewValue $caption -DefaultType 1

    Set-RegistryValueInINF -File $FilePath `
        -KeyPattern "MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System\LegalNoticeText" `
        -NewValue $text -DefaultType 7

    Write-Host "✅ Legal notice updated (UTF-16LE preserved)."
}

# --- Rename accounts if requested ---
if ($Rename) {
    Write-Host "📝 Rename mode: Updating Administrator and Guest accounts."
    $newAdmin = Read-Host "Enter new Administrator account name"
    $newGuest = Read-Host "Enter new Guest account name"

    # Read file with UTF-16LE encoding
    $lines = Get-Content -Path $FilePath -Encoding Unicode
    $list = [System.Collections.Generic.List[string]]::new()
    $list.AddRange([string[]]$lines)

    # Find [System Access] section
    $sysIndex = -1
    for ($i = 0; $i -lt $list.Count; $i++) {
        if ($list[$i].Trim() -eq "[System Access]") {
            $sysIndex = $i
            break
        }
    }

    if ($sysIndex -lt 0) {
        Write-Host "❌ [System Access] section not found"
        exit 1
    }

    Ensure-SystemAccessLine -List $list -SysIndex $sysIndex -Key "NewAdministratorName" -Value $newAdmin
    Ensure-SystemAccessLine -List $list -SysIndex $sysIndex -Key "NewGuestName" -Value $newGuest

    # Write back with UTF-16LE encoding
    Set-Content -Path $FilePath -Value $list -Encoding Unicode
    Write-Host "✅ Account rename applied (UTF-16LE preserved)."
}

# --- Create or get GPO ---
try {
    $existing = Get-GPO -Name $gpoName -ErrorAction SilentlyContinue
    if (-not $existing) {
        Write-Host "Creating GPO '$gpoName'..."
        New-GPO -Name $gpoName -Comment "Imported by script"
    } else { Write-Host "GPO already exists." }
} catch {
    Write-Error "Failed to create/find GPO: $_"
    throw
}


# --- Import GPO from backup ---
try {
    $guidObj = [guid]$backupId
} catch {
    Write-Error "BackupId invalid: $backupId"; exit 1
}

$params = @{
    BackupId   = $guidObj
    Path       = $backupRoot
    TargetName = $gpoName
    CreateIfNeeded = $true
}

Write-Host "Importing GPO from backup..."
try {
    Import-GPO @params -Verbose
    Write-Host "✅ Import completed."
} catch {
    Write-Error "Import-GPO failed: $($_.Exception.Message)"
    throw
}

# --- Create OU and link GPO if requested ---
if ($OU) {
    Import-Module ActiveDirectory -ErrorAction Stop
    Import-Module GroupPolicy -ErrorAction Stop

    $ouName = Read-Host "Enter OU name to create"
    $domainDN = (Get-ADDomain).DistinguishedName
    $ouDN = "OU=$ouName,$domainDN"

    if (-not (Get-ADOrganizationalUnit -LDAPFilter "(distinguishedName=$ouDN)" -ErrorAction SilentlyContinue)) {
        Write-Host "Creating OU: $ouDN"
        New-ADOrganizationalUnit -Name $ouName -Path $domainDN
    } else { Write-Host "OU already exists: $ouDN" }

    Write-Host "Linking GPO '$gpoName' to $ouDN..."
    New-GPLink -Name $gpoName -Target $ouDN -Enforced No
    Write-Host "✅ GPO linked to OU."
}
