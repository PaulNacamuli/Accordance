<#
.SYNOPSIS
    Migrates Accordance Desktop user data, preferences, and modules from an old PC to a new PC.

.DESCRIPTION
    Implements the Windows half of the vendor migration article:
    https://support.accordancebible.com/hc/en-us/articles/35804210381851-Migrating-Accordance-Desktop-to-Another-Computer

    The article names four locations on Windows. This script handles them as four named items:

        Modules      C:\ProgramData\Accordance        modules and support files
        Preferences  %LOCALAPPDATA%\Accordance        holds "Accordance Preferences"
        UserFiles    <Documents>\Accordance Files     workspaces, highlights, user notes, user tools
        Application  C:\Program Files (x86)\Oaktree   the program itself, opt in only

    Requires PowerShell 5.1 - do not use PS 6+ features.

    Five modes:

        Inventory   Reports which items exist on this PC, their file counts and sizes, and the
                    staging space needed. Copies nothing. Run this first on the old PC.
        Export      Copies the items found on this PC into a staging folder (USB drive, external
                    disk, or network share) and writes manifest.json describing what was captured.
        Compare     Run on the new PC against a staging folder. For each item, reports how many
                    files are identical, differ, exist only in the export, or exist only on this
                    PC. Changes nothing. Worth running before Import whenever the new PC already
                    holds a populated Accordance install, because a fresh install that has already
                    downloaded modules can hold a larger and newer library than the old PC, and
                    importing over it would be a downgrade.
        Merge       Adds only the files the export has and this PC does not. Nothing existing is
                    renamed, replaced, or deleted. This is the answer when neither machine's module
                    library is a superset of the other: an Import would gain the old PC's extra
                    modules at the cost of burying the new PC's, and a skip does the reverse.
        Import      Reads manifest.json from the staging folder and copies each item into the
                    matching location on this PC. Any existing target folder is renamed to
                    <name>.bak-yyyyMMddHHmmss first, which is the article's "remove or rename the
                    existing folder" step. Renaming requires -Force so an accidental run cannot
                    disturb a working install.

    Import replaces a whole folder; Merge only ever adds to one. Use Import for preferences and
    user content, where the old PC's version is the one you want throughout. Use Merge for modules,
    where both machines may hold something the other lacks.

    Application is skipped unless -IncludeApplication is passed, because a fresh install of
    Accordance on the new PC supplies the program files already. Copying the old ones over a
    newer install is a downgrade, not a migration.

    Paths are resolved from the live environment on each machine rather than read back from the
    manifest, so the two PCs can have different user names and the Documents folder can be
    redirected into OneDrive on either side.

    Copying is done by robocopy so that large module libraries and long paths are handled
    predictably. Every item is verified after copying by comparing file count and total bytes
    against the source. A log is written beside the staged data.

.PARAMETER Mode
    Inventory, Export, Compare, Merge, or Import.

.PARAMETER Path
    Staging folder. Required for every mode except Inventory. An "AccordanceMigration" subfolder is
    created under it on Export and looked for by the other modes, so the same -Path value works
    throughout.

.PARAMETER ReportPath
    Compare only. Folder to write per-item CSVs listing every file that differs or exists on only
    one side. Without it, Compare prints counts and a sample to the console and writes nothing.

.PARAMETER Hash
    Compare only. SHA256 the files that match in size but not in write time, to say for certain
    whether the content is the same. Without it those files are counted separately and left
    undecided, because a differing timestamp on its own proves nothing: two machines that installed
    the same module at different times produce exactly that. Slower, and worth it before deciding
    what to do about a module library.

.PARAMETER IncludeApplication
    Also handle C:\Program Files (x86)\Oaktree. Needs an elevated session on Import.

.PARAMETER Force
    On Import, allows an existing target folder to be renamed to <name>.bak-yyyyMMddHHmmss before
    the staged copy is placed. Without it, an item whose target already has content is skipped and
    reported. On Export, allows writing into a staging folder that already holds an export.

.PARAMETER SkipItem
    One or more item names to leave alone: Modules, Preferences, UserFiles, Application.

.EXAMPLE
    .\Migrate-Accordance.ps1 -Mode Inventory

    Old PC, step one. Reports what is present and how much staging space an export needs.

.EXAMPLE
    .\Migrate-Accordance.ps1 -Mode Export -Path E:\

    Old PC, step two. Stages everything found into E:\AccordanceMigration.

.EXAMPLE
    .\Migrate-Accordance.ps1 -Mode Compare -Path E:\ -ReportPath C:\Temp\AccordanceCompare

    New PC, before importing. Shows per item what the export holds that this PC does not, and what
    this PC holds that the export does not, with the full detail in CSVs.

.EXAMPLE
    .\Migrate-Accordance.ps1 -Mode Import -Path E:\ -Force

    New PC, after Accordance has been installed, launched once, and closed. Renames the freshly
    created Accordance folders aside and puts the old PC's data in place.

.EXAMPLE
    .\Migrate-Accordance.ps1 -Mode Import -Path E:\ -Force -SkipItem Modules

    The usual shape when the new PC's install has already downloaded its own module library:
    migrate preferences and user content, leave the newer modules alone.

.EXAMPLE
    .\Migrate-Accordance.ps1 -Mode Merge -Path E:\ -SkipItem Preferences,UserFiles

    Then, elevated, add just the modules the old PC had and this one does not, keeping every module
    this PC already holds.

.EXAMPLE
    .\Migrate-Accordance.ps1 -Mode Import -Path E:\ -WhatIf

    Shows every rename and copy the import would perform, without touching anything.

.NOTES
    Author:  Paul Nacamuli
    Version: 1.00
    Date:    09/11/2026
    Source:  Accordance

    Close Accordance on both machines before running. The script refuses to continue while
    Accordance.exe is running, because preferences and module indexes are written on exit.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateSet('Inventory', 'Export', 'Compare', 'Merge', 'Import')]
    [string]$Mode,

    [Parameter(Position = 1)]
    [string]$Path,

    [string]$ReportPath,

    [switch]$Hash,

    [switch]$IncludeApplication,

    [switch]$Force,

    [ValidateSet('Modules', 'Preferences', 'UserFiles', 'Application')]
    [string[]]$SkipItem
)
$scriptVer = "1.00" # Paul Nacamuli 09/11/2026
$scriptName = "Migrate-Accordance.ps1"
$stageFolderName = "AccordanceMigration"
$manifestName = "manifest.json"
$stamp = Get-Date -Format "yyyyMMddHHmmss"
$script:LogFile = $null
$script:Problems = @()

Write-Host ""
Write-Host "$scriptName v$scriptVer - Accordance Desktop migration - mode: $Mode" -ForegroundColor Cyan
Write-Host ""

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Good', 'Warn', 'Bad', 'Head')]
        [string]$Level = 'Info'
    )
    $colors = @{ Info = 'Gray'; Good = 'Green'; Warn = 'Yellow'; Bad = 'Red'; Head = 'White' }
    Write-Host $Message -ForegroundColor $colors[$Level]
    if ($script:LogFile) {
        $line = "{0} [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level.ToUpper(), $Message
        # -WhatIf:$false keeps a dry run from narrating every log append. $script:LogFile is left
        # unset during -WhatIf anyway, so a dry run writes no file at all.
        try { Add-Content -LiteralPath $script:LogFile -Value $line -Encoding ASCII -WhatIf:$false -ErrorAction Stop } catch { }
    }
}

function Add-Problem {
    param([string]$Message)
    $script:Problems += $Message
    Write-Log $Message -Level Bad
}

function Format-Size {
    param([double]$Bytes)
    if ($Bytes -ge 1GB) { return ("{0:N2} GB" -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ("{0:N2} MB" -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ("{0:N2} KB" -f ($Bytes / 1KB)) }
    return ("{0} bytes" -f [int]$Bytes)
}

function Test-IsElevated {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($id)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-FolderStat {
    # PS 5.1: a lone PSCustomObject has a $null .Count, and Measure-Object Sum is $null over an
    # empty set. Both are normalised here so callers can compare numbers without guarding.
    param([string]$FolderPath)
    $files = @()
    if ($FolderPath -and (Test-Path -LiteralPath $FolderPath)) {
        $files = @(Get-ChildItem -LiteralPath $FolderPath -Recurse -File -Force -ErrorAction SilentlyContinue)
    }
    $sum = ($files | Measure-Object -Property Length -Sum).Sum
    if (-not $sum) { $sum = 0 }
    return [pscustomobject]@{
        Files = $files.Count
        Bytes = [int64]$sum
    }
}

function Get-DocumentsCandidate {
    # Documents can be the plain profile folder or redirected into a OneDrive root whose name
    # carries the tenant, so every plausible root is returned and the caller picks.
    $list = @()
    $myDocs = [Environment]::GetFolderPath('MyDocuments')
    if ($myDocs) { $list += $myDocs }
    $list += (Join-Path $env:USERPROFILE 'Documents')
    foreach ($root in @($env:OneDrive, $env:OneDriveCommercial, $env:OneDriveConsumer)) {
        if ($root) { $list += (Join-Path $root 'Documents') }
    }
    $oneDriveDirs = @(Get-ChildItem -LiteralPath $env:USERPROFILE -Directory -Filter 'OneDrive*' -ErrorAction SilentlyContinue)
    foreach ($dir in $oneDriveDirs) { $list += (Join-Path $dir.FullName 'Documents') }
    return @($list | Where-Object { $_ } | Select-Object -Unique)
}

function Get-AccordanceItem {
    # One row per location named in the vendor article. Source is where the data sits on this PC
    # right now, empty when absent. Target is where an import should put it on this PC.
    $items = @()

    $programDataAccordance = Join-Path $env:ProgramData 'Accordance'
    $items += [pscustomobject]@{
        Name        = 'Modules'
        Description = 'Modules and support files'
        Source      = if (Test-Path -LiteralPath $programDataAccordance) { $programDataAccordance } else { '' }
        Target      = $programDataAccordance
        Optional    = $false
        NeedsAdmin  = $true
    }

    $localAccordance = Join-Path $env:LOCALAPPDATA 'Accordance'
    $items += [pscustomobject]@{
        Name        = 'Preferences'
        Description = 'Accordance Preferences'
        Source      = if (Test-Path -LiteralPath $localAccordance) { $localAccordance } else { '' }
        Target      = $localAccordance
        Optional    = $false
        NeedsAdmin  = $false
    }

    $docCandidates = Get-DocumentsCandidate
    $userFilesSource = ''
    foreach ($candidate in $docCandidates) {
        $probe = Join-Path $candidate 'Accordance Files'
        if (Test-Path -LiteralPath $probe) { $userFilesSource = $probe; break }
    }
    $docRoot = $docCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
    if (-not $docRoot) { $docRoot = Join-Path $env:USERPROFILE 'Documents' }
    # When an "Accordance Files" folder already exists, that is the folder to target. On the new PC
    # a first launch creates it wherever Windows currently points Documents, and the staged copy
    # belongs there rather than in whichever candidate root happens to sort first.
    $userFilesTarget = Join-Path $docRoot 'Accordance Files'
    if ($userFilesSource) { $userFilesTarget = $userFilesSource }
    $items += [pscustomobject]@{
        Name        = 'UserFiles'
        Description = 'Accordance Files - workspaces, highlights, user notes, user tools'
        Source      = $userFilesSource
        Target      = $userFilesTarget
        Optional    = $false
        NeedsAdmin  = $false
    }

    $appCandidates = @()
    if (${env:ProgramFiles(x86)}) { $appCandidates += (Join-Path ${env:ProgramFiles(x86)} 'Oaktree') }
    $appCandidates += (Join-Path $env:ProgramFiles 'Oaktree')
    $appSource = $appCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
    if (-not $appSource) { $appSource = '' }
    $items += [pscustomobject]@{
        Name        = 'Application'
        Description = 'Oaktree program folder - normally supplied by a fresh install instead'
        Source      = $appSource
        Target      = $appCandidates[0]
        Optional    = $true
        NeedsAdmin  = $true
    }

    return $items
}

function Select-MigrationItem {
    # Applies -SkipItem and the Application opt in, and explains every exclusion.
    param([object[]]$AllItems)
    $selected = @()
    foreach ($item in $AllItems) {
        if ($SkipItem -and ($SkipItem -contains $item.Name)) {
            Write-Log ("  {0,-12} skipped by -SkipItem" -f $item.Name) -Level Warn
            continue
        }
        if ($item.Name -eq 'Application' -and -not $IncludeApplication) {
            Write-Log ("  {0,-12} skipped - pass -IncludeApplication to migrate the program folder" -f $item.Name) -Level Info
            continue
        }
        $selected += $item
    }
    return @($selected)
}

function Get-FileMap {
    # Relative path (lower case, for a case-insensitive filesystem) to FileInfo, so two trees can
    # be compared by where files sit inside them rather than by absolute path.
    param([string]$Root)
    $map = @{}
    if (-not $Root -or -not (Test-Path -LiteralPath $Root)) { return $map }
    $rootFull = (Get-Item -LiteralPath $Root).FullName.TrimEnd('\')
    $prefix = $rootFull.Length + 1
    foreach ($file in (Get-ChildItem -LiteralPath $Root -Recurse -File -Force -ErrorAction SilentlyContinue)) {
        if ($file.FullName.Length -le $prefix) { continue }
        $map[$file.FullName.Substring($prefix).ToLowerInvariant()] = $file
    }
    return $map
}

function Compare-Tree {
    # A same-sized file whose write time differs is the normal case for two machines that installed
    # the same module separately, so it is counted apart from a genuine size difference. Timestamps
    # alone say nothing about content. -Hash settles those cases for certain.
    param(
        [string]$Label,
        [string]$StagedRoot,
        [string]$LocalRoot,
        [switch]$UseHash
    )
    $staged = Get-FileMap -Root $StagedRoot
    $local = Get-FileMap -Root $LocalRoot
    $rows = @()
    $identical = 0
    $sameSize = 0
    $diffSize = 0
    $diffContent = 0
    $onlyStaged = 0
    $onlyLocal = 0
    $onlyStagedBytes = [int64]0
    $onlyLocalBytes = [int64]0
    $hashed = 0

    foreach ($key in $staged.Keys) {
        $s = $staged[$key]
        if ($local.ContainsKey($key)) {
            $l = $local[$key]
            $relative = $s.FullName.Substring($s.FullName.Length - $key.Length)
            $newer = 'old PC'
            if ($l.LastWriteTimeUtc -gt $s.LastWriteTimeUtc) { $newer = 'this PC' }

            if ($s.Length -ne $l.Length) {
                $diffSize++
                $rows += [pscustomobject]@{
                    Item          = $Label
                    Status        = 'DifferentSize'
                    Newer         = $newer
                    RelativePath  = $relative
                    OldPCBytes    = $s.Length
                    ThisPCBytes   = $l.Length
                    OldPCWritten  = $s.LastWriteTime
                    ThisPCWritten = $l.LastWriteTime
                }
                continue
            }

            $gap = [math]::Abs(($s.LastWriteTimeUtc - $l.LastWriteTimeUtc).TotalSeconds)
            if ($gap -le 2) {
                $identical++
                continue
            }

            if (-not $UseHash) {
                # Counted, deliberately not listed: on a real module library this is thousands of
                # rows of noise that would bury the handful of findings that matter.
                $sameSize++
                continue
            }

            $hashed++
            if (($hashed % 250) -eq 0) {
                Write-Progress -Activity "Hashing $Label" -Status "$hashed same-sized files compared" -Id 1
            }
            $sh = $null
            $lh = $null
            try {
                $sh = (Get-FileHash -LiteralPath $s.FullName -Algorithm SHA256 -ErrorAction Stop).Hash
                $lh = (Get-FileHash -LiteralPath $l.FullName -Algorithm SHA256 -ErrorAction Stop).Hash
            } catch { }
            if ($sh -and $lh -and $sh -eq $lh) {
                $identical++
                continue
            }
            $diffContent++
            $rows += [pscustomobject]@{
                Item          = $Label
                Status        = 'DifferentContent'
                Newer         = $newer
                RelativePath  = $relative
                OldPCBytes    = $s.Length
                ThisPCBytes   = $l.Length
                OldPCWritten  = $s.LastWriteTime
                ThisPCWritten = $l.LastWriteTime
            }
        } else {
            $onlyStaged++
            $onlyStagedBytes += $s.Length
            $rows += [pscustomobject]@{
                Item         = $Label
                Status       = 'OnlyOnOldPC'
                Newer        = ''
                RelativePath = $s.FullName.Substring($s.FullName.Length - $key.Length)
                OldPCBytes   = $s.Length
                ThisPCBytes  = ''
                OldPCWritten = $s.LastWriteTime
                ThisPCWritten = ''
            }
        }
    }
    foreach ($key in $local.Keys) {
        if ($staged.ContainsKey($key)) { continue }
        $l = $local[$key]
        $onlyLocal++
        $onlyLocalBytes += $l.Length
        $rows += [pscustomobject]@{
            Item         = $Label
            Status       = 'OnlyOnThisPC'
            Newer        = ''
            RelativePath = $l.FullName.Substring($l.FullName.Length - $key.Length)
            OldPCBytes   = ''
            ThisPCBytes  = $l.Length
            OldPCWritten = ''
            ThisPCWritten = $l.LastWriteTime
        }
    }

    if ($hashed -gt 0) { Write-Progress -Activity "Hashing $Label" -Id 1 -Completed }

    return [pscustomobject]@{
        Label           = $Label
        StagedFiles     = $staged.Count
        LocalFiles      = $local.Count
        Identical       = $identical
        SameSize        = $sameSize
        DiffSize        = $diffSize
        DiffContent     = $diffContent
        Hashed          = $hashed
        OnlyStaged      = $onlyStaged
        OnlyLocal       = $onlyLocal
        OnlyStagedBytes = $onlyStagedBytes
        OnlyLocalBytes  = $onlyLocalBytes
        Rows            = $rows
    }
}

function Test-AccordanceRunning {
    $procs = @(Get-Process -Name 'Accordance' -ErrorAction SilentlyContinue)
    return ($procs.Count -gt 0)
}

function Invoke-RoboCopyFolder {
    # /E keeps empty folders so a workspace layout survives intact. /COPY:DAT and /DCOPY:DAT carry
    # timestamps without attempting ACLs, which would fail across two different user accounts.
    param(
        [string]$Source,
        [string]$Destination,
        [string[]]$ExtraArg = @()
    )
    $src = $Source.TrimEnd('\')
    $dst = $Destination.TrimEnd('\')
    $roboArgs = @($src, $dst, '/E', '/COPY:DAT', '/DCOPY:DAT', '/R:2', '/W:2', '/NP', '/NFL', '/NDL', '/NJH')
    if ($ExtraArg.Count -gt 0) { $roboArgs += $ExtraArg }
    $output = & robocopy.exe @roboArgs 2>&1
    $code = $LASTEXITCODE
    if ($script:LogFile) {
        try {
            Add-Content -LiteralPath $script:LogFile -Value ("--- robocopy " + $src + " -> " + $dst) -Encoding ASCII -WhatIf:$false -ErrorAction Stop
            Add-Content -LiteralPath $script:LogFile -Value ($output | Out-String) -Encoding ASCII -WhatIf:$false -ErrorAction Stop
        } catch { }
    }
    return $code
}

function Copy-AndVerify {
    # Robocopy exit codes below 8 are success. 8 and above mean at least one file failed.
    param(
        [string]$Label,
        [string]$Source,
        [string]$Destination
    )
    $before = Get-FolderStat -FolderPath $Source
    Write-Log ("  {0,-12} {1} files, {2}" -f $Label, $before.Files, (Format-Size $before.Bytes)) -Level Info
    Write-Log ("               from {0}" -f $Source) -Level Info
    Write-Log ("               to   {0}" -f $Destination) -Level Info

    if (-not $PSCmdlet.ShouldProcess($Destination, "Copy $Label from $Source")) { return $null }

    $code = Invoke-RoboCopyFolder -Source $Source -Destination $Destination
    if ($code -ge 8) {
        Add-Problem ("  {0,-12} robocopy reported failures, exit code {1} - see the log" -f $Label, $code)
    }

    $after = Get-FolderStat -FolderPath $Destination
    if ($after.Files -eq $before.Files -and $after.Bytes -eq $before.Bytes) {
        Write-Log ("  {0,-12} verified, {1} files, {2}" -f $Label, $after.Files, (Format-Size $after.Bytes)) -Level Good
    } else {
        Add-Problem ("  {0,-12} verify mismatch - source {1} files / {2} bytes, copy {3} files / {4} bytes" -f $Label, $before.Files, $before.Bytes, $after.Files, $after.Bytes)
    }
    return $before
}

# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------

if ($Mode -ne 'Inventory' -and -not $Path) {
    Write-Log "-Path is required for $Mode. Point it at the USB drive, external disk, or share both PCs can reach." -Level Bad
    exit 1
}

if (Test-AccordanceRunning) {
    Write-Log "Accordance is running. Close it and run this again - preferences and module indexes are written on exit." -Level Bad
    exit 1
}

$allItems = Get-AccordanceItem

# ---------------------------------------------------------------------------
# Inventory
# ---------------------------------------------------------------------------

if ($Mode -eq 'Inventory') {
    Write-Log "Accordance locations on $env:COMPUTERNAME (user $env:USERNAME):" -Level Head
    Write-Log ""
    $total = [int64]0
    foreach ($item in $allItems) {
        if ($item.Source) {
            $stat = Get-FolderStat -FolderPath $item.Source
            $total += $stat.Bytes
            Write-Log ("  {0,-12} FOUND  {1,8} files  {2,10}  {3}" -f $item.Name, $stat.Files, (Format-Size $stat.Bytes), $item.Source) -Level Good
        } else {
            $level = 'Warn'
            if ($item.Optional) { $level = 'Info' }
            Write-Log ("  {0,-12} absent                           looked for {1}" -f $item.Name, $item.Target) -Level $level
        }
        Write-Log ("               {0}" -f $item.Description) -Level Info
    }
    Write-Log ""
    Write-Log ("Total data found: {0}" -f (Format-Size $total)) -Level Head
    $appItem = $allItems | Where-Object { $_.Name -eq 'Application' }
    if ($appItem.Source) {
        $appStat = Get-FolderStat -FolderPath $appItem.Source
        Write-Log ("A default export excludes the Application folder and needs {0} of staging space." -f (Format-Size ($total - $appStat.Bytes))) -Level Head
    }
    Write-Log ""
    Write-Log "Next: .\$scriptName -Mode Export -Path <drive or share>" -Level Head
    exit 0
}

# ---------------------------------------------------------------------------
# Staging folder
# ---------------------------------------------------------------------------

$leaf = Split-Path -Path $Path.TrimEnd('\') -Leaf
if ($leaf -eq $stageFolderName) {
    $stageRoot = $Path.TrimEnd('\')
} else {
    $stageRoot = Join-Path $Path.TrimEnd('\') $stageFolderName
}
$manifestPath = Join-Path $stageRoot $manifestName

# ---------------------------------------------------------------------------
# Compare
# ---------------------------------------------------------------------------

if ($Mode -eq 'Compare') {
    if (-not (Test-Path -LiteralPath $manifestPath)) {
        Write-Log "No $manifestName found in $stageRoot. Point -Path at the folder the export was written to." -Level Bad
        exit 1
    }
    $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
    # ConvertFrom-Json hands back one object rather than enumerating, so this stays wrapped.
    $manifestItems = @($manifest.Items)

    Write-Log ("Export from {0} (user {1}), created {2}" -f $manifest.SourcePC, $manifest.SourceUser, $manifest.Created) -Level Head
    Write-Log ("Compared against {0} (user {1})" -f $env:COMPUTERNAME, $env:USERNAME) -Level Head
    Write-Log ""

    $allRows = @()
    foreach ($item in $allItems) {
        if ($SkipItem -and ($SkipItem -contains $item.Name)) { continue }
        $record = $manifestItems | Where-Object { $_.Name -eq $item.Name } | Select-Object -First 1
        if (-not $record) { continue }
        $stagePath = Join-Path $stageRoot $record.StageFolder
        if (-not (Test-Path -LiteralPath $stagePath)) {
            Add-Problem ("  {0,-12} manifest lists it but {1} is missing" -f $item.Name, $stagePath)
            continue
        }

        $result = Compare-Tree -Label $item.Name -StagedRoot $stagePath -LocalRoot $item.Target -UseHash:$Hash
        $allRows += $result.Rows

        Write-Log ("  {0}" -f $item.Name) -Level Head
        Write-Log ("    old PC {0} files, this PC {1} files, at {2}" -f $result.StagedFiles, $result.LocalFiles, $item.Target) -Level Info
        if ($Hash) {
            Write-Log ("    identical            {0}   (content verified by hash where sizes matched)" -f $result.Identical) -Level Info
            Write-Log ("    different content    {0}   <- same size, different bytes" -f $result.DiffContent) -Level Info
        } else {
            Write-Log ("    identical            {0}   (same size and write time)" -f $result.Identical) -Level Info
            Write-Log ("    same size, new time  {0}   <- almost certainly the same content, installed separately. Add -Hash to confirm" -f $result.SameSize) -Level Info
        }
        Write-Log ("    different size       {0}   <- genuinely different content" -f $result.DiffSize) -Level Info
        Write-Log ("    only on the old PC   {0}  ({1})  <- what an import would add" -f $result.OnlyStaged, (Format-Size $result.OnlyStagedBytes)) -Level Info
        Write-Log ("    only on this PC      {0}  ({1})  <- what an import would move into the .bak folder" -f $result.OnlyLocal, (Format-Size $result.OnlyLocalBytes)) -Level Info

        $sample = @($result.Rows | Where-Object { $_.Status -eq 'OnlyOnOldPC' } | Select-Object -First 8)
        if ($sample.Count -gt 0) {
            Write-Log "    sample of files only on the old PC:" -Level Info
            foreach ($row in $sample) { Write-Log ("      {0}" -f $row.RelativePath) -Level Info }
        }
        $sample = @($result.Rows | Where-Object { $_.Status -eq 'OnlyOnThisPC' } | Select-Object -First 8)
        if ($sample.Count -gt 0) {
            Write-Log "    sample of files only on this PC:" -Level Info
            foreach ($row in $sample) { Write-Log ("      {0}" -f $row.RelativePath) -Level Info }
        }
        Write-Log ""
    }

    if ($ReportPath) {
        if (-not (Test-Path -LiteralPath $ReportPath)) {
            New-Item -ItemType Directory -Path $ReportPath -Force | Out-Null
        }
        foreach ($group in ($allRows | Group-Object -Property Item)) {
            $csv = Join-Path $ReportPath ("compare-{0}-{1}.csv" -f $group.Name, $stamp)
            $group.Group | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding ASCII
            Write-Log ("Report: {0}  ({1} rows)" -f $csv, $group.Count) -Level Good
        }
        if ($allRows.Count -eq 0) { Write-Log "No differences to report." -Level Good }
    } else {
        Write-Log "Add -ReportPath <folder> to write the full per-file detail to CSV." -Level Warn
    }

    Write-Log ""
    Write-Log "Nothing was changed. Per item:" -Level Head
    Write-Log "  only on this PC is zero   ->  Import is safe, it takes nothing away." -Level Info
    Write-Log "  only on this PC is not    ->  Import would bury those files. Use Merge to add the" -Level Info
    Write-Log "                                old PC's missing files without touching anything else." -Level Info
    if ($script:Problems.Count -gt 0) { exit 1 }
    exit 0
}

# ---------------------------------------------------------------------------
# Merge
# ---------------------------------------------------------------------------

if ($Mode -eq 'Merge') {
    if (-not (Test-Path -LiteralPath $manifestPath)) {
        Write-Log "No $manifestName found in $stageRoot. Point -Path at the folder the export was written to." -Level Bad
        exit 1
    }
    $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
    # ConvertFrom-Json hands back one object rather than enumerating, so this stays wrapped.
    $manifestItems = @($manifest.Items)

    if (-not $WhatIfPreference) {
        $script:LogFile = Join-Path $stageRoot ("merge-{0}-{1}.log" -f $env:COMPUTERNAME, $stamp)
    }

    Write-Log ("Export from {0} (user {1}), created {2}" -f $manifest.SourcePC, $manifest.SourceUser, $manifest.Created) -Level Head
    Write-Log ("Merging into {0} (user {1})" -f $env:COMPUTERNAME, $env:USERNAME) -Level Head
    Write-Log "Only files absent from this PC are added. Nothing here is renamed, replaced, or deleted." -Level Head
    Write-Log ""

    Write-Log "Selecting items:" -Level Head
    $items = Select-MigrationItem -AllItems $allItems
    $elevated = Test-IsElevated

    $plan = @()
    foreach ($item in $items) {
        $record = $manifestItems | Where-Object { $_.Name -eq $item.Name } | Select-Object -First 1
        if (-not $record) {
            Write-Log ("  {0,-12} not in the manifest, nothing staged for it" -f $item.Name) -Level Warn
            continue
        }
        $stagePath = Join-Path $stageRoot $record.StageFolder
        if (-not (Test-Path -LiteralPath $stagePath)) {
            Add-Problem ("  {0,-12} manifest lists it but {1} is missing" -f $item.Name, $stagePath)
            continue
        }
        if ($item.NeedsAdmin -and -not $elevated) {
            Add-Problem ("  {0,-12} target {1} needs an elevated session - re-run PowerShell as administrator" -f $item.Name, $item.Target)
            continue
        }
        $plan += [pscustomobject]@{
            Name      = $item.Name
            StagePath = $stagePath
            Target    = $item.Target
        }
    }

    if ($script:Problems.Count -gt 0) {
        Write-Log ""
        Write-Log ("Nothing has been changed. {0} item(s) cannot be merged:" -f $script:Problems.Count) -Level Bad
        foreach ($problem in $script:Problems) { Write-Log $problem -Level Bad }
        exit 1
    }
    if ($plan.Count -eq 0) {
        Write-Log ""
        Write-Log "Nothing to merge." -Level Bad
        exit 1
    }

    Write-Log ""
    foreach ($step in $plan) {
        # The comparison runs first so the expected number of additions is known, which turns the
        # post-copy count into a real check rather than a restatement of what robocopy did.
        $diff = Compare-Tree -Label $step.Name -StagedRoot $step.StagePath -LocalRoot $step.Target
        $before = Get-FolderStat -FolderPath $step.Target

        Write-Log ("  {0,-12} {1} file(s) to add ({2}), {3} existing file(s) left untouched" -f $step.Name, $diff.OnlyStaged, (Format-Size $diff.OnlyStagedBytes), $before.Files) -Level Info
        Write-Log ("               into {0}" -f $step.Target) -Level Info
        if ($diff.OnlyStaged -eq 0) {
            Write-Log ("  {0,-12} nothing to add, skipped" -f $step.Name) -Level Good
            continue
        }
        foreach ($row in @($diff.Rows | Where-Object { $_.Status -eq 'OnlyOnOldPC' } | Select-Object -First 20)) {
            Write-Log ("                 + {0}" -f $row.RelativePath) -Level Info
        }
        if ($diff.OnlyStaged -gt 20) {
            Write-Log ("                 ... and {0} more" -f ($diff.OnlyStaged - 20)) -Level Info
        }

        if (-not $PSCmdlet.ShouldProcess($step.Target, ("Add {0} missing file(s) from {1}" -f $diff.OnlyStaged, $step.StagePath))) { continue }

        # /XC /XN /XO exclude changed, newer, and older files, which leaves only the files that do
        # not exist at the destination at all. Identical files are skipped by robocopy anyway.
        $code = Invoke-RoboCopyFolder -Source $step.StagePath -Destination $step.Target -ExtraArg @('/XC', '/XN', '/XO')
        if ($code -ge 8) {
            Add-Problem ("  {0,-12} robocopy reported failures, exit code {1} - see the log" -f $step.Name, $code)
        }

        $after = Get-FolderStat -FolderPath $step.Target
        $added = $after.Files - $before.Files
        if ($added -eq $diff.OnlyStaged) {
            Write-Log ("  {0,-12} added {1} file(s), now {2} files, {3}" -f $step.Name, $added, $after.Files, (Format-Size $after.Bytes)) -Level Good
        } else {
            Add-Problem ("  {0,-12} expected to add {1} file(s) but the count moved by {2}" -f $step.Name, $diff.OnlyStaged, $added)
        }
    }

    Write-Log ""
    if ($script:Problems.Count -gt 0) {
        Write-Log ("Merge finished with {0} problem(s):" -f $script:Problems.Count) -Level Bad
        foreach ($problem in $script:Problems) { Write-Log $problem -Level Bad }
        exit 1
    }
    Write-Log "Merge complete. Nothing was removed, so there is no backup folder to clean up." -Level Good
    Write-Log "Start Accordance and confirm the added modules appear in your library." -Level Good
    if ($script:LogFile) { Write-Log "Log: $script:LogFile" -Level Info }
    exit 0
}

# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------

if ($Mode -eq 'Export') {
    $parent = Split-Path -Path $stageRoot -Parent
    if (-not (Test-Path -LiteralPath $parent)) {
        Write-Log "Staging location not reachable: $parent" -Level Bad
        exit 1
    }

    Write-Log "Selecting items:" -Level Head
    $items = Select-MigrationItem -AllItems $allItems
    $present = @($items | Where-Object { $_.Source })
    $missing = @($items | Where-Object { -not $_.Source })

    foreach ($item in $missing) {
        $level = 'Warn'
        if ($item.Optional) { $level = 'Info' }
        Write-Log ("  {0,-12} absent on this PC, nothing to export - looked for {1}" -f $item.Name, $item.Target) -Level $level
    }

    if ($present.Count -eq 0) {
        Write-Log ""
        Write-Log "No Accordance data found on this PC. Run -Mode Inventory to see every path that was checked." -Level Bad
        exit 1
    }

    $needed = [int64]0
    foreach ($item in $present) { $needed += (Get-FolderStat -FolderPath $item.Source).Bytes }
    $driveRoot = [System.IO.Path]::GetPathRoot($stageRoot)
    try {
        $drive = Get-PSDrive -Name $driveRoot.Substring(0, 1) -ErrorAction Stop
        if ($null -ne $drive.Free -and $drive.Free -lt $needed) {
            Write-Log ("Not enough free space on {0}: need {1}, have {2}." -f $driveRoot, (Format-Size $needed), (Format-Size $drive.Free)) -Level Bad
            exit 1
        }
    } catch {
        Write-Log ("Could not read free space on {0}, continuing. Needed: {1}" -f $driveRoot, (Format-Size $needed)) -Level Warn
    }

    if (Test-Path -LiteralPath $manifestPath) {
        if (-not $Force) {
            Write-Log "$stageRoot already holds an export. Re-run with -Force to refresh it, or pick an empty folder." -Level Bad
            exit 1
        }
        Write-Log "$stageRoot already holds an export - robocopy will refresh it in place." -Level Warn
    }

    if ($PSCmdlet.ShouldProcess($stageRoot, "Create staging folder")) {
        if (-not (Test-Path -LiteralPath $stageRoot)) {
            New-Item -ItemType Directory -Path $stageRoot -Force | Out-Null
        }
        $script:LogFile = Join-Path $stageRoot ("export-{0}-{1}.log" -f $env:COMPUTERNAME, $stamp)
    }

    Write-Log ""
    Write-Log "Exporting to $stageRoot" -Level Head
    $records = @()
    foreach ($item in $present) {
        $dest = Join-Path $stageRoot $item.Name
        $stat = Copy-AndVerify -Label $item.Name -Source $item.Source -Destination $dest
        if ($null -eq $stat) { continue }
        $records += [pscustomobject]@{
            Name        = $item.Name
            Description = $item.Description
            SourcePath  = $item.Source
            StageFolder = $item.Name
            Files       = $stat.Files
            Bytes       = $stat.Bytes
            NeedsAdmin  = $item.NeedsAdmin
        }
    }

    if ($PSCmdlet.ShouldProcess($manifestPath, "Write manifest")) {
        $manifest = [pscustomobject]@{
            Tool         = $scriptName
            ToolVersion  = $scriptVer
            Created      = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            SourcePC     = $env:COMPUTERNAME
            SourceUser   = $env:USERNAME
            WindowsBuild = (Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue).Version
            Items        = $records
        }
        $manifest | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $manifestPath -Encoding ASCII
        Write-Log ""
        Write-Log "Manifest written: $manifestPath" -Level Good
    }

    Write-Log ""
    if ($script:Problems.Count -gt 0) {
        Write-Log ("Export finished with {0} problem(s) - review them before trusting this staging copy." -f $script:Problems.Count) -Level Bad
        exit 1
    }
    Write-Log "Export complete. On the new PC: install Accordance, launch it once, close it, then run" -Level Good
    Write-Log "  .\$scriptName -Mode Import -Path $Path -Force" -Level Good
    exit 0
}

# ---------------------------------------------------------------------------
# Import
# ---------------------------------------------------------------------------

if (-not (Test-Path -LiteralPath $manifestPath)) {
    Write-Log "No $manifestName found in $stageRoot. Point -Path at the folder the export was written to." -Level Bad
    exit 1
}

$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
# ConvertFrom-Json hands back one object rather than enumerating, so this stays wrapped.
$manifestItems = @($manifest.Items)

Write-Log ("Manifest from {0} (user {1}), created {2} by v{3}" -f $manifest.SourcePC, $manifest.SourceUser, $manifest.Created, $manifest.ToolVersion) -Level Head
Write-Log ("Importing onto {0} (user {1})" -f $env:COMPUTERNAME, $env:USERNAME) -Level Head
Write-Log ""

if (-not $WhatIfPreference) {
    $script:LogFile = Join-Path $stageRoot ("import-{0}-{1}.log" -f $env:COMPUTERNAME, $stamp)
}

Write-Log "Selecting items:" -Level Head
$items = Select-MigrationItem -AllItems $allItems
$elevated = Test-IsElevated

$plan = @()
foreach ($item in $items) {
    $record = $manifestItems | Where-Object { $_.Name -eq $item.Name } | Select-Object -First 1
    if (-not $record) {
        Write-Log ("  {0,-12} not in the manifest, nothing staged for it" -f $item.Name) -Level Warn
        continue
    }
    $stagePath = Join-Path $stageRoot $record.StageFolder
    if (-not (Test-Path -LiteralPath $stagePath)) {
        Add-Problem ("  {0,-12} manifest lists it but {1} is missing" -f $item.Name, $stagePath)
        continue
    }
    $staged = Get-FolderStat -FolderPath $stagePath
    if ($staged.Files -ne $record.Files -or $staged.Bytes -ne $record.Bytes) {
        Add-Problem ("  {0,-12} staged data does not match the manifest - expected {1} files / {2} bytes, found {3} files / {4} bytes" -f $item.Name, $record.Files, $record.Bytes, $staged.Files, $staged.Bytes)
        continue
    }
    if ($item.NeedsAdmin -and -not $elevated) {
        Add-Problem ("  {0,-12} target {1} needs an elevated session - re-run PowerShell as administrator" -f $item.Name, $item.Target)
        continue
    }
    $existing = Get-FolderStat -FolderPath $item.Target
    if ($existing.Files -gt 0 -and -not $Force) {
        Add-Problem ("  {0,-12} target already holds {1} files - re-run with -Force to rename it aside first: {2}" -f $item.Name, $existing.Files, $item.Target)
        continue
    }
    $plan += [pscustomobject]@{
        Name          = $item.Name
        StagePath     = $stagePath
        Target        = $item.Target
        ExistingFiles = $existing.Files
    }
}

# Every check above runs before anything is renamed or copied. A partial import is the one
# outcome worth avoiding here: new modules against old preferences is a state neither the old
# nor the new PC ever had, and it is not obvious from inside Accordance that it happened.
if ($script:Problems.Count -gt 0) {
    Write-Log ""
    Write-Log ("Nothing has been changed. {0} item(s) cannot be imported:" -f $script:Problems.Count) -Level Bad
    foreach ($problem in $script:Problems) { Write-Log $problem -Level Bad }
    Write-Log ""
    Write-Log "Fix the above and run again, or name the blocked items in -SkipItem to migrate the rest deliberately." -Level Warn
    exit 1
}

if ($plan.Count -eq 0) {
    Write-Log ""
    Write-Log "Nothing to import." -Level Bad
    exit 1
}

Write-Log ""
$renamed = 0
foreach ($step in $plan) {
    if ($step.ExistingFiles -gt 0) {
        $backup = "{0}.bak-{1}" -f $step.Target.TrimEnd('\'), $stamp
        if ($PSCmdlet.ShouldProcess($step.Target, "Rename to $backup")) {
            try {
                Move-Item -LiteralPath $step.Target -Destination $backup -ErrorAction Stop
                $renamed++
                Write-Log ("  {0,-12} existing folder renamed to {1}" -f $step.Name, $backup) -Level Warn
            } catch {
                Add-Problem ("  {0,-12} could not rename {1} - {2}" -f $step.Name, $step.Target, $_.Exception.Message)
                continue
            }
        }
    }
    Copy-AndVerify -Label $step.Name -Source $step.StagePath -Destination $step.Target | Out-Null
}

Write-Log ""
if ($script:Problems.Count -gt 0) {
    Write-Log ("Import finished with {0} problem(s):" -f $script:Problems.Count) -Level Bad
    foreach ($problem in $script:Problems) { Write-Log $problem -Level Bad }
    if ($renamed -gt 0) {
        Write-Log ""
        Write-Log "Folders renamed during this run are still on disk as <name>.bak-$stamp if you need to go back." -Level Warn
    }
    exit 1
}

Write-Log "Import complete. Start Accordance and check your workspaces, user notes, and highlights." -Level Good
if ($renamed -gt 0) {
    Write-Log "Once it all looks right, delete the <name>.bak-$stamp folders to reclaim the space." -Level Good
}
if ($script:LogFile) { Write-Log "Log: $script:LogFile" -Level Info }
exit 0
