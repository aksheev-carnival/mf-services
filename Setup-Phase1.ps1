<#
.SYNOPSIS
  Phase 1 - SharePoint/OneDrive information architecture for the Mainframe Investigation Assistant.

.DESCRIPTION
  DRY RUN BY DEFAULT. Nothing is created, moved or changed unless you pass -Apply.

  What it does:
    1. Takes a before-snapshot of every file under MF Codes (CSV).
    2. Creates the target folder structure (idempotent - existing folders are kept).
    3. Moves only files that match explicit, conservative patterns:
         *.bak / *.old                 -> MF Codes\Archive\Backups
         PST*/PRD* schema + RDMS files -> MF Codes\Schemas\DMS|RDMS\PST|PRD
         ESDR-*.txt, *-QLP.txt, YYYAKSSQL*.txt -> MF Codes\SQL QLP results
         Sort-MFCodes.ps1              -> Work Tracker\Agent Configuration\Scripts
    4. NEVER overwrites. On a name clash it appends a timestamp to the moved file.
    5. Reports (read-only, never moves): duplicate element names, files in the wrong
       folder per your naming rules, and unclassified files.
    6. Creates the Investigations\Txx folder tree for the open tracker rows.
    7. Writes a README with the naming/versioning rules.

  Existing folders (BSRC, RSRC, PROCS, ECL, COPYLIB, Schemas, prod, dev, SQL QLP results)
  are used as-is. prod and dev are never touched.

.EXAMPLE
  .\Setup-Phase1.ps1                       # dry run, prints what it WOULD do
  .\Setup-Phase1.ps1 -Apply                # does it

.NOTES
  Not tested in the authoring environment (no PowerShell available there).
  Run the dry run first and read the report before using -Apply.
  If Windows blocks the script:  Unblock-File .\Setup-Phase1.ps1
#>
[CmdletBinding()]
param(
    [string]$WorkTrackerRoot,
    [string]$MfCodesRoot,
    [switch]$Apply
)

$ErrorActionPreference = 'Stop'
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'

# ---------------------------------------------------------------- locate roots
function Find-First([string[]]$candidates) {
    foreach ($c in $candidates) { if ($c -and (Test-Path -LiteralPath $c)) { return (Resolve-Path -LiteralPath $c).Path } }
    return $null
}
$od = @($env:OneDriveCommercial, $env:OneDrive) | Where-Object { $_ }

if (-not $WorkTrackerRoot) {
    $WorkTrackerRoot = Find-First ($od | ForEach-Object { Join-Path $_ 'Work Tracker' })
}
if (-not $MfCodesRoot) {
    $cands = @()
    if ($WorkTrackerRoot) { $cands += (Join-Path $WorkTrackerRoot 'MF Codes') }
    $cands += ($od | ForEach-Object { Join-Path $_ 'MF Codes' })
    $MfCodesRoot = Find-First $cands
}
if (-not $WorkTrackerRoot) { throw "Could not find the 'Work Tracker' folder. Re-run with -WorkTrackerRoot '<full path>'." }
if (-not $MfCodesRoot)     { throw "Could not find the 'MF Codes' folder. Re-run with -MfCodesRoot '<full path>'." }

$mode = if ($Apply) { 'APPLY' } else { 'DRY RUN' }
Write-Host "Mode            : $mode" -ForegroundColor Yellow
Write-Host "Work Tracker    : $WorkTrackerRoot"
Write-Host "MF Codes        : $MfCodesRoot"
Write-Host ""

$cfgRoot  = Join-Path $WorkTrackerRoot 'Agent Configuration'
$logDir   = Join-Path $cfgRoot 'Logs'
$scrDir   = Join-Path $cfgRoot 'Scripts'
$invRoot  = Join-Path $WorkTrackerRoot 'Investigations'
$report   = New-Object System.Collections.Generic.List[object]

function Add-Report($action, $path, $target, $note) {
    $report.Add([pscustomobject]@{ Action = $action; Path = $path; Target = $target; Note = $note })
}

# ---------------------------------------------------------------- 1. snapshot
$allFiles = Get-ChildItem -LiteralPath $MfCodesRoot -Recurse -File -ErrorAction SilentlyContinue
Write-Host ("Files currently under MF Codes: {0}" -f $allFiles.Count)

# ---------------------------------------------------------------- 2. folders
$mfFolders = @(
    'BSRC','RSRC','PROCS','ECL','COPYLIB',
    'Schemas\DMS\PST','Schemas\DMS\PRD','Schemas\RDMS\PST','Schemas\RDMS\PRD',
    'SQL QLP results','Archive\Backups'
)
$wtFolders = @('Agent Configuration','Agent Configuration\Scripts','Agent Configuration\Logs','Investigations','Investigations\_Archive')

$invFolders = @(
    'T01 - Excursion price set to zero (SXR75)',
    'T02 - CABIN_NET vs YMBCCD count gap',
    'T03 - GoCCL post PPO change',
    'T04 - MCA connecting cabins fields',
    'T05 - Booking claimant validation',
    'T06 - DPBP0U file overflow',
    'T07 - AGENT_DIM vs AGENT_EMAIL_ADDR_DIM',
    'T08 - PAX phone seq num to DWH',
    'T09 - YMBIBR error file email card',
    'T10 - Empty ShoreX files Azure',
    'T11 - GIBR supercombo bug 2404922',
    'T12 - SXBFTS abort suppression'
)
$invSub = @('Evidence','Query Results','Email Exports','Code Snapshots')

function Ensure-Dir($path) {
    if (Test-Path -LiteralPath $path) { return }
    Add-Report 'CREATE_FOLDER' $path '' ''
    if ($Apply) { New-Item -ItemType Directory -Path $path -Force | Out-Null }
}
foreach ($f in $mfFolders) { Ensure-Dir (Join-Path $MfCodesRoot $f) }
foreach ($f in $wtFolders) { Ensure-Dir (Join-Path $WorkTrackerRoot $f) }
foreach ($t in $invFolders) {
    Ensure-Dir (Join-Path $invRoot $t)
    foreach ($s in $invSub) { Ensure-Dir (Join-Path (Join-Path $invRoot $t) $s) }
}

# ---------------------------------------------------------------- 3. moves
function Safe-Move($file, $destDir, $why) {
    if ($file.DirectoryName -eq $destDir) { return }   # already in place
    $dest = Join-Path $destDir $file.Name
    if (Test-Path -LiteralPath $dest) {
        $dest = Join-Path $destDir ("{0}__{1}{2}" -f $file.BaseName, $stamp, $file.Extension)
        $why  = "$why (name clash - timestamp added)"
    }
    Add-Report 'MOVE' $file.FullName $dest $why
    if ($Apply) {
        if (-not (Test-Path -LiteralPath $destDir)) { New-Item -ItemType Directory -Path $destDir -Force | Out-Null }
        Move-Item -LiteralPath $file.FullName -Destination $dest
    }
}

$protected = @('\prod\','\dev\')   # never touch these
function Is-Protected($p) { foreach ($x in $protected) { if ($p -like "*$x*") { return $true } }; return $false }

$archiveBak = Join-Path $MfCodesRoot 'Archive\Backups'
$qlpDir     = Join-Path $MfCodesRoot 'SQL QLP results'

foreach ($f in $allFiles) {
    if (Is-Protected $f.FullName) { continue }
    $name = $f.Name

    if ($f.Extension -in '.bak','.old') { Safe-Move $f $archiveBak 'backup file kept out of active search folders'; continue }

    # schemas and RDMS catalogs
    if ($name -match '^(PST|PRD)(100|200|300)SCHEMA\.txt$') {
        Safe-Move $f (Join-Path $MfCodesRoot ("Schemas\DMS\{0}" -f $Matches[1])) 'DMS schema'; continue
    }
    if ($name -match '^(PST|PRD)RDMS\.txt$') {
        Safe-Move $f (Join-Path $MfCodesRoot ("Schemas\RDMS\{0}" -f $Matches[1])) 'RDMS catalog'; continue
    }

    # evidence files (not source)
    if ($name -match '^ESDR-.*\.txt$' -or $name -match '.*-QLP\.txt$' -or $name -match '^YYYAKSSQL.*\.txt$') {
        Safe-Move $f $qlpDir 'query/history evidence, not source'; continue
    }

    if ($name -ieq 'Sort-MFCodes.ps1') { Safe-Move $f $scrDir 'helper script'; continue }
}

# ---------------------------------------------------------------- 5. read-only checks
function Expected-Folder([string]$el) {
    if ($el.Length -eq 12 -and $el[2] -eq 'S') { return 'BSRC' }
    if ($el.Length -eq 6)  { return 'ECL' }
    if ($el.Length -eq 8) {
        switch ($el[2]) {
            'B' { return 'BSRC' } 'O' { return 'BSRC' } 'C' { return 'BSRC' }
            'R' { return 'RSRC' } 'U' { return 'RSRC' }
            'P' { return 'PROCS' }
        }
    }
    return $null
}

$active = Get-ChildItem -LiteralPath $MfCodesRoot -Recurse -File -ErrorAction SilentlyContinue |
          Where-Object { -not (Is-Protected $_.FullName) -and $_.Extension -notin '.bak','.old' }

# duplicates: same element name in more than one active location
$byEl = $active | Where-Object { $_.Extension -eq '.txt' } |
        Group-Object { ($_.Name -split '\.')[0].ToUpper() }
foreach ($g in $byEl) {
    if ($g.Count -gt 1) {
        Add-Report 'CHECK_DUPLICATE' ($g.Group.FullName -join ' | ') '' ("Element {0} exists {1} times - confirm which is current" -f $g.Name, $g.Count)
    }
}

# wrong folder per naming rules (only for the five source folders)
$srcFolders = 'BSRC','RSRC','PROCS','ECL'
foreach ($f in $active) {
    $rel = $f.FullName.Substring($MfCodesRoot.Length).TrimStart('\')
    $top = ($rel -split '\\')[0]
    if ($top -notin $srcFolders) { continue }
    $el  = ($f.Name -split '\.')[0].ToUpper()
    $exp = Expected-Folder $el
    if (-not $exp)          { Add-Report 'CHECK_UNCLASSIFIED' $f.FullName '' "Name '$el' does not match any naming rule" }
    elseif ($exp -ne $top)  { Add-Report 'CHECK_WRONG_FOLDER'  $f.FullName (Join-Path $MfCodesRoot $exp) "Rules say $exp, file is in $top" }
}

# loose files at the MF Codes root
foreach ($f in (Get-ChildItem -LiteralPath $MfCodesRoot -File -ErrorAction SilentlyContinue)) {
    Add-Report 'CHECK_LOOSE_FILE' $f.FullName '' 'File sits at the MF Codes root, not in a category folder'
}

# ---------------------------------------------------------------- 7. README
$readme = @"
MF CODES - NAMING AND VERSIONING RULES  (created by Setup-Phase1.ps1, $stamp)

FOLDERS
  BSRC     batch jobs (3rd char B), subroutines (O), subsystems (C), screens (12 chars, 3rd char S)
  RSRC     real-time transactions (3rd char R or U)
  PROCS    procedures (3rd char P)
  ECL      runstreams / ECL (6 characters)
  COPYLIB  copy members and copybooks
  Schemas  DMS and RDMS schema files, split by PST (test) and PRD (production)
  SQL QLP results   saved QLP / SQL / history output used as evidence
  Archive\Backups   .bak and superseded files - NOT searched for current logic
  prod, dev         existing folders, left untouched

RULES
  1. One active file per element, named exactly after the element, extension .txt.
  2. A file in prod/dev is treated as that environment's version. Never mix them in the active folders.
  3. Never overwrite evidence. Add a date to the name instead (Txx_YYYYMMDD_description.txt).
  4. A line with * in column 7 is a comment, not active logic.
  5. If the same element appears twice, stop and resolve which is current before analysing.
  6. Query results saved as: Txx_YYYYMMDD_ENV_object_identifier.txt
     Example: T02_20261007_PST_SAIL_CATEGORY_SAIL22844.txt
"@
$readmePath = Join-Path $cfgRoot 'README - MF Codes rules.txt'
Add-Report 'WRITE_FILE' $readmePath '' 'naming and versioning rules'
if ($Apply) {
    if (-not (Test-Path -LiteralPath $cfgRoot)) { New-Item -ItemType Directory -Path $cfgRoot -Force | Out-Null }
    Set-Content -LiteralPath $readmePath -Value $readme -Encoding UTF8
}

# ---------------------------------------------------------------- outputs
$summary = $report | Group-Object Action | Sort-Object Name | ForEach-Object { "{0,-20}{1}" -f $_.Name, $_.Count }
Write-Host "---- SUMMARY ($mode) ----" -ForegroundColor Cyan
$summary | ForEach-Object { Write-Host $_ }

$reportCsv = Join-Path $env:TEMP ("Phase1_Report_{0}.csv" -f $stamp)
$snapCsv   = Join-Path $env:TEMP ("Phase1_Snapshot_before_{0}.csv" -f $stamp)
$report   | Export-Csv -LiteralPath $reportCsv -NoTypeInformation -Encoding UTF8
$allFiles | Select-Object FullName, Length, LastWriteTime | Export-Csv -LiteralPath $snapCsv -NoTypeInformation -Encoding UTF8

if ($Apply -and (Test-Path -LiteralPath $logDir)) {
    Copy-Item -LiteralPath $reportCsv -Destination $logDir -Force
    Copy-Item -LiteralPath $snapCsv   -Destination $logDir -Force
    Write-Host "Report and before-snapshot copied to: $logDir"
} else {
    Write-Host "Report   : $reportCsv"
    Write-Host "Snapshot : $snapCsv"
}
if (-not $Apply) { Write-Host "`nDRY RUN ONLY - nothing was changed. Re-run with -Apply after reviewing the report." -ForegroundColor Yellow }
