<#
.SYNOPSIS
    Scans a folder for split archive parts and combines or extracts them.

.DESCRIPTION
    Combine-Archives.ps1 automatically detects split archive sets in a folder
    and combines or extracts them. Run without parameters for interactive mode.
    It supports the following formats:

      - Generic binary splits  : file.zip.001, file.zip.002 ...
      - HJSplit / 7-Zip splits : file.001, file.002 ...
      - Multi-part RAR         : file.part1.rar, file.part01.rar ...
      - Old-style RAR          : file.rar + file.r00, file.r01 ...

    ZIP/binary splits are joined using pure PowerShell (no tools required).
    RAR files require 7-Zip to be installed (auto-detected in default locations).
    Uses chunked streaming so files larger than 2GB are handled correctly.
    Displays progress bars in console and a text fallback in PowerShell ISE.

.PARAMETER FolderPath
    The folder to scan for split archive parts.
    Defaults to the current directory if not specified.

.PARAMETER OutputFolder
    The folder where combined/extracted files will be saved.
    Defaults to the same folder as FolderPath if not specified.

.PARAMETER SevenZipPath
    Full path to 7z.exe if it is not installed in the default location
    (C:\Program Files\7-Zip\7z.exe).

.PARAMETER DryRun
    Preview mode. Shows what the script would do without actually
    combining or extracting anything. Useful for verifying detection.

.PARAMETER DeletePartsAfter
    Automatically deletes the individual part files after they have been
    successfully combined or extracted.

.PARAMETER BufferSizeMB
    Size of the read/write buffer in megabytes. Default is 64MB.
    Increase for faster transfers on systems with plenty of RAM.

.PARAMETER NonInteractive
    Force non-interactive mode even when no parameters are provided.
    In non-interactive mode all detected archives are processed automatically.

.EXAMPLE
    .\Combine-Archives.ps1

    Launches the interactive TUI menu.

.EXAMPLE
    .\Combine-Archives.ps1 -FolderPath "D:\Downloads"

    Scans D:\Downloads and combines any split archives found there.

.EXAMPLE
    .\Combine-Archives.ps1 -FolderPath "D:\Downloads" -DryRun

    Previews what would be combined in D:\Downloads without doing anything.

.EXAMPLE
    .\Combine-Archives.ps1 -FolderPath "D:\Downloads" -OutputFolder "D:\Combined"

    Scans D:\Downloads and saves all combined/extracted files to D:\Combined.

.EXAMPLE
    .\Combine-Archives.ps1 -FolderPath "D:\Downloads" -DeletePartsAfter

    Combines archives in D:\Downloads and deletes the parts afterwards.

.EXAMPLE
    .\Combine-Archives.ps1 -FolderPath "D:\Downloads" -SevenZipPath "C:\Tools\7z.exe"

    Uses a custom 7-Zip location for extracting RAR files.

.EXAMPLE
    .\Combine-Archives.ps1 -FolderPath "D:\Downloads" -OutputFolder "D:\Out" -DeletePartsAfter

    Combines archives, saves results to D:\Out, and cleans up part files.

.NOTES
    - RAR extraction requires 7-Zip (https://www.7-zip.org)
    - ZIP and generic binary splits need no additional tools
    - Files larger than 2GB are supported via chunked streaming
    - Progress bars shown in console; text fallback used in PowerShell ISE
    - If script is blocked, run:
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
    - Use -DryRun first to verify the script detects your files correctly

.LINK
    https://www.7-zip.org
#>

param (
    [string] $FolderPath   = "",
    [string] $OutputFolder = "",
    [string] $SevenZipPath = "",
    [int]    $BufferSizeMB = 64,
    [switch] $DryRun,
    [switch] $DeletePartsAfter,
    [switch] $NonInteractive
)

# =============================================================================
#  ENVIRONMENT DETECTION
# =============================================================================

$script:IsISE = ($null -ne $psISE)

# =============================================================================
#  UI HELPERS
# =============================================================================

$script:width = 64

function Draw-Line  { Write-Host ("+$("-" * ($script:width - 2))+") -ForegroundColor DarkCyan }
function Draw-Blank { Write-Host ("|$(" " * ($script:width - 2))|") -ForegroundColor DarkCyan }

function Draw-Title {
    param([string]$Text)
    $pad   = $script:width - 2 - $Text.Length
    $left  = [math]::Floor($pad / 2)
    $right = $pad - $left
    Write-Host ("|$(" " * $left)") -ForegroundColor DarkCyan -NoNewline
    Write-Host $Text -ForegroundColor White -NoNewline
    Write-Host ("$(" " * $right)|") -ForegroundColor DarkCyan
}

function Draw-Header {
    Clear-Host
    Draw-Line
    Draw-Blank
    Draw-Title "Combine-Archives"
    Draw-Title "Split Archive Combiner v1.2"
    Draw-Blank
    Draw-Line
}

function Draw-Section {
    param([string]$Title)
    Write-Host ""
    Write-Host "  $Title" -ForegroundColor Cyan
    Write-Host "  $("-" * ($Title.Length))" -ForegroundColor DarkCyan
}

function Write-Status {
    param([string]$Label, [string]$Value, [ConsoleColor]$ValueColor = "White")
    Write-Host "  $($Label.PadRight(20)): " -ForegroundColor Gray -NoNewline
    Write-Host $Value -ForegroundColor $ValueColor
}

function Write-OK   { param($m) Write-Host "  [OK]  $m" -ForegroundColor Green }
function Write-Info { param($m) Write-Host "  [..]  $m" -ForegroundColor Yellow }
function Write-Fail { param($m) Write-Host "  [!!]  $m" -ForegroundColor Red }
function Write-Dry  { param($m) Write-Host "  [DRY] $m" -ForegroundColor Magenta }

function Prompt-Input {
    param([string]$Prompt, [string]$Default = "")
    $hint = if ($Default) { " (default: $Default)" } else { "" }
    Write-Host "  $Prompt$hint" -ForegroundColor Cyan -NoNewline
    Write-Host " > " -ForegroundColor DarkCyan -NoNewline
    $val = Read-Host
    if ([string]::IsNullOrWhiteSpace($val) -and $Default) { return $Default }
    return $val
}

function Prompt-YesNo {
    param([string]$Prompt, [bool]$Default = $false)
    $hint = if ($Default) { "[Y/n]" } else { "[y/N]" }
    Write-Host "  $Prompt $hint" -ForegroundColor Cyan -NoNewline
    Write-Host " > " -ForegroundColor DarkCyan -NoNewline
    $val = Read-Host
    if ([string]::IsNullOrWhiteSpace($val)) { return $Default }
    return $val -match '^[Yy]'
}

function Format-Bytes {
    param([long]$Bytes)
    if     ($Bytes -ge 1GB) { "{0:N2} GB" -f ($Bytes / 1GB) }
    elseif ($Bytes -ge 1MB) { "{0:N1} MB" -f ($Bytes / 1MB) }
    elseif ($Bytes -ge 1KB) { "{0:N1} KB" -f ($Bytes / 1KB) }
    else                    { "$Bytes B" }
}

function Pause-Screen {
    Write-Host ""
    Write-Host "  Press Enter to continue..." -ForegroundColor DarkGray -NoNewline
    Read-Host | Out-Null
}

# =============================================================================
#  ARCHIVE SELECTION MENU
#
#  Shows a numbered list of detected archive sets.
#  User can type:
#    - Individual numbers    : 1,3,5
#    - Ranges                : 2-4
#    - Mix of both           : 1,3-5,7
#    - "a" or blank Enter    : select all
#    - "n"                   : select none / cancel
# =============================================================================

function Select-Archives {
    param([object[]]$Groups)

    $typeLabel = @{
        BinaryParts  = "ZIP/Binary"
        MultiPartRar = "Multi-RAR "
        OldStyleRar  = "Old RAR   "
        SplitZip     = "Split ZIP "
    }

    Draw-Section "Select Archives to Process"
    Write-Host "  Enter numbers to select (e.g. 1,3 or 2-4 or 1,3-5)." -ForegroundColor Gray
    Write-Host "  Press Enter or type 'a' to select all. Type 'n' to cancel." -ForegroundColor DarkGray
    Write-Host ""

    # Pre-calculate sizes for all groups so we can align columns
    $groupMeta = @()
    foreach ($g in $Groups) {
        $bytes = [long]($g.Parts | ForEach-Object { (Get-Item -LiteralPath $_).Length } | Measure-Object -Sum).Sum
        $groupMeta += [PSCustomObject]@{ Group = $g; Bytes = $bytes; SizeStr = Format-Bytes $bytes }
    }

    # Column widths driven by longest values
    $maxNameLen = ($groupMeta | ForEach-Object { $_.Group.Name.Length } | Measure-Object -Maximum).Maximum
    $maxSizeLen = ($groupMeta | ForEach-Object { $_.SizeStr.Length }    | Measure-Object -Maximum).Maximum
    $maxNameLen = [math]::Max($maxNameLen, 12)
    $maxSizeLen = [math]::Max($maxSizeLen, 7)

    # Header row
    Write-Host ("  [ ##]  {0,-12}  {1}  {2}  {3}  {4}" -f "Type", "Name".PadRight($maxNameLen), "Pts", "Size".PadRight($maxSizeLen), "Cumulative") -ForegroundColor DarkGray
    Write-Host "  $("-" * 66)" -ForegroundColor DarkCyan

    # Rows with running cumulative total
    $cumulative = [long]0
    for ($i = 0; $i -lt $groupMeta.Count; $i++) {
        $m         = $groupMeta[$i]
        $g         = $m.Group
        $num       = ($i + 1).ToString().PadLeft(3)
        $label     = $typeLabel[$g.Type]
        $partCount = $g.Parts.Count.ToString().PadLeft(3)
        $sizeStr   = $m.SizeStr.PadLeft($maxSizeLen)
        $cumulative += $m.Bytes
        $cumStr    = Format-Bytes $cumulative
        $nameStr   = $g.Name.PadRight($maxNameLen)
        $cumColor  = "DarkGray"
        if ($i -eq $groupMeta.Count - 1) { $cumColor = "Yellow" }

        Write-Host "  [" -ForegroundColor DarkCyan       -NoNewline
        Write-Host $num  -ForegroundColor Yellow          -NoNewline
        Write-Host "]  " -ForegroundColor DarkCyan       -NoNewline
        Write-Host $label.PadRight(12) -ForegroundColor Cyan     -NoNewline
        Write-Host "$nameStr " -ForegroundColor White    -NoNewline
        Write-Host $partCount  -ForegroundColor DarkGray -NoNewline
        Write-Host "  "        -NoNewline
        Write-Host $sizeStr    -ForegroundColor DarkGray -NoNewline
        Write-Host "  "        -NoNewline
        Write-Host $cumStr     -ForegroundColor $cumColor
    }

    # Footer grand total
    Write-Host "  $("-" * 66)" -ForegroundColor DarkCyan
    $grandBytes = [long]($groupMeta | ForEach-Object { $_.Bytes } | Measure-Object -Sum).Sum
    Write-Host "  Total if all selected: " -ForegroundColor Gray -NoNewline
    Write-Host (Format-Bytes $grandBytes) -ForegroundColor Yellow

    Write-Host ""
    Write-Host "  Selection" -ForegroundColor Cyan -NoNewline
    Write-Host " > " -ForegroundColor DarkCyan -NoNewline
    $raw = Read-Host

    # Parse selection
    $raw = $raw.Trim()

    # All / blank
    if ([string]::IsNullOrWhiteSpace($raw) -or $raw -eq 'a') {
        Write-Host ""
        Write-OK "All $($Groups.Count) archive(s) selected."
        return $Groups
    }

    # None / cancel
    if ($raw -eq 'n') {
        Write-Host ""
        Write-Info "No archives selected. Returning to menu."
        return @()
    }

    # Parse numbers and ranges
    $selectedIndices = [System.Collections.Generic.HashSet[int]]::new()

    foreach ($token in ($raw -split ',')) {
        $token = $token.Trim()
        if ($token -match '^(\d+)-(\d+)$') {
            $from = [int]$Matches[1]
            $to   = [int]$Matches[2]
            if ($from -gt $to) { $from, $to = $to, $from }   # allow reverse ranges
            $from..$to | ForEach-Object { [void]$selectedIndices.Add($_) }
        } elseif ($token -match '^\d+$') {
            [void]$selectedIndices.Add([int]$token)
        } else {
            Write-Fail "Unrecognised token '$token' — ignored."
        }
    }

    # Map 1-based numbers to objects, filtering out-of-range entries
    $selected = @()
    foreach ($n in ($selectedIndices | Sort-Object)) {
        if ($n -lt 1 -or $n -gt $Groups.Count) {
            Write-Fail "Number $n is out of range (1-$($Groups.Count)) — skipped."
            continue
        }
        $selected += $Groups[$n - 1]
    }

    if ($selected.Count -eq 0) {
        Write-Host ""
        Write-Info "No valid archives selected."
        return @()
    }

    Write-Host ""
    Write-OK "$($selected.Count) archive(s) selected:"
    foreach ($g in $selected) {
        Write-Host "    - $($g.Name)" -ForegroundColor Gray
    }

    return $selected
}

# =============================================================================
#  PROGRESS — Console bars + ISE text fallback
# =============================================================================

$script:LastProgressLen = 0

function Show-Progress {
    param(
        [int]    $Id,
        [int]    $ParentId    = -1,
        [string] $Activity,
        [string] $Status      = "",
        [int]    $Percent     = 0,
        [switch] $Completed
    )

    if (-not $script:IsISE) {
        $params = @{
            Id              = $Id
            Activity        = $Activity
            Status          = $Status
            PercentComplete = $Percent
        }
        if ($ParentId -ge 0) { $params['ParentId'] = $ParentId }
        if ($Completed)      { $params['Completed'] = $true; $params.Remove('PercentComplete') | Out-Null }
        Write-Progress @params
    } else {
        if ($Completed) {
            Write-Host ("`r$(" " * $script:LastProgressLen)`r") -NoNewline
            $script:LastProgressLen = 0
            return
        }
        if ($Id -ne 1) { return }

        $barWidth = 30
        $filled   = [math]::Floor($barWidth * $Percent / 100)
        $empty    = $barWidth - $filled
        $bar      = "[" + ("#" * $filled) + ("-" * $empty) + "]"
        $line     = "  $bar $Percent%  $Status"

        try {
            $maxW = [Console]::WindowWidth - 2
            if ($line.Length -gt $maxW) { $line = $line.Substring(0, $maxW - 3) + "..." }
        } catch { }

        $padded = $line.PadRight($script:LastProgressLen)
        $script:LastProgressLen = $line.Length
        Write-Host "`r$padded" -NoNewline -ForegroundColor DarkCyan
    }
}

function Clear-AllProgress {
    if (-not $script:IsISE) {
        Write-Progress -Id 1 -Activity "Done" -Completed
        Write-Progress -Id 0 -Activity "Done" -Completed
    } else {
        Write-Host ("`r$(" " * $script:LastProgressLen)`r") -NoNewline
        $script:LastProgressLen = 0
    }
}

# =============================================================================
#  CORE FUNCTIONS
# =============================================================================

function Find-7Zip {
    if ($SevenZipPath -and (Test-Path $SevenZipPath)) { return $SevenZipPath }
    $candidates = @(
        "C:\Program Files\7-Zip\7z.exe",
        "C:\Program Files (x86)\7-Zip\7z.exe",
        (Get-Command "7z" -ErrorAction SilentlyContinue)?.Source
    )
    foreach ($c in $candidates) { if ($c -and (Test-Path $c)) { return $c } }
    return $null
}

function Join-BinaryParts {
    param(
        [string[]] $Parts,
        [string]   $OutputFile,
        [string]   $ArchiveName
    )

    $bufferSize   = $BufferSizeMB * 1024 * 1024
    $buffer       = New-Object byte[] $bufferSize
    $totalBytes   = [long]($Parts | ForEach-Object { (Get-Item -LiteralPath $_).Length } | Measure-Object -Sum).Sum
    $writtenTotal = [long]0
    $partIndex    = 0
    $partCount    = $Parts.Count
    $stopwatch    = [System.Diagnostics.Stopwatch]::StartNew()

    if ($script:IsISE) {
        Write-Host "  Combining: $ArchiveName  ($partCount parts, $(Format-Bytes $totalBytes) total)" -ForegroundColor Cyan
    }

    $outStream = [System.IO.File]::Open(
        $OutputFile,
        [System.IO.FileMode]::Create,
        [System.IO.FileAccess]::Write,
        [System.IO.FileShare]::None
    )

    try {
        foreach ($part in $Parts) {
            $partIndex++
            $partName    = Split-Path $part -Leaf
            $partSize    = [long](Get-Item -LiteralPath $part).Length
            $writtenPart = [long]0

            $overallPct = [math]::Min([int](($writtenTotal / [math]::Max([long]$totalBytes, [long]1)) * 100), 100)
            Show-Progress -Id 0 -Activity "Combining: $ArchiveName" `
                -Status  "Part $partIndex of $partCount  |  $(Format-Bytes $writtenTotal) of $(Format-Bytes $totalBytes)" `
                -Percent $overallPct

            $inStream = [System.IO.File]::Open(
                $part,
                [System.IO.FileMode]::Open,
                [System.IO.FileAccess]::Read,
                [System.IO.FileShare]::Read
            )

            try {
                while (($read = $inStream.Read($buffer, 0, $buffer.Length)) -gt 0) {
                    $outStream.Write($buffer, 0, $read)
                    $writtenPart  += $read
                    $writtenTotal += $read

                    $partPct = [math]::Min([int](($writtenPart / [math]::Max([long]$partSize, [long]1)) * 100), 100)
                    $elapsed = $stopwatch.Elapsed.TotalSeconds

                    if ($elapsed -gt 0) {
                        $speedMBs  = [math]::Round(($writtenTotal / 1MB) / $elapsed, 1)
                        $remaining = if ($speedMBs -gt 0) {
                            [math]::Round((($totalBytes - $writtenTotal) / 1MB) / $speedMBs)
                        } else { 0 }
                        $statusMsg = "$partName  |  $(Format-Bytes $writtenPart) / $(Format-Bytes $partSize)  |  $speedMBs MB/s  |  ~${remaining}s left"
                    } else {
                        $statusMsg = "$partName  |  $(Format-Bytes $writtenPart) / $(Format-Bytes $partSize)"
                    }

                    Show-Progress -Id 1 -ParentId 0 `
                        -Activity "Writing part $partIndex of $partCount" `
                        -Status   $statusMsg `
                        -Percent  $partPct
                }
            } finally {
                $inStream.Close()
            }

            if ($script:IsISE) {
                Write-Host ("`r$(" " * $script:LastProgressLen)`r") -NoNewline
                $script:LastProgressLen = 0
                Write-Host "    Part $partIndex/$partCount  $partName  $(Format-Bytes $partSize)  [done]" -ForegroundColor DarkGray
            } else {
                Show-Progress -Id 1 -ParentId 0 -Activity "Writing part $partIndex of $partCount" `
                    -Status "$partName - Done" -Percent 100
            }
        }
    } finally {
        $outStream.Close()
        Clear-AllProgress
    }
}

function Invoke-7Zip {
    param([string]$FirstPart)
    $7z = Find-7Zip
    if (-not $7z) {
        Write-Fail "7-Zip not found. Install it or set -SevenZipPath."
        return $false
    }
    Write-Info "Running 7-Zip on: $(Split-Path $FirstPart -Leaf)"
    if ($DryRun) { Write-Dry "7z x `"$FirstPart`" -o`"$OutputFolder`""; return $true }
    $result = & $7z x "$FirstPart" "-o$OutputFolder" -y 2>&1
    if ($LASTEXITCODE -eq 0) { return $true }
    Write-Fail "7-Zip error:`n$result"
    return $false
}

# =============================================================================
#  SCAN
# =============================================================================

function Get-ArchiveGroups {
    param([string]$Path)
    $allFiles = Get-ChildItem -LiteralPath $Path -File
    $groups   = @{}

    foreach ($file in $allFiles) {
        $name = $file.Name
        $full = $file.FullName

        # Pattern 1: file.zip.001 / file.rar.001 / file.tar.001
        if ($name -match '^(.+\.(zip|rar|tar|gz|7z))\.\d+$') {
            $base = $Matches[1]
            if (-not $groups[$base]) { $groups[$base] = @{ Type = "BinaryParts"; Parts = @() } }
            $groups[$base].Parts += $full; continue
        }
        # Pattern 2: file.001 / file.002 (HJSplit / generic)
        if ($name -match '^(.+)\.\d{3,}$') {
            $base = $Matches[1]
            if (-not $groups[$base]) { $groups[$base] = @{ Type = "BinaryParts"; Parts = @() } }
            $groups[$base].Parts += $full; continue
        }
        # Pattern 3: file.part1.rar / file.part01.rar
        if ($name -match '^(.+)\.part\d+\.rar$') {
            $base = $Matches[1]
            if (-not $groups[$base]) { $groups[$base] = @{ Type = "MultiPartRar"; Parts = @() } }
            $groups[$base].Parts += $full; continue
        }
        # Pattern 4: old-style RAR -- file.rar + file.r00, file.r01
        if ($name -match '^(.+)\.(rar|r\d+)$') {
            $base = $Matches[1]
            if (-not $groups[$base]) { $groups[$base] = @{ Type = "OldStyleRar"; Parts = @() } }
            $groups[$base].Parts += $full; continue
        }
        # Pattern 5: split ZIP -- file.z01, file.z02 ... + file.zip (WinZip/7-Zip split)
        if ($name -match '^(.+)\.(z\d+|zip)$') {
            $base = $Matches[1]
            if (-not $groups[$base]) { $groups[$base] = @{ Type = "SplitZip"; Parts = @() } }
            $groups[$base].Parts += $full; continue
        }
    }

    return $groups.GetEnumerator() |
           Where-Object { $_.Value.Parts.Count -gt 1 } |
           ForEach-Object { [PSCustomObject]@{
               Name  = $_.Key
               Type  = $_.Value.Type
               Parts = ($_.Value.Parts | Sort-Object)
           }}
}

# =============================================================================
#  PROCESS ONE GROUP
# =============================================================================

function Invoke-CombineGroup {
    param([PSCustomObject]$Group)

    $base  = $Group.Name
    $type  = $Group.Type
    $parts = $Group.Parts

    switch ($type) {
        "BinaryParts" {
            $outFile = Join-Path $OutputFolder $base
            if (Test-Path $outFile) {
                Write-Info "Output already exists, skipping: $base"
                return "Skipped"
            }
            if ($DryRun) { Write-Dry "Would concat $($parts.Count) parts -> $outFile"; return "DryRun" }
            try {
                Join-BinaryParts -Parts $parts -OutputFile $outFile -ArchiveName $base
                Write-OK "Created: $outFile  ($(Format-Bytes ([long](Get-Item -LiteralPath $outFile).Length)))"
                if ($DeletePartsAfter) {
                    $parts | ForEach-Object { Remove-Item $_ -Force; Write-Info "Deleted: $(Split-Path $_ -Leaf)" }
                }
                return "Success"
            } catch {
                Write-Fail "Failed: $_"
                if (Test-Path $outFile) { Remove-Item $outFile -Force }
                return "Failed"
            }
        }
        { $_ -in "MultiPartRar","OldStyleRar" } {
            $firstPart = if ($type -eq "OldStyleRar") {
                $parts | Where-Object { $_ -match '\.rar$' } | Select-Object -First 1
            } else { $parts[0] }
            if (-not $firstPart) { $firstPart = $parts[0] }
            if ($DryRun) { Write-Dry "Would run 7-Zip on: $(Split-Path $firstPart -Leaf)"; return "DryRun" }
            if (Invoke-7Zip -FirstPart $firstPart) {
                Write-OK "Extracted: $base"
                if ($DeletePartsAfter) {
                    $parts | ForEach-Object { Remove-Item $_ -Force; Write-Info "Deleted: $(Split-Path $_ -Leaf)" }
                }
                return "Success"
            }
            return "Failed"
        }
        # Split ZIP (.z01 / .z02 ... + .zip) — pass the .zip to 7-Zip which handles the rest
        "SplitZip" {
            $firstPart = $parts | Where-Object { $_ -match '\.zip$' } | Select-Object -First 1
            if (-not $firstPart) { $firstPart = ($parts | Sort-Object)[0] }
            if ($DryRun) { Write-Dry "Would run 7-Zip on: $(Split-Path $firstPart -Leaf)"; return "DryRun" }
            if (Invoke-7Zip -FirstPart $firstPart) {
                Write-OK "Extracted: $base"
                if ($DeletePartsAfter) {
                    $parts | ForEach-Object { Remove-Item $_ -Force; Write-Info "Deleted: $(Split-Path $_ -Leaf)" }
                }
                return "Success"
            }
            return "Failed"
        }
    }
}

# =============================================================================
#  RUN ALL SELECTED GROUPS
# =============================================================================

function Invoke-AllGroups {
    param([object[]]$Groups)
    $results = @{ Success = 0; Skipped = 0; Failed = 0; DryRun = 0 }
    foreach ($group in $Groups) {
        Write-Host ""
        Write-Host "  --> $($group.Name)" -ForegroundColor Cyan
        Write-Host "      $($group.Parts.Count) parts  |  $($group.Type)" -ForegroundColor DarkGray
        $outcome = Invoke-CombineGroup -Group $group
        $results[$outcome]++
    }
    return $results
}

# =============================================================================
#  SUMMARY
# =============================================================================

function Show-Summary {
    param([hashtable]$Results, [bool]$WasDryRun = $false)
    if ($WasDryRun) { $summaryTitle = "Dry Run Complete" } else { $summaryTitle = "Complete" }
    Write-Host ""
    Draw-Line
    Draw-Blank
    Draw-Title $summaryTitle
    Draw-Blank
    Draw-Line
    Write-Host ""
    Write-Host "  Success  : $($Results.Success)" -ForegroundColor Green
    Write-Host "  Skipped  : $($Results.Skipped)" -ForegroundColor Yellow
    Write-Host "  Failed   : $($Results.Failed)"  -ForegroundColor Red
    if ($WasDryRun) {
        Write-Host "  Dry runs : $($Results.DryRun)" -ForegroundColor Magenta
    }
    Write-Host ""
    Draw-Line
    Write-Host ""
}

# =============================================================================
#  SETTINGS DISPLAY
# =============================================================================

function Show-Settings {
    Draw-Section "Current Settings"
    $sz = Find-7Zip

    # Pre-compute values and colours for PS 5.1 compatibility
    # (inline `if` expressions inside function arguments are not supported in PS 5.1)

    if ($FolderPath)   { $srcVal = $FolderPath;   $srcCol = "White"  }
    else               { $srcVal = "(not set)";    $srcCol = "Red"    }

    if ($OutputFolder) { $outVal = $OutputFolder }
    else               { $outVal = "(same as source)" }

    if ($SevenZipPath) { $szVal = $SevenZipPath;      $szCol = "Green"  }
    elseif ($sz)       { $szVal = "$sz (auto)";        $szCol = "Green"  }
    else               { $szVal = "Not found";         $szCol = "Yellow" }

    if ($DryRun)           { $drVal = "Yes"; $drCol = "Magenta" }
    else                   { $drVal = "No";  $drCol = "Gray"    }

    if ($DeletePartsAfter) { $dpVal = "Yes"; $dpCol = "Yellow" }
    else                   { $dpVal = "No";  $dpCol = "Gray"   }

    if ($script:IsISE) { $pgVal = "Text fallback (ISE)" }
    else               { $pgVal = "Progress bars"       }

    Write-Status "Source folder"  $srcVal $srcCol
    Write-Status "Output folder"  $outVal "Gray"
    Write-Status "7-Zip path"     $szVal  $szCol
    Write-Status "Buffer size"    "$BufferSizeMB MB"
    Write-Status "Dry run"        $drVal  $drCol
    Write-Status "Delete parts"   $dpVal  $dpCol
    Write-Status "Progress mode"  $pgVal  "DarkGray"
}

# =============================================================================
#  INTERACTIVE FLOW
# =============================================================================

function Start-InteractiveMode {

    # ── Welcome ────────────────────────────────────────────────────────────────
    Draw-Header
    Draw-Section "Welcome"
    Write-Host "  This tool combines split archive files." -ForegroundColor Gray
    Write-Host "  Supports .zip.001, .part1.rar, .001 and more." -ForegroundColor Gray
    Write-Host ""
    Write-Host "  Run with -Help for command-line / scripting usage." -ForegroundColor DarkGray
    Write-Host ""

    # ── Step 1: Source folder ──────────────────────────────────────────────────
    Draw-Section "Step 1 of 5 - Source Folder"
    Write-Host "  Enter the folder containing your split archives." -ForegroundColor Gray
    Write-Host "  Press Enter to use the current directory." -ForegroundColor DarkGray
    Write-Host ""

    $inputPath = Prompt-Input "Source folder" (Get-Location).Path
    if (-not (Test-Path $inputPath)) {
        Write-Fail "Folder not found: $inputPath"
        Pause-Screen; return
    }
    $script:FolderPath = (Resolve-Path $inputPath).Path

    # ── Scan ───────────────────────────────────────────────────────────────────
    Write-Host ""
    Write-Host "  Scanning..." -ForegroundColor DarkGray
    $allGroups = @(Get-ArchiveGroups -Path $FolderPath)

    Draw-Header

    if (-not $allGroups -or $allGroups.Count -eq 0) {
        Draw-Section "Scan Results"
        Write-Host "  No split archives found in:" -ForegroundColor Yellow
        Write-Host "  $FolderPath" -ForegroundColor Gray
        Pause-Screen; return
    }

    # ── Step 2: Select archives ────────────────────────────────────────────────
    Draw-Section "Step 2 of 5 - Select Archives"
    Write-Host "  Found $($allGroups.Count) archive set(s) in:" -ForegroundColor Gray
    Write-Host "  $FolderPath" -ForegroundColor DarkGray
    Write-Host ""

    $groups = @(Select-Archives -Groups $allGroups)

    if (-not $groups -or $groups.Count -eq 0) {
        Pause-Screen; return
    }

    # ── Step 3: Output folder ──────────────────────────────────────────────────
    Draw-Header
    Draw-Section "Step 3 of 5 - Output Folder"
    Write-Host "  Where should combined files be saved?" -ForegroundColor Gray
    Write-Host "  Press Enter to save in the same folder as the source." -ForegroundColor DarkGray
    Write-Host ""

    $outInput = Prompt-Input "Output folder" $FolderPath
    if (-not (Test-Path $outInput)) {
        $create = Prompt-YesNo "Folder does not exist. Create it?" $true
        if ($create) {
            New-Item -ItemType Directory -Path $outInput | Out-Null
            Write-OK "Created: $outInput"
        } else {
            Write-Fail "Output folder not set. Aborting."
            Pause-Screen; return
        }
    }
    $script:OutputFolder = (Resolve-Path $outInput).Path

    # ── Step 4: Options ────────────────────────────────────────────────────────
    Draw-Header
    Draw-Section "Step 4 of 5 - Options"

    $script:DryRun           = Prompt-YesNo "Dry run? (preview only, no files written)" $false
    $script:DeletePartsAfter = Prompt-YesNo "Delete part files after successful combine?" $false

    Write-Host ""
    $bufInput = Prompt-Input "Buffer size in MB" "64"
    if ($bufInput -match '^\d+$') { $script:BufferSizeMB = [int]$bufInput }

    # ── Step 5: Confirm ────────────────────────────────────────────────────────
    Draw-Header
    Show-Settings

    # Show selected archives
    Draw-Section "Selected Archives ($($groups.Count) of $($allGroups.Count))"
    $typeLabel = @{ BinaryParts = "ZIP/Binary"; MultiPartRar = "Multi-RAR "; OldStyleRar = "Old RAR   "; SplitZip = "Split ZIP " }
    foreach ($g in $groups) {
        $totalSize = Format-Bytes ([long](($g.Parts | ForEach-Object { (Get-Item -LiteralPath $_).Length }) | Measure-Object -Sum).Sum)
        Write-Host "  [" -ForegroundColor DarkCyan -NoNewline
        Write-Host $typeLabel[$g.Type] -ForegroundColor Cyan -NoNewline
        Write-Host "]  " -ForegroundColor DarkCyan -NoNewline
        Write-Host $g.Name -ForegroundColor White -NoNewline
        Write-Host "  ($($g.Parts.Count) parts, $totalSize)" -ForegroundColor DarkGray
    }
    Write-Host ""

    $confirm = Prompt-YesNo "Ready to proceed?" $true
    if (-not $confirm) {
        Write-Host ""
        Write-Info "Cancelled. No files were changed."
        Write-Host ""
        return
    }

    # ── Run ────────────────────────────────────────────────────────────────────
    Draw-Header
    Draw-Section "Step 5 of 5 - Combining"

    $results = Invoke-AllGroups -Groups $groups

    # ── Summary + dry-run follow-up ────────────────────────────────────────────
    Show-Summary -Results $results -WasDryRun $DryRun

    if ($DryRun) {
        Write-Host "  Everything above was a preview. No files were written." -ForegroundColor Magenta
        Write-Host ""
        $proceed = Prompt-YesNo "Would you like to proceed with a real run using the same settings?" $false

        if ($proceed) {
            $script:DryRun = $false
            Draw-Header
            Draw-Section "Running for real now"
            Show-Settings
            Write-Host ""
            $confirm2 = Prompt-YesNo "Confirm and start combining?" $true
            if ($confirm2) {
                Draw-Header
                Draw-Section "Combining"
                $realResults = Invoke-AllGroups -Groups $groups
                Show-Summary -Results $realResults -WasDryRun $false
            } else {
                Write-Host ""
                Write-Info "Cancelled. No files were changed."
                Write-Host ""
            }
        } else {
            Write-Host ""
            Write-Info "Exiting. No files were changed."
            Write-Host ""
        }
    }
}

# =============================================================================
#  ENTRY POINT
# =============================================================================

if (-not $NonInteractive -and [string]::IsNullOrWhiteSpace($FolderPath)) {
    Start-InteractiveMode
    exit 0
}

# ── Non-interactive / scripted mode ───────────────────────────────────────────
if ([string]::IsNullOrWhiteSpace($FolderPath)) { $FolderPath = (Get-Location).Path }
$FolderPath = (Resolve-Path $FolderPath).Path
if (-not $OutputFolder) { $OutputFolder = $FolderPath }
if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder | Out-Null }

Write-Host ""
Write-Host "  Combine-Archives  [Non-interactive]" -ForegroundColor Cyan
Show-Settings

$groups = @(Get-ArchiveGroups -Path $FolderPath)

if (-not $groups -or $groups.Count -eq 0) {
    Write-Info "No split archive sets found in: $FolderPath"
    exit 0
}

$results = Invoke-AllGroups -Groups $groups
Show-Summary -Results $results -WasDryRun $DryRun

if ($DryRun) {
    Write-Host "  Everything above was a preview. No files were written." -ForegroundColor Magenta
    Write-Host ""
    Write-Host "  Would you like to proceed with a real run using the same settings? [y/N]" -ForegroundColor Cyan -NoNewline
    Write-Host " > " -ForegroundColor DarkCyan -NoNewline
    $ans = Read-Host
    if ($ans -match '^[Yy]') {
        $script:DryRun = $false
        $DryRun        = $false
        Write-Host ""
        Write-Host "  Running for real now..." -ForegroundColor Cyan
        $realResults = Invoke-AllGroups -Groups $groups
        Show-Summary -Results $realResults -WasDryRun $false
    } else {
        Write-Host ""
        Write-Info "Exiting. No files were changed."
        Write-Host ""
    }
}
