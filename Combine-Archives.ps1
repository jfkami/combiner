<#
.SYNOPSIS
    Scans a folder for split archive parts and combines or extracts them.

.DESCRIPTION
    Combine-Archives.ps1 automatically detects split archive sets in a folder
    and combines or extracts them. It supports the following formats:

      - Generic binary splits  : file.zip.001, file.zip.002 ...
      - HJSplit / 7-Zip splits : file.001, file.002 ...
      - Multi-part RAR         : file.part1.rar, file.part01.rar ...
      - Old-style RAR          : file.rar + file.r00, file.r01 ...

    ZIP/binary splits are joined using pure PowerShell (no tools required).
    RAR files require 7-Zip to be installed (auto-detected in default locations).
    Uses chunked streaming so files larger than 2GB are handled correctly.
    Displays a progress bar per part file and an overall progress bar.

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

.EXAMPLE
    .\Combine-Archives.ps1

    Scans the current directory and combines any split archives found.

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
    - Run Set-ExecutionPolicy RemoteSigned -Scope CurrentUser if scripts are blocked
    - Use -DryRun first to verify the script detects your files correctly

.LINK
    https://www.7-zip.org
#>

param (
    [string] $FolderPath   = ".",
    [string] $OutputFolder = "",
    [string] $SevenZipPath = "",
    [int]    $BufferSizeMB = 64,
    [switch] $DryRun,
    [switch] $DeletePartsAfter
)

# ── Colour helpers ─────────────────────────────────────────────────────────────
function Write-Header { param($m) Write-Host "`n==> $m" -ForegroundColor Cyan }
function Write-OK     { param($m) Write-Host "  [OK] $m"  -ForegroundColor Green }
function Write-Info   { param($m) Write-Host "  [..] $m"  -ForegroundColor Yellow }
function Write-Fail   { param($m) Write-Host "  [!!] $m"  -ForegroundColor Red }
function Write-Dry    { param($m) Write-Host "  [DRY] $m" -ForegroundColor Magenta }

# ── Format bytes helper ────────────────────────────────────────────────────────
function Format-Bytes {
    param([long]$Bytes)
    if     ($Bytes -ge 1GB) { "{0:N2} GB" -f ($Bytes / 1GB) }
    elseif ($Bytes -ge 1MB) { "{0:N1} MB" -f ($Bytes / 1MB) }
    elseif ($Bytes -ge 1KB) { "{0:N1} KB" -f ($Bytes / 1KB) }
    else                    { "$Bytes B" }
}

# ── Resolve paths ──────────────────────────────────────────────────────────────
$FolderPath = (Resolve-Path $FolderPath).Path
if (-not $OutputFolder) { $OutputFolder = $FolderPath }
if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder | Out-Null }

# ── Locate 7-Zip ───────────────────────────────────────────────────────────────
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
$7z = Find-7Zip

# ── Chunked streaming concat with progress bars ────────────────────────────────
#
#   Progress bar layout:
#     ID 0  - Overall progress  (how many parts done out of total)
#     ID 1  - Current part      (bytes written for this part)
#
function Join-BinaryParts {
    param (
        [string[]] $Parts,
        [string]   $OutputFile,
        [string]   $ArchiveName
    )

    $bufferSize  = $BufferSizeMB * 1024 * 1024
    $buffer      = New-Object byte[] $bufferSize

    # Pre-calculate total bytes across all parts for the overall bar
    $totalBytes  = ($Parts | ForEach-Object { (Get-Item $_).Length } |
                    Measure-Object -Sum).Sum
    $writtenTotal = [long]0
    $partIndex    = 0
    $partCount    = $Parts.Count
    $stopwatch    = [System.Diagnostics.Stopwatch]::StartNew()

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
            $partSize    = (Get-Item $part).Length
            $writtenPart = [long]0

            # ── Overall progress bar (ID 0) ────────────────────────────────
            $overallPct = [math]::Min(
                [int](($writtenTotal / [math]::Max($totalBytes, 1)) * 100), 100)

            Write-Progress `
                -Id       0 `
                -Activity "Combining: $ArchiveName" `
                -Status   "Part $partIndex of $partCount  |  $(Format-Bytes $writtenTotal) of $(Format-Bytes $totalBytes)" `
                -PercentComplete $overallPct

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

                    # ── Per-part progress bar (ID 1) ───────────────────────
                    $partPct = [math]::Min(
                        [int](($writtenPart / [math]::Max($partSize, 1)) * 100), 100)

                    # Calculate speed (MB/s)
                    $elapsed = $stopwatch.Elapsed.TotalSeconds
                    if ($elapsed -gt 0) {
                        $speedMBs  = [math]::Round(($writtenTotal / 1MB) / $elapsed, 1)
                        $remaining = if ($speedMBs -gt 0) {
                            $bytesLeft = $totalBytes - $writtenTotal
                            [math]::Round(($bytesLeft / 1MB) / $speedMBs)
                        } else { 0 }
                        $statusMsg = "$partName  |  $(Format-Bytes $writtenPart) / $(Format-Bytes $partSize)  |  $speedMBs MB/s  |  ~$remaining s remaining"
                    } else {
                        $statusMsg = "$partName  |  $(Format-Bytes $writtenPart) / $(Format-Bytes $partSize)"
                    }

                    Write-Progress `
                        -Id               1 `
                        -ParentId         0 `
                        -Activity         "Writing part $partIndex of $partCount" `
                        -Status           $statusMsg `
                        -PercentComplete  $partPct
                }
            } finally {
                $inStream.Close()
            }

            # Mark part complete on the child bar
            Write-Progress -Id 1 -ParentId 0 -Activity "Writing part $partIndex of $partCount" `
                           -Status "$partName  - Done" -PercentComplete 100
        }

    } finally {
        $outStream.Close()

        # Clear both progress bars
        Write-Progress -Id 1 -Activity "Done" -Completed
        Write-Progress -Id 0 -Activity "Done" -Completed
    }
}

# ── Extract with 7-Zip ─────────────────────────────────────────────────────────
function Invoke-7Zip {
    param ([string] $FirstPart)
    if (-not $7z) {
        Write-Fail "7-Zip not found. Install it or set -SevenZipPath. Skipping: $FirstPart"
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
#  SCAN AND GROUP
# =============================================================================
Write-Header "Scanning: $FolderPath"

$allFiles = Get-ChildItem -Path $FolderPath -File
$groups   = @{}

foreach ($file in $allFiles) {
    $name = $file.Name
    $full = $file.FullName

    # Pattern 1: file.zip.001 / file.rar.001 / file.tar.001
    if ($name -match '^(.+\.(zip|rar|tar|gz|7z))\.\d+$') {
        $base = $Matches[1]
        if (-not $groups[$base]) { $groups[$base] = @{ Type = "BinaryParts"; Parts = @() } }
        $groups[$base].Parts += $full
        continue
    }

    # Pattern 2: file.001 / file.002 (HJSplit / generic)
    if ($name -match '^(.+)\.\d{3,}$') {
        $base = $Matches[1]
        if (-not $groups[$base]) { $groups[$base] = @{ Type = "BinaryParts"; Parts = @() } }
        $groups[$base].Parts += $full
        continue
    }

    # Pattern 3: file.part1.rar / file.part01.rar
    if ($name -match '^(.+)\.part\d+\.rar$') {
        $base = $Matches[1]
        if (-not $groups[$base]) { $groups[$base] = @{ Type = "MultiPartRar"; Parts = @() } }
        $groups[$base].Parts += $full
        continue
    }

    # Pattern 4: old-style RAR -- file.rar + file.r00, file.r01
    if ($name -match '^(.+)\.(rar|r\d+)$') {
        $base = $Matches[1]
        if (-not $groups[$base]) { $groups[$base] = @{ Type = "OldStyleRar"; Parts = @() } }
        $groups[$base].Parts += $full
        continue
    }
}

# Filter out lone files (no actual split set)
$groups = $groups.GetEnumerator() | Where-Object { $_.Value.Parts.Count -gt 1 } |
          ForEach-Object { @{ Key = $_.Key; Value = $_.Value } }

if (-not $groups) {
    Write-Info "No split archive sets found in: $FolderPath"
    exit 0
}

# =============================================================================
#  PROCESS EACH GROUP
# =============================================================================
$results = @{ Success = 0; Skipped = 0; Failed = 0 }

foreach ($entry in $groups) {
    $base  = $entry.Key
    $info  = $entry.Value
    $type  = $info.Type
    $parts = $info.Parts | Sort-Object

    Write-Header "Found: $base  ($($parts.Count) parts, type: $type)"
    foreach ($p in $parts) { Write-Info (Split-Path $p -Leaf) }

    switch ($type) {

        "BinaryParts" {
            $outFile = Join-Path $OutputFolder $base

            if (Test-Path $outFile) {
                Write-Info "Output already exists, skipping: $base"
                $results.Skipped++
                continue
            }

            if ($DryRun) {
                Write-Dry "Would concat $($parts.Count) parts -> $outFile"
                $results.Success++
                continue
            }

            try {
                Join-BinaryParts -Parts $parts -OutputFile $outFile -ArchiveName $base
                $finalSize = Format-Bytes (Get-Item $outFile).Length
                Write-OK "Created: $outFile  ($finalSize)"
                $results.Success++
                if ($DeletePartsAfter) {
                    $parts | ForEach-Object {
                        Remove-Item $_ -Force
                        Write-Info "Deleted: $(Split-Path $_ -Leaf)"
                    }
                }
            } catch {
                Write-Fail "Failed: $_"
                if (Test-Path $outFile) { Remove-Item $outFile -Force }
                $results.Failed++
            }
        }

        "MultiPartRar" {
            $firstPart = ($parts | Sort-Object)[0]
            if ($DryRun) {
                Write-Dry "Would run 7-Zip on: $(Split-Path $firstPart -Leaf)"
                $results.Success++
                continue
            }
            if (Invoke-7Zip -FirstPart $firstPart) {
                Write-OK "Extracted: $base"
                $results.Success++
                if ($DeletePartsAfter) {
                    $parts | ForEach-Object {
                        Remove-Item $_ -Force
                        Write-Info "Deleted: $(Split-Path $_ -Leaf)"
                    }
                }
            } else {
                $results.Failed++
            }
        }

        "OldStyleRar" {
            $firstPart = $parts | Where-Object { $_ -match '\.rar$' } | Select-Object -First 1
            if (-not $firstPart) { $firstPart = ($parts | Sort-Object)[0] }
            if ($DryRun) {
                Write-Dry "Would run 7-Zip on: $(Split-Path $firstPart -Leaf)"
                $results.Success++
                continue
            }
            if (Invoke-7Zip -FirstPart $firstPart) {
                Write-OK "Extracted: $base"
                $results.Success++
                if ($DeletePartsAfter) {
                    $parts | ForEach-Object {
                        Remove-Item $_ -Force
                        Write-Info "Deleted: $(Split-Path $_ -Leaf)"
                    }
                }
            } else {
                $results.Failed++
            }
        }
    }
}

# ── Summary ────────────────────────────────────────────────────────────────────
Write-Header "Done"
Write-Host "  Success : $($results.Success)" -ForegroundColor Green
Write-Host "  Skipped : $($results.Skipped)" -ForegroundColor Yellow
Write-Host "  Failed  : $($results.Failed)"  -ForegroundColor Red
