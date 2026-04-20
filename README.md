# Combine-Archives.ps1 · v1.3

A PowerShell script that automatically scans a folder for split archive parts and combines or extracts them — with an interactive step-by-step menu, archive selection, live progress bars, speed readout, and ETA.

---

## Features

- Interactive TUI menu when run without parameters
- Numbered archive selection — choose which sets to process
- Detects and groups split archive sets automatically
- Supports files larger than 2 GB via chunked streaming
- Progress bars in Windows Terminal / console; text fallback in PowerShell ISE
- Dry-run mode with option to immediately proceed with a real run afterwards
- Optional cleanup of part files after a successful combine
- No external dependencies for ZIP/binary splits
- Auto-detects 7-Zip for RAR extraction
- Handles filenames with special characters including brackets `[ ]`
- Full command-line parameter support for scripting/automation

---

## Supported Formats

| Pattern | Example |
|---|---|
| Generic binary splits | `archive.zip.001`, `archive.zip.002` ... |
| HJSplit / 7-Zip splits | `archive.001`, `archive.002` ... |
| Multi-part RAR | `archive.part1.rar`, `archive.part01.rar` ... |
| Old-style RAR | `archive.rar` + `archive.r00`, `archive.r01` ... |
| Split ZIP (WinZip/7-Zip) | `archive.z01`, `archive.z02` ... + `archive.zip` |

ZIP and generic binary splits are handled entirely by PowerShell — no extra tools needed.
Split ZIP and RAR files require [7-Zip](https://www.7-zip.org) to be installed.

---

## Requirements

- Windows PowerShell 5.1 or PowerShell 7+
- [7-Zip](https://www.7-zip.org) *(required for RAR and Split ZIP `.z01` files)*

---

## Installation

1. Download `Combine-Archives.ps1` to a folder of your choice (e.g. `C:\Scripts\`)

2. **Unblock the file** — Windows automatically marks files downloaded from the internet
   as untrusted. Run this once after downloading:

```powershell
Unblock-File "C:\Scripts\Combine-Archives.ps1"
```

> This removes the hidden **Mark of the Web** tag Windows attached when you downloaded
> the file. It is the command-line equivalent of right-clicking the file → Properties →
> ticking the **Unblock** checkbox. It only affects this one file and is the recommended
> approach.

3. **Allow local scripts to run** *(if not already set — one-time, per user)*:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

> This tells PowerShell to allow local scripts to run while still blocking unsigned
> scripts downloaded from the internet. Only needs to be set once per user account.

### Which one do I need?

| Situation | Solution |
|---|---|
| Script is blocked after downloading | `Unblock-File` ← **recommended, surgical** |
| Script runs but execution policy error appears | `Set-ExecutionPolicy RemoteSigned` |
| Both errors appear | Run both commands, `Unblock-File` first |

Both are safe as long as you trust the script. `Unblock-File` is preferred because it
targets only the specific file rather than relaxing policy for your entire account.

---

## Usage

### Interactive mode (recommended)

Just run the script with no arguments:

```powershell
.\Combine-Archives.ps1
```

You will be guided through 5 steps:

```
  Step 1 of 5 - Source Folder
  Step 2 of 5 - Select Archives
  Step 3 of 5 - Output Folder
  Step 4 of 5 - Options
  Step 5 of 5 - Combining
```

### Archive Selection

After scanning, all detected archive sets are shown as a numbered list.
You can choose exactly which ones to process:

```
  Step 2 of 5 - Select Archives
  --------------------------------
  Found 4 archive set(s) in: D:\Downloads

  [ ##]  Type          Name                  Pts  Size      Cumulative
  ------------------------------------------------------------------
  [  1]  ZIP/Binary    movie.zip               5  4.80 GB   4.80 GB
  [  2]  Multi-RAR     tv.show.part1.rar       8  7.20 GB  12.00 GB
  [  3]  ZIP/Binary    software.zip            2  980 MB   12.96 GB
  [  4]  Old RAR       backup.rar              3  2.10 GB  15.06 GB
  ------------------------------------------------------------------
  Total if all selected: 15.06 GB

  Selection > 1,3          <- individual numbers
  Selection > 2-4          <- a range
  Selection > 1,3-4        <- mix of both
  Selection >              <- press Enter to select all
  Selection > a            <- also selects all
  Selection > n            <- cancel / select none
```

### After a Dry Run

When dry-run mode is used, the script will ask before exiting whether you want to immediately proceed with a real run — no need to re-enter anything:

```
  Dry Run Complete
  ----------------
  Dry runs : 2

  Everything above was a preview. No files were written.

  Would you like to proceed with a real run using the same settings? [y/N] >
```

### Command-line / scripted mode

Pass `-FolderPath` (or any parameter) to skip the interactive menu entirely.
In this mode all detected archives are processed without prompting.

```powershell
.\Combine-Archives.ps1 [[-FolderPath] <string>] [[-OutputFolder] <string>]
                       [[-SevenZipPath] <string>] [[-BufferSizeMB] <int>]
                       [-DryRun] [-DeletePartsAfter] [-NonInteractive]
```

---

## Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-FolderPath` | string | *(interactive prompt)* | Folder to scan for split archive parts |
| `-OutputFolder` | string | same as `-FolderPath` | Folder to save combined/extracted files |
| `-SevenZipPath` | string | auto-detected | Full path to `7z.exe` if not in a standard location |
| `-BufferSizeMB` | int | `64` | Read/write buffer size in MB (see table below) |
| `-DryRun` | switch | off | Preview mode — shows what would happen without doing anything |
| `-DeletePartsAfter` | switch | off | Delete part files after a successful combine or extract |
| `-NonInteractive` | switch | off | Force non-interactive mode even without parameters |

---

## Buffer Size Guide

The buffer controls how much data is read and written at a time. A larger buffer can improve throughput on fast drives, but uses more RAM.

| `-BufferSizeMB` | RAM used | Recommended for |
|---|---|---|
| `16` | 16 MB | Low-end systems or VMs with under 4 GB RAM |
| `32` | 32 MB | Systems with 4 GB RAM |
| `64` *(default)* | 64 MB | Systems with 8 GB RAM — good general default |
| `128` | 128 MB | Systems with 16 GB RAM and fast NVMe/SSD |
| `256` | 256 MB | Systems with 32 GB+ RAM, large archives, fast storage |
| `512` | 512 MB | Workstations with 64 GB+ RAM, maximum throughput |

> **Note:** Going above `256` rarely gives a meaningful speed boost unless you are working with very large files on very fast storage. The bottleneck is usually disk speed, not buffer size.

---

## Progress Display

The script automatically detects the environment and adjusts accordingly.

**Windows Terminal / PowerShell console** — native nested progress bars:
```
Combining: archive.zip
[=================>        ] 68%   Part 3 of 5  |  4.2 GB of 6.1 GB

  Writing part 3 of 5
  [=============>            ] 54%
  archive.zip.003  |  891.2 MB / 1.2 GB  |  487.3 MB/s  |  ~2s left
```

**PowerShell ISE** — inline text progress (ISE does not support native bars):
```
  Combining: archive.zip  (5 parts, 4.8 GB total)
    Part 1/5  archive.zip.001  1.2 GB  [done]
    Part 2/5  archive.zip.002  1.2 GB  [done]
  [##############----------------]  47%  archive.zip.003 | 487.3 MB/s | ~2s left
```

---

## Examples

**Interactive mode:**
```powershell
.\Combine-Archives.ps1
```

**Scan a specific folder:**
```powershell
.\Combine-Archives.ps1 -FolderPath "D:\Downloads"
```

**Preview without combining (then optionally proceed):**
```powershell
.\Combine-Archives.ps1 -FolderPath "D:\Downloads" -DryRun
```

**Save combined files to a different folder:**
```powershell
.\Combine-Archives.ps1 -FolderPath "D:\Downloads" -OutputFolder "D:\Combined"
```

**Combine and delete parts afterwards:**
```powershell
.\Combine-Archives.ps1 -FolderPath "D:\Downloads" -DeletePartsAfter
```

**Use a custom 7-Zip path:**
```powershell
.\Combine-Archives.ps1 -FolderPath "D:\Downloads" -SevenZipPath "C:\Tools\7z.exe"
```

**Use a larger buffer for faster copying:**
```powershell
.\Combine-Archives.ps1 -FolderPath "D:\Downloads" -BufferSizeMB 128
```

---

## Built-in Help

```powershell
Get-Help .\Combine-Archives.ps1
Get-Help .\Combine-Archives.ps1 -Full
Get-Help .\Combine-Archives.ps1 -Examples
```

---

## Notes

- Run without any parameters for the guided interactive menu
- Archive selection supports individual numbers (`1,3`), ranges (`2-4`), or a mix (`1,3-5`)
- Press Enter or type `a` at the selection prompt to process all detected archives
- After a dry run, the script offers to immediately proceed with a real run — no re-entry needed
- If an output file already exists it will be skipped — delete it first to re-combine
- If a combine fails partway through, any incomplete output file is automatically removed
- In non-interactive / scripted mode all detected archives are processed without prompting

---

## Changelog

### v1.3
- **Docs:** Expanded Installation section — explains `Unblock-File` vs `Set-ExecutionPolicy`,
  when to use each, and why neither presents a security risk when you trust the script

### v1.2
- **Fix:** All file size calculations now use explicit `[long]` (64-bit) casting throughout —
  previously `Measure-Object -Sum` could silently downcast `Int64` to `Int32`, causing an
  overflow error on archives larger than ~2.1 GB total
- **Fix:** Replaced all `Get-Item` and `Get-ChildItem` calls with `-LiteralPath` variants —
  previously filenames containing square brackets (e.g. `Purple Teaming [BlackHat].7z`)
  were misinterpreted as PowerShell wildcard patterns, returning 0 bytes and failing to combine
- **Added:** Support for Split ZIP format (`.z01`, `.z02` ... + `.zip`) created by WinZip and 7-Zip

### v1.1
- **Added:** Interactive TUI menu (5-step guided flow)
- **Added:** Numbered archive selection with range support (`1,3`, `2-4`, `1,3-5`)
- **Added:** Size and cumulative size columns in archive selection list
- **Added:** Dry-run follow-up prompt — offer to proceed with real run immediately after preview
- **Added:** ISE detection with text-based progress fallback
- **Added:** `-NonInteractive` switch for scripted/scheduled use
- **Fix:** `Show-Settings` inline `if` expressions rewritten for PowerShell 5.1 compatibility

### v1.0
- Initial release
- Chunked streaming for files larger than 2 GB
- Dual nested progress bars with MB/s speed and ETA
- Supports ZIP binary splits, HJSplit, multi-part RAR, old-style RAR
- Auto-detects 7-Zip
- `-DryRun`, `-DeletePartsAfter`, `-BufferSizeMB` parameters
- Full `Get-Help` documentation block

---

## License

MIT — free to use, modify, and distribute.
