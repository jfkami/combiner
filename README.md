This powershell script was created using claude AI.

# Combine-Archives.ps1

A PowerShell script that automatically scans a folder for split archive parts and combines or extracts them — with live progress bars, speed readout, and ETA.

---

## Features

- Detects and groups split archive sets automatically
- Supports files larger than 2 GB via chunked streaming
- Dual nested progress bars with MB/s speed and time remaining
- Dry-run mode to preview without touching any files
- Optional cleanup of part files after a successful combine
- No external dependencies for ZIP/binary splits
- Auto-detects 7-Zip for RAR extraction

---

## Supported Formats

| Pattern | Example |
|---|---|
| Generic binary splits | `archive.zip.001`, `archive.zip.002` ... |
| HJSplit / 7-Zip splits | `archive.001`, `archive.002` ... |
| Multi-part RAR | `archive.part1.rar`, `archive.part01.rar` ... |
| Old-style RAR | `archive.rar` + `archive.r00`, `archive.r01` ... |

ZIP and generic binary splits are handled entirely by PowerShell — no extra tools needed.
RAR files require [7-Zip](https://www.7-zip.org) to be installed.

---

## Requirements

- Windows PowerShell 5.1 or PowerShell 7+
- [7-Zip](https://www.7-zip.org) *(only required for RAR files)*

---

## Installation

1. Download `Combine-Archives.ps1`
2. If required, allow script execution (one-time, current user only):

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

---

## Usage

```powershell
.\Combine-Archives.ps1 [[-FolderPath] <string>] [[-OutputFolder] <string>]
                       [[-SevenZipPath] <string>] [[-BufferSizeMB] <int>]
                       [-DryRun] [-DeletePartsAfter]
```

---

## Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-FolderPath` | string | `.` (current dir) | Folder to scan for split archive parts |
| `-OutputFolder` | string | same as `-FolderPath` | Folder to save combined/extracted files |
| `-SevenZipPath` | string | auto-detected | Full path to `7z.exe` if not in a standard location |
| `-BufferSizeMB` | int | `64` | Read/write buffer size in MB. Increase for faster copies on systems with more RAM |
| `-DryRun` | switch | off | Preview mode — shows what would happen without doing anything |
| `-DeletePartsAfter` | switch | off | Delete part files after a successful combine or extract |

---

## Examples

**Scan the current folder:**
```powershell
.\Combine-Archives.ps1
```

**Scan a specific folder:**
```powershell
.\Combine-Archives.ps1 -FolderPath "D:\Downloads"
```

**Preview without combining anything:**
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

**Combine, save elsewhere, and clean up:**
```powershell
.\Combine-Archives.ps1 -FolderPath "D:\Downloads" -OutputFolder "D:\Out" -DeletePartsAfter
```

**Use a larger buffer for faster copying:**
```powershell
.\Combine-Archives.ps1 -FolderPath "D:\Downloads" -BufferSizeMB 128
```

---

## Progress Output

While combining, two nested progress bars are shown:

```
Combining: archive.zip
[=================>        ] 68%   Part 3 of 5  |  4.2 GB of 6.1 GB

  Writing part 3 of 5
  [=============>            ] 54%
  archive.zip.003  |  891.2 MB / 1.2 GB  |  487.3 MB/s  |  ~2 s remaining
```

- **Outer bar** — overall progress across all parts (total bytes written vs total size)
- **Inner bar** — current part being written, with transfer speed and estimated time remaining

Both bars clear automatically on completion.

---

## Built-in Help

Full parameter documentation is available directly in PowerShell:

```powershell
Get-Help .\Combine-Archives.ps1
Get-Help .\Combine-Archives.ps1 -Full
Get-Help .\Combine-Archives.ps1 -Examples
```

---

## Notes

- Always use `-DryRun` first to confirm the script has detected your files correctly before combining
- If an output file already exists it will be skipped — delete it first to re-combine
- If a combine fails partway through, any incomplete output file is automatically removed
- The script processes multiple archive sets in a single run if more than one is found in the folder

---

## License

MIT — free to use, modify, and distribute.
