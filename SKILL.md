---
name: beyond-compare-document-reports
description: Generate headless Beyond Compare 5 reports on Windows for document, spreadsheet, text, binary, and folder comparisons. Use when Codex needs to compare files or folders with Beyond Compare without opening the GUI, especially for .doc, .docx, .xls, .xlsx, .txt, CSV-like files, binary same/different checks, report files, HTML text reports, or folder XML reports.
---

# Beyond Compare Document Reports

## Core Rule

Use the hardcoded Windows install path:

```powershell
$BComp = "J:\Program Files\Beyond Compare 5\BComp.com"
```

Do not search for the Beyond Compare installation unless this path fails.

Do not use `BCompare.exe`, `BComp.exe`, or direct commands like `BComp.com left right` for report-only work. Those can open GUI windows.

## Default Workflow

1. Identify whether the inputs are files or folders.
2. Use `scripts/New-BeyondCompareReport.ps1` for headless report generation.
3. Inspect the generated report for results and `Conversion Error`.
4. Use `/qc=binary` only when the user needs a quick same/different exit code for two files.
5. Read `references/beyond-compare-5-windows-cli-report-guide.md` only when needing exact switch behavior, exit-code nuance, or examples.

## Generate Reports

Use the bundled helper script:

```powershell
& ".\scripts\New-BeyondCompareReport.ps1" -Left "C:\left.docx" -Right "C:\right.docx" -Output "C:\report.txt"
```

For folder XML:

```powershell
& ".\scripts\New-BeyondCompareReport.ps1" -Mode Folder -Left "C:\left_folder" -Right "C:\right_folder" -Output "C:\folder_report.xml"
```

For HTML text reports:

```powershell
& ".\scripts\New-BeyondCompareReport.ps1" -Mode TextHtml -Left "C:\left.txt" -Right "C:\right.txt" -Output "C:\report.html"
```

## Quick Same/Different Checks

For two files only:

```powershell
& "J:\Program Files\Beyond Compare 5\BComp.com" /qc=binary "C:\left.file" "C:\right.file"
$LASTEXITCODE
```

Important exit codes:

| Code | Meaning |
|---:|---|
| `1` | Binary same |
| `2` | Rules-based same, but not binary-identical |
| `11` | Binary differences |
| `12` | Similar |
| `13` | Rules-based differences |
| `100` | Error |
| `105` | Script file load error |
| `106` | Script syntax error |
| `107` | Script failed to load folders or files |

Avoid `/qc` for folders. Generate a folder XML report instead.

## Office Files

Beyond Compare can generate reports for `.doc`, `.docx`, `.xls`, and `.xlsx`, but conversion can fail depending on file contents and installed converters.

After report generation, check:

```powershell
$report = Get-Content "C:\report.txt" -Raw
$report -match "Conversion Error"
```

If `Conversion Error` appears, report that Beyond Compare could not convert one or both files for content comparison. Do not claim a meaningful document/spreadsheet content diff from that report.

## HTML Reports

A `.html` output filename is not enough. The BC script must include:

```text
output-options:html-color
```

The bundled helper adds this automatically for `-Mode TextHtml`.

## Safety

Never run mirror sync, delete, move, or direct GUI compare commands unless the user explicitly asks for that destructive or interactive behavior.

