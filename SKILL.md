---
name: document-compare
description: Generate headless Beyond Compare 5 reports on Windows. Use for document, spreadsheet, text, binary, or folder comparisons, especially .doc, .docx, .xls, .xlsx, .txt, HTML reports, XML folder reports, or same/different checks without opening the GUI.
---

# Document Compare

Use the hardcoded console helper:

```powershell
$BComp = "J:\Program Files\Beyond Compare 5\BComp.com"
```

Do not search for Beyond Compare unless that path fails. Do not use `BCompare.exe`, `BComp.exe`, or `BComp.com left right` for report-only work; those can open GUI windows.

## Default Path

Prefer the bundled helper:

```powershell
& ".\scripts\New-BeyondCompareReport.ps1" -Left "C:\left.docx" -Right "C:\right.docx" -Output "C:\report.txt"
```

Modes:

| Mode | Use | Output |
|---|---|---|
| `File` | Default file/document/spreadsheet report | Text report |
| `TextHtml` | Text side-by-side HTML report | HTML report |
| `Folder` | Recursive folder comparison | XML report |

Examples:

```powershell
& ".\scripts\New-BeyondCompareReport.ps1" -Mode TextHtml -Left "C:\a.txt" -Right "C:\b.txt" -Output "C:\diff.html"
& ".\scripts\New-BeyondCompareReport.ps1" -Mode Folder -Left "C:\old" -Right "C:\new" -Output "C:\folder.xml"
```

After Office reports, check for conversion failure:

```powershell
(Get-Content "C:\report.txt" -Raw) -match "Conversion Error"
```

If true, say Beyond Compare could not convert one or both files; do not claim a meaningful content diff.

## Quick Exit Code

For two files only:

```powershell
& $BComp /qc=binary "C:\left.file" "C:\right.file"
$LASTEXITCODE
```

Key codes: `1` binary same, `2` rules-based same but not byte-identical, `11` binary different, `12` similar, `13` rules-based different, `100` error, `105` script load error, `106` script syntax error, `107` file/folder load failure.

Avoid `/qc` for folders; generate folder XML instead. Avoid `/qc=size`; observed behavior was not reliable as size-only comparison.

## Reference

Read `references/beyond-compare-5-windows-cli-report-guide.md` only when needing exact script syntax, observed exit-code nuance, or troubleshooting details.

## Safety

Never run mirror sync, delete, move, or direct GUI compare commands unless the user explicitly asks for destructive or interactive behavior.
