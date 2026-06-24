---
name: document-compare
description: Generate headless Beyond Compare 5 reports on Windows with LLM-friendly JSONL by default. Use for comparing documents, spreadsheets, text files, binary files, or folders, especially .doc, .docx, .xls, .xlsx, .txt, folder trees, HTML reports for humans, XML folder reports, or same/different checks without opening the GUI.
---

# Document Compare

Prefer PowerShell for this skill.

Use the bundled helper unless the user only needs a quick same/different exit code:

```powershell
& ".\scripts\New-BeyondCompareReport.ps1" -Left "C:\left.docx" -Right "C:\right.docx" -Output "C:\diff.jsonl"
```

Default output is LLM-readable JSON Lines for `File` and `Folder` modes. Read the JSONL, check for `conversion_error`, then summarize changes for the human.

The helper checks only fixed Beyond Compare console paths, in order: explicit `-BCompPath` (folder or `BComp.com`), `J:\Program Files\Beyond Compare 5\BComp.com`, then `D:\Program Files\Beyond Compare 5\BComp.com`. Do not search the filesystem unless all fixed paths fail and the user asks you to locate it.

## Modes

| Mode | Use | Default output |
|---|---|---|
| `File` | Documents, spreadsheets, text, binary-like files | JSONL wrapping a mismatch-only text report |
| `Folder` | Recursive folder comparison | JSONL converted from Beyond Compare XML |
| `TextHtml` | Human side-by-side text report | HTML |

Use explicit raw output only when needed:

```powershell
& ".\scripts\New-BeyondCompareReport.ps1" -Mode TextHtml -OutputFormat Raw -Left "C:\a.txt" -Right "C:\b.txt" -Output "C:\diff.html"
& ".\scripts\New-BeyondCompareReport.ps1" -Mode Folder -OutputFormat Raw -Left "C:\old" -Right "C:\new" -Output "C:\folder.xml"
```

## JSONL Contract

Expect one JSON object per line:

- `metadata`: inputs, mode, engine, schema version
- `entry`: folder item from XML, with `path`, optional `status`, and raw `details`
- `report_line`: numbered line from a file/document report
- `conversion_error`: Beyond Compare could not convert one or both Office files
- `xml_parse_error`: folder XML could not be parsed; raw lines follow

If `conversion_error` appears, say content comparison failed. Do not claim a meaningful document/spreadsheet diff.

## Quick Exit Code

For two files only:

```powershell
$BComp = @(
  "J:\Program Files\Beyond Compare 5\BComp.com",
  "D:\Program Files\Beyond Compare 5\BComp.com"
) | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if (-not $BComp) { throw "Beyond Compare console helper not found." }
& $BComp /qc=binary "C:\left.file" "C:\right.file"
$LASTEXITCODE
```

Key codes: `1` binary same, `2` rules-based same but not byte-identical, `11` binary different, `12` similar, `13` rules-based different, `100` error, `105` script load error, `106` script syntax error, `107` file/folder load failure.

Avoid `/qc` for folders; generate folder JSONL instead. Avoid `/qc=size`; observed behavior was not reliable as size-only comparison.

## Reference

Use only the fixed console helper paths above. Do not use `BCompare.exe`, `BComp.exe`, or direct `BComp.com left right` for reports; those can open GUI windows.

Read `references/beyond-compare-5-windows-cli-report-guide.md` only for exact Beyond Compare script syntax, exit-code nuance, or troubleshooting.

## Safety

Never run mirror sync, delete, move, or direct GUI compare commands unless the user explicitly asks for destructive or interactive behavior.
