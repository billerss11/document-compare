---
name: document-compare
description: "MUST USE for file or folder comparison tasks on Windows: compare, diff, find differences, check what changed, compare two versions, same/different checks, or create comparison reports for documents, spreadsheets, text files, binary files, or folders. Handles .doc, .docx, .xls, .xlsx, .txt, folder trees, XML folder reports, and explicit HTML side-by-side report requests with headless Beyond Compare 5, defaults to machine-friendly JSONL, then summarizes results in plain language without opening the GUI. Do not use for purely conceptual comparisons where no files or folders are being compared."
---

# Document Compare

Prefer PowerShell for this skill.

## When To Use This Skill

Use this skill whenever the user asks to compare actual files, folders, exports, versions, revisions, or before/after copies. This includes casual wording like "what changed?", "show differences", "old vs new", "v1 vs v2", "are these the same?", "compare these docs", "diff these spreadsheets", or "make me a comparison report".

Do not skip this skill just because the user did not say "Beyond Compare". The point of the skill is to choose the correct headless Beyond Compare workflow for them.

Do not use this skill for comparing ideas, APIs, products, plans, code approaches, or other conceptual topics unless the task includes concrete files or folders to compare.

Use the bundled helper unless the user only needs a quick same/different exit code:

```powershell
& ".\scripts\New-BeyondCompareReport.ps1" -Left "C:\left.docx" -Right "C:\right.docx" -Output "C:\diff.jsonl"
```

Default to machine-friendly JSONL for `File` and `Folder` modes. Read it, check for `conversion_error`, then summarize the result in plain language. Do not create HTML or expose raw JSONL unless the user asks for it.

The helper checks only fixed Beyond Compare console paths, in order: explicit `-BCompPath` (folder or `BComp.com`), `J:\Program Files\Beyond Compare 5\BComp.com`, then `D:\Program Files\Beyond Compare 5\BComp.com`. Do not search the filesystem unless all fixed paths fail and the user asks you to locate it.

## Modes

| Mode | Use | Default output |
|---|---|---|
| `File` | Documents, spreadsheets, text, binary-like files | JSONL wrapping a mismatch-only text report |
| `Folder` | Recursive folder comparison | JSONL converted from Beyond Compare XML |
| `TextHtml` | Only when the user asks for HTML or a visual side-by-side report | HTML |

Use explicit raw/HTML output only when requested or needed for troubleshooting:

```powershell
& ".\scripts\New-BeyondCompareReport.ps1" -Mode TextHtml -OutputFormat Raw -Left "C:\a.txt" -Right "C:\b.txt" -Output "C:\diff.html"
& ".\scripts\New-BeyondCompareReport.ps1" -Mode Folder -OutputFormat Raw -Left "C:\old" -Right "C:\new" -Output "C:\folder.xml"
```

## Excel Sheets

Beyond Compare CLI cannot select Excel worksheets by sheet name. For multisheet `.xls`/`.xlsx` or sheet-name requests, first extract requested/all worksheets to normalized UTF-8 TSV files, then compare those exports. Default to matching sheets by name, reporting added/removed sheets, and comparing cell values only. Compare formulas, formatting, charts, macros, pivots, comments, or structure only when explicitly asked.

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

Source/update repo: https://github.com/billerss11/document-compare

Use only the fixed console helper paths above. Do not use `BCompare.exe`, `BComp.exe`, or direct `BComp.com left right` for reports; those can open GUI windows.

Read `references/beyond-compare-5-windows-cli-report-guide.md` only for exact Beyond Compare script syntax, exit-code nuance, or troubleshooting.

## Safety

Never run mirror sync, delete, move, or direct GUI compare commands unless the user explicitly asks for destructive or interactive behavior.
