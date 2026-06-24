# Beyond Compare 5 Windows CLI Report Notes

Use this reference only when SKILL.md is not enough.

## Verified Environment

| Item | Value |
|---|---|
| Beyond Compare version tested | `5.0.0.29773` |
| Console helper paths | Explicit `-BCompPath` folder/file, `J:\Program Files\Beyond Compare 5\BComp.com`, `D:\Program Files\Beyond Compare 5\BComp.com` |
| Safe report pattern | PowerShell running `BComp.com /silent /closescript @"script.bc" left right report` |

Check only the fixed helper paths above unless the user asks to locate the install. Avoid `BCompare.exe`, `BComp.exe`, and direct `BComp.com left right` for automation because they can open GUI windows.

## Report Scripts

Mismatch-only file report used before JSONL wrapping:

```text
file-report layout:side-by-side options:display-mismatches,line-numbers output-to:"%3" "%1" "%2"
```

Human HTML text report:

```text
text-report layout:side-by-side options:display-mismatches,line-numbers output-to:"%3" output-options:html-color "%1" "%2"
```

Folder XML report:

```text
criteria rules-based
load "%1" "%2"
expand all
folder-report layout:xml options:display-mismatches output-to:"%3"
```

Older fallback folder XML report:

```text
load "%1" "%2"
expand all
folder-report layout:xml output-to:"%3"
```

Important: `output-to:"report.html"` alone writes plain text. Use `output-options:html-color` for real HTML.

## Exit Codes

| Code | Meaning |
|---:|---|
| `0` | Script/command succeeded |
| `1` | Binary same |
| `2` | Rules-based same but not binary-identical |
| `11` | Binary differences |
| `12` | Similar |
| `13` | Rules-based differences |
| `14` | Conflicts detected |
| `100` | Error |
| `105` | Error loading script file |
| `106` | Script syntax error |
| `107` | Script failed to load files/folders |

Observed nuance:

| Case | Exit |
|---|---:|
| Identical text with `/qc=rules-based` | `1` |
| LF vs CRLF with `/qc=rules-based` | `2` |
| Whitespace-only text change | `12` |
| Different text | `13` |
| `/qc` on folders | `100` |

Do not rely on `/qc=size` as size-only comparison; same-size different files returned `11` in testing.

## Office Reports

Use `file-report` for `.doc`, `.docx`, `.xls`, and `.xlsx`, then inspect the generated report.

| Report text | Meaning |
|---|---|
| `Text Compare` | Text/extracted-text comparison ran |
| `Table Compare` | Table comparison ran |
| `Conversion Error` | BC could not convert one or both files |

If `Conversion Error` appears, report that content comparison failed. Binary `/qc=binary` can still prove bytes differ, but it does not prove a meaningful document/spreadsheet diff.

## Destructive Operations

Mirror sync is headless but destructive:

```text
log normal "%3"
load "%1" "%2"
sync create-empty mirror:left->right
```

Do not use sync/delete/move scripts unless the user explicitly asks for destructive folder modification.
