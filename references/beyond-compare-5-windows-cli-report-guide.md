# Beyond Compare 5 Windows CLI Report Guide

This guide is for Windows automation with Beyond Compare 5 installed at:

```bat
J:\Program Files\Beyond Compare 5
```

The main goal is to generate comparison results and reports without opening the GUI.

Use:

```bat
"J:\Program Files\Beyond Compare 5\BComp.com"
```

Do not use `BCompare.exe` or `BComp.exe` for report-only automation unless you intentionally want GUI behavior.

## Key Rules

| Rule | Recommendation |
|---|---|
| Headless reports | Use `BComp.com /silent /closescript @"script.bc" ...` |
| Exit-code-only file checks | Use `BComp.com /qc=binary left right` |
| Folder comparison | Generate `folder-report layout:xml` and parse the XML |
| HTML reports | Add `output-options:html-color`; `.html` filename alone is not enough |
| GUI compare windows | Avoid direct commands like `BComp.com left right` in automation |
| Office files | Reports can work, but check output for `Conversion Error` |

## Executables

| Command | Use | Automation verdict |
|---|---|---|
| `BComp.com` | Windows console helper | Recommended for scripts, reports, and exit codes |
| `BComp.exe` | Windows GUI helper | Avoid for report-only automation |
| `BCompare.exe` | Main GUI application | Avoid for report-only automation |

## Safe Headless Patterns

### Generate a Text File Report

`file_report_text.bc`:

```text
file-report layout:summary output-to:"%3" "%1" "%2"
```

Run:

```bat
"J:\Program Files\Beyond Compare 5\BComp.com" /silent /closescript @"file_report_text.bc" "C:\left.txt" "C:\right.txt" "C:\report.txt"
```

Expected result:

| Result | Meaning |
|---|---|
| Exit code `0` | Script ran successfully |
| Report file exists | Comparison report was generated |
| Report content | Must be inspected to determine the actual differences |

### Generate an HTML Text Compare Report

Important: `output-to:"report.html"` does not automatically create HTML. Add `output-options:html-color`.

`text_report_html.bc`:

```text
text-report layout:side-by-side options:display-mismatches,line-numbers output-to:"%3" output-options:html-color "%1" "%2"
```

Run:

```bat
"J:\Program Files\Beyond Compare 5\BComp.com" /silent /closescript @"text_report_html.bc" "C:\left.txt" "C:\right.txt" "C:\report.html"
```

Expected result:

| Result | Meaning |
|---|---|
| Exit code `0` | Script ran successfully |
| Report starts with `<!DOCTYPE HTML>` | Real HTML report was generated |

### Generate a Folder XML Report

`folder_report_xml.bc`:

```text
load "%1" "%2"
expand all
folder-report layout:xml output-to:"%3"
```

Run:

```bat
"J:\Program Files\Beyond Compare 5\BComp.com" /silent /closescript @"folder_report_xml.bc" "C:\left_folder" "C:\right_folder" "C:\folder_report.xml"
```

Expected result:

| Result | Meaning |
|---|---|
| Exit code `0` | Script ran successfully |
| XML root `bcreport` | Folder report was generated |
| Parse XML | Use the XML content to determine equal, different, left-only, and right-only files |

## Quick File Compare With Exit Codes

Use quick compare only for two files, not folders.

```bat
"J:\Program Files\Beyond Compare 5\BComp.com" /qc=binary "C:\left.bin" "C:\right.bin"
```

Then read `%ERRORLEVEL%`.

Example:

```bat
"J:\Program Files\Beyond Compare 5\BComp.com" /qc=binary "C:\left.bin" "C:\right.bin"
echo %ERRORLEVEL%
```

## Quick Compare Types

| Command | Meaning | Recommended? |
|---|---|---|
| `/qc=binary` | Byte-by-byte comparison | Yes |
| `/qc=crc` | CRC comparison | Yes |
| `/quickcompare=crc` | Long-form CRC alias | Yes |
| `/qc=rules-based` | Uses file format rules | Yes, but understand the exit codes |
| `/qc` | Default quick compare | Acceptable, but be explicit when possible |
| `/qc=size` | Intended as size comparison | Do not rely on it; test showed different same-size files returned different |

If you need size-only comparison, use PowerShell file sizes instead:

```powershell
(Get-Item "C:\left.dat").Length -eq (Get-Item "C:\right.dat").Length
```

## Exit Codes

Observed on Beyond Compare `5.0.0.29773`.

| Code | Meaning |
|---:|---|
| `0` | Command or script completed successfully |
| `1` | Binary same |
| `2` | Rules-based same, but not binary-identical |
| `11` | Binary differences |
| `12` | Similar |
| `13` | Rules-based differences |
| `14` | Conflicts detected |
| `100` | Error |
| `101` | Conflicts detected; merge output not saved |
| `102+` | Other errors |
| `105` | Error loading script file |
| `106` | Script syntax error |
| `107` | Script failed to load folders or files |

Important nuance:

| Scenario | Result |
|---|---:|
| Identical text files with `/qc=rules-based` | `1` |
| Same text with LF vs CRLF using `/qc=rules-based` | `2` |
| Whitespace-only text change using `/qc=rules-based` | `12` |
| Different text using `/qc=rules-based` | `13` |

## Office File Reports

Beyond Compare can generate reports for document and spreadsheet files, but Office conversion can fail depending on the file format and installed converters.

Use generic `file-report`:

`office_file_report.bc`:

```text
file-report layout:summary output-to:"%3" "%1" "%2"
```

Run:

```bat
"J:\Program Files\Beyond Compare 5\BComp.com" /silent /closescript @"office_file_report.bc" "C:\left.docx" "C:\right.docx" "C:\docx_report.txt"
```

Check the report content.

| Report content | Meaning |
|---|---|
| `Text Compare` | BC compared the file as text/extracted text |
| `Table Compare` | BC compared the file as tabular data |
| `Conversion Error` | BC could not convert one or both files for content comparison |

Recommended handling:

```powershell
$report = Get-Content "C:\report.txt" -Raw
if ($report -match "Conversion Error") {
  Write-Error "Beyond Compare could not convert one or both files."
}
```

## Folder Reports

Do not use `/qc` for folders.

Bad:

```bat
"J:\Program Files\Beyond Compare 5\BComp.com" /qc "C:\left_folder" "C:\right_folder"
```

Use this instead:

```bat
"J:\Program Files\Beyond Compare 5\BComp.com" /silent /closescript @"folder_report_xml.bc" "C:\left_folder" "C:\right_folder" "C:\folder_report.xml"
```

Then parse `folder_report.xml`.

## Mirror Sync

This is headless, but it is destructive. It can delete right-side files.

Only use it when you really want the right folder to match the left folder.

`mirror_left_to_right.bc`:

```text
log normal "%3"
load "%1" "%2"
sync create-empty mirror:left->right
```

Run:

```bat
"J:\Program Files\Beyond Compare 5\BComp.com" /silent /closescript @"mirror_left_to_right.bc" "C:\source" "D:\target" "C:\sync.log"
```

Result:

| Effect | Description |
|---|---|
| Copies/updates files | Right side receives left-side files |
| Deletes orphans | Right-side files missing from left are deleted |
| Writes log | Log is written to the path passed as `%3` |

## Useful Switches For Report Automation

| Switch | Use |
|---|---|
| `/silent` | Run script without showing a window |
| `/closescript` | Close script window after completion |
| `/qc=binary` | Headless two-file binary comparison |
| `/qc=crc` | Headless two-file CRC comparison |
| `/ro` | Read-only mode for interactive views; not usually needed for reports |
| `/title1=...`, `/title2=...` | Label sides in interactive/file views |
| `/vcs1=...`, `/vcs2=...` | Show VCS paths and help file-format selection |

## Commands To Avoid In Report-Only Automation

These are valid Beyond Compare commands, but they are not the right tool for headless report generation.

| Command pattern | Why to avoid |
|---|---|
| `BComp.com "C:\left.txt" "C:\right.txt"` | Opens a compare window |
| `BComp.com "C:\left_folder" "C:\right_folder"` | Opens a folder compare window |
| `BCompare.exe ...` | Main GUI app |
| `BComp.exe ...` | GUI helper |
| `BCompare.exe "patch.diff"` | Opens patch view |
| `dir | BCompare.exe -` | Opens stdin content in GUI |
| `BCompare.exe "settings.bcpkg"` | Imports settings package |
| Saved session/workspace commands | Usually open GUI |

## Recommended Agent-Safe Recipes

### Need yes/no same/different for two files

```bat
"J:\Program Files\Beyond Compare 5\BComp.com" /qc=binary "C:\left.file" "C:\right.file"
```

Use `%ERRORLEVEL%`.

### Need a human-readable file report

```bat
"J:\Program Files\Beyond Compare 5\BComp.com" /silent /closescript @"file_report_text.bc" "C:\left.file" "C:\right.file" "C:\report.txt"
```

Check the report file.

### Need an HTML text report

```bat
"J:\Program Files\Beyond Compare 5\BComp.com" /silent /closescript @"text_report_html.bc" "C:\left.txt" "C:\right.txt" "C:\report.html"
```

The script must include:

```text
output-options:html-color
```

### Need a folder comparison result

```bat
"J:\Program Files\Beyond Compare 5\BComp.com" /silent /closescript @"folder_report_xml.bc" "C:\left_folder" "C:\right_folder" "C:\folder_report.xml"
```

Parse the XML.

## Validation Summary

These checks were performed headlessly against Beyond Compare `5.0.0.29773`.

| Area | Result |
|---|---|
| `BComp.com /help` | Valid |
| `/qc=binary` | Valid |
| `/qc=crc` and `/quickcompare=crc` | Valid |
| `/qc=rules-based` | Valid, with exit-code nuance |
| `/qc=size` | Not reliable based on observed behavior |
| `/qc` on folders | Not valid for folder comparison; returned error |
| Silent folder XML report | Valid |
| Silent text report | Valid |
| HTML report with `output-options:html-color` | Valid |
| HTML filename without HTML output option | Misleading; creates plain text |
| Script error codes `105`, `106`, `107` | Valid |
| Office file reports | Possible, but report must be checked for conversion errors |

## Bottom Line

For Windows report automation, use this pattern:

```bat
"J:\Program Files\Beyond Compare 5\BComp.com" /silent /closescript @"script.bc" "left" "right" "report"
```

Use direct `BComp.com left right` only when you intentionally want the GUI.

