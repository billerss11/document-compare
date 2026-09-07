param(
    [ValidateSet("File", "Folder", "TextHtml")]
    [string] $Mode = "File",

    [ValidateSet("Auto", "Jsonl", "Raw")]
    [string] $OutputFormat = "Auto",

    [Parameter(Mandatory = $true)]
    [string] $Left,

    [Parameter(Mandatory = $true)]
    [string] $Right,

    [Parameter(Mandatory = $true)]
    [string] $Output,

    [string] $BCompPath,

    [switch] $FailOnConversionError
)

$ErrorActionPreference = "Stop"

function Resolve-BCompCandidate {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    if (Test-Path -LiteralPath $Path -PathType Container) {
        $candidate = Join-Path $Path "BComp.com"
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return $candidate
        }
    }

    if (Test-Path -LiteralPath $Path -PathType Leaf) {
        return $Path
    }

    return $null
}

function New-TempReportPath {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Extension
    )

    Join-Path ([System.IO.Path]::GetTempPath()) ("document-compare-{0}{1}" -f [guid]::NewGuid().ToString("N"), $Extension)
}

$BCompCandidates = @()
if ($BCompPath) {
    $BCompCandidates += $BCompPath
}
$BCompCandidates += @(
    "J:\Program Files\Beyond Compare 5\BComp.com",
    "D:\Program Files\Beyond Compare 5\BComp.com"
)
$BComp = $null
foreach ($candidate in $BCompCandidates) {
    $resolved = Resolve-BCompCandidate -Path $candidate
    if ($resolved) {
        $BComp = $resolved
        break
    }
}

function New-JsonLine {
    param(
        [Parameter(Mandatory = $true)]
        [object] $InputObject
    )

    $InputObject | ConvertTo-Json -Compress -Depth 12
}

function Write-JsonLines {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IEnumerable] $Records,

        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    $writer = New-Object System.IO.StreamWriter($Path, $false, $utf8NoBom)
    try {
        foreach ($record in $Records) {
            $writer.WriteLine((New-JsonLine -InputObject $record))
        }
    }
    finally {
        $writer.Dispose()
    }
}

function New-MetadataRecord {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Mode,

        [Parameter(Mandatory = $true)]
        [string] $Left,

        [Parameter(Mandatory = $true)]
        [string] $Right,

        [Parameter(Mandatory = $true)]
        [string] $RawFormat
    )

    [ordered]@{
        type = "metadata"
        schema_version = 2
        mode = $Mode
        left = (Resolve-Path -LiteralPath $Left).Path
        right = (Resolve-Path -LiteralPath $Right).Path
        engine = "Beyond Compare 5"
        raw_format = $RawFormat
    }
}

function Get-AttributeValue {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary] $Attributes,

        [Parameter(Mandatory = $true)]
        [string[]] $Names
    )

    foreach ($name in $Names) {
        foreach ($key in $Attributes.Keys) {
            if ($key -ieq $name) {
                return $Attributes[$key]
            }
        }
    }

    return $null
}

function Get-ElementAttributes {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlElement] $Element
    )

    $attributes = [ordered]@{}
    foreach ($attribute in $Element.Attributes) {
        $attributes[$attribute.Name] = $attribute.Value
    }
    $attributes
}

function Get-ChildText {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlElement] $Element,

        [Parameter(Mandatory = $true)]
        [string[]] $Paths
    )

    foreach ($path in $Paths) {
        $node = $Element.SelectSingleNode($path)
        if ($node -and -not [string]::IsNullOrWhiteSpace($node.InnerText)) {
            return $node.InnerText
        }
    }

    return $null
}

function Get-SideInfo {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlElement] $Element,

        [Parameter(Mandatory = $true)]
        [ValidateSet("lt", "rt")]
        [string] $Side
    )

    $node = $Element.SelectSingleNode($Side)
    if (-not $node) {
        return $null
    }

    $info = [ordered]@{}
    foreach ($childName in @("name", "size", "modified")) {
        $child = $node.SelectSingleNode($childName)
        if ($child -and -not [string]::IsNullOrWhiteSpace($child.InnerText)) {
            $info[$childName] = $child.InnerText
        }
    }

    if ($info.Count -eq 0) {
        return $null
    }

    $info
}

function Join-ReportPath {
    param(
        [string] $Parent,
        [string] $Name
    )

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return $Parent
    }

    if ([string]::IsNullOrWhiteSpace($Parent)) {
        return $Name
    }

    $trimChars = [char[]] @("/", "\")
    "{0}/{1}" -f $Parent.TrimEnd($trimChars), $Name.TrimStart($trimChars)
}

function Add-FolderXmlRecords {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlElement] $Element,

        [string] $ParentPath,

        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.List[object]] $Records
    )

    $attributes = Get-ElementAttributes -Element $Element
    $path = Get-AttributeValue -Attributes $attributes -Names @("path", "relativePath", "relpath")

    if (-not $path) {
        $name = Get-AttributeValue -Attributes $attributes -Names @("name", "filename", "file", "folder")
        if (-not $name) {
            $name = Get-ChildText -Element $Element -Paths @("lt/name", "rt/name", "name")
        }
        $path = Join-ReportPath -Parent $ParentPath -Name $name
    }

    $status = Get-AttributeValue -Attributes $attributes -Names @("status", "state", "comparison", "result", "side")
    $leftInfo = Get-SideInfo -Element $Element -Side "lt"
    $rightInfo = Get-SideInfo -Element $Element -Side "rt"
    # BC does not put a status attribute on one-sided folders, including empty ones.
    if ($Element.LocalName -eq "foldercomp" -and -not $status) {
        if ($leftInfo -and -not $rightInfo) { $status = "ltonly" }
        if ($rightInfo -and -not $leftInfo) { $status = "rtonly" }
    }
    $isRoot = $Element.OwnerDocument.DocumentElement -eq $Element
    if (-not $isRoot -and ($attributes.Count -gt 0 -or $status) -and $status -ine "same") {
        $record = [ordered]@{
            type = "entry"
            element = $Element.LocalName
        }

        if ($path) {
            $record["path"] = $path
        }

        if ($status) {
            $record["status"] = $status
        }

        if ($leftInfo) {
            $record["left_item"] = $leftInfo
        }

        if ($rightInfo) {
            $record["right_item"] = $rightInfo
        }

        $record["details"] = $attributes
        $Records.Add($record)
    }

    foreach ($child in $Element.ChildNodes) {
        if ($child.NodeType -eq [System.Xml.XmlNodeType]::Element) {
            Add-FolderXmlRecords -Element ([System.Xml.XmlElement] $child) -ParentPath $path -Records $Records
        }
    }
}

function Test-ReportConversionError {
    param([string] $Content)

    # Match BC diagnostic headers, not the numbered file content or HTML cells.
    $Content -match '(?m)^(?:Left|Right) error:[ \t]*Conversion Error\b'
}

function Convert-TextReportToJsonl {
    param(
        [Parameter(Mandatory = $true)]
        [string] $RawPath,

        [Parameter(Mandatory = $true)]
        [string] $OutputPath,

        [Parameter(Mandatory = $true)]
        [string] $Mode,

        [Parameter(Mandatory = $true)]
        [string] $Left,

        [Parameter(Mandatory = $true)]
        [string] $Right,

        [Parameter(Mandatory = $true)]
        [string] $BComp
    )

    $content = Get-Content -LiteralPath $RawPath -Raw
    $conversionError = Test-ReportConversionError -Content $content
    $records = New-Object System.Collections.Generic.List[object]

    $records.Add((New-MetadataRecord -Mode $Mode -Left $Left -Right $Right -RawFormat "text"))
    $summary = [ordered]@{
        type = "summary"
        outcome = "error"
        comparison = "rules-based"
    }
    $records.Add($summary)

    $reportErrors = [regex]::Matches($content, '(?im)^(?:Left|Right) error:[^\r\n]*')
    if ($reportErrors.Count -gt 0) {
        $records.Add([ordered]@{
            type = $(if ($conversionError) { "conversion_error" } else { "comparison_error" })
            message = ($reportErrors.Value -join "; ")
        })
    }
    else {
        # Ask the engine for the outcome; do not infer equality from report layout.
        & $BComp /silent /qc=rules-based $Left $Right | Out-Null
        $quickExitCode = $LASTEXITCODE
        $summary["quick_compare_exit_code"] = $quickExitCode
        if ($quickExitCode -in @(1, 2)) {
            $summary["outcome"] = "same"
        }
        elseif ($quickExitCode -in @(11, 12, 13, 14)) {
            $summary["outcome"] = "different"
        }
        else {
            $records.Add([ordered]@{
                type = "comparison_error"
                message = "Beyond Compare quick comparison failed or returned an unrecognized exit code: $quickExitCode"
            })
        }
    }

    $lineNumber = 0
    foreach ($line in ($content -split "`r?`n")) {
        $lineNumber++
        $text = $line.TrimEnd()
        if ($text.Length -eq 0) {
            continue
        }

        $records.Add([ordered]@{
            type = "report_line"
            line = $lineNumber
            text = $text
        })
    }

    Write-JsonLines -Records $records -Path $OutputPath
    $conversionError
}

function Convert-FolderReportToJsonl {
    param(
        [Parameter(Mandatory = $true)]
        [string] $RawPath,

        [Parameter(Mandatory = $true)]
        [string] $OutputPath,

        [Parameter(Mandatory = $true)]
        [string] $Left,

        [Parameter(Mandatory = $true)]
        [string] $Right
    )

    $records = New-Object System.Collections.Generic.List[object]
    $records.Add((New-MetadataRecord -Mode "Folder" -Left $Left -Right $Right -RawFormat "xml"))
    $summary = [ordered]@{
        type = "summary"
        outcome = "error"
        comparison = "rules-based"
    }
    $records.Add($summary)

    try {
        [xml] $xml = Get-Content -LiteralPath $RawPath -Raw
        if (-not $xml.SelectSingleNode('/bcreport/foldercomp')) {
            throw "Expected a Beyond Compare folder report (bcreport/foldercomp)."
        }
        Add-FolderXmlRecords -Element $xml.DocumentElement -ParentPath "" -Records $records

        $entries = @($records | Where-Object { $_.type -eq "entry" })
        $unknownEntries = @($entries | Where-Object { $_.status -notin @("diff", "ltonly", "rtonly") })
        $summary["entry_count"] = $entries.Count
        if ($unknownEntries.Count -gt 0 -or $xml.SelectNodes('//error').Count -gt 0) {
            $records.Add([ordered]@{
                type = "comparison_error"
                message = "Folder report contains errors or unrecognized entry statuses; inspect the raw XML before drawing conclusions."
            })
        }
        else {
            $summary["outcome"] = $(if ($entries.Count -gt 0) { "different" } else { "same" })
        }
    }
    catch {
        $records.Add([ordered]@{
            type = "xml_parse_error"
            message = $_.Exception.Message
        })

        $lineNumber = 0
        foreach ($line in (Get-Content -LiteralPath $RawPath)) {
            $lineNumber++
            if ($line.Trim().Length -eq 0) {
                continue
            }

            $records.Add([ordered]@{
                type = "raw_line"
                line = $lineNumber
                text = $line
            })
        }
    }

    Write-JsonLines -Records $records -Path $OutputPath
    $false
}

if (-not $BComp) {
    throw "Beyond Compare console helper not found. Checked: $($BCompCandidates -join '; ')"
}

if (-not (Test-Path -LiteralPath $Left)) {
    throw "Left path not found: $Left"
}

if (-not (Test-Path -LiteralPath $Right)) {
    throw "Right path not found: $Right"
}

$outputDir = Split-Path -Parent $Output
if ($outputDir -and -not (Test-Path -LiteralPath $outputDir)) {
    New-Item -ItemType Directory -Force -Path $outputDir | Out-Null
}

$effectiveOutputFormat = $OutputFormat
if ($effectiveOutputFormat -eq "Auto") {
    if ($Mode -eq "TextHtml") {
        $effectiveOutputFormat = "Raw"
    }
    else {
        $effectiveOutputFormat = "Jsonl"
    }
}

if ($Mode -eq "TextHtml" -and $effectiveOutputFormat -eq "Jsonl") {
    throw "TextHtml is a human HTML report mode. Use -OutputFormat Raw, or use -Mode File for JSONL."
}

$scriptPath = New-TempReportPath -Extension ".bc"
$rawOutputPath = $Output

if ($effectiveOutputFormat -eq "Jsonl") {
    $rawExtension = if ($Mode -eq "Folder") { ".xml" } else { ".txt" }
    $rawOutputPath = New-TempReportPath -Extension $rawExtension
}

try {
    switch ($Mode) {
        "Folder" {
            @'
criteria rules-based
load "%1" "%2"
expand all
folder-report layout:xml options:display-mismatches output-to:"%3"
'@ | Set-Content -LiteralPath $scriptPath -Encoding ASCII
        }
        "TextHtml" {
            @'
text-report layout:side-by-side options:display-mismatches,line-numbers output-to:"%3" output-options:html-color "%1" "%2"
'@ | Set-Content -LiteralPath $scriptPath -Encoding ASCII
        }
        default {
            @'
file-report layout:side-by-side options:display-mismatches,line-numbers output-to:"%3" "%1" "%2"
'@ | Set-Content -LiteralPath $scriptPath -Encoding ASCII
        }
    }

    & $BComp /silent /closescript "@$scriptPath" $Left $Right $rawOutputPath | Out-Null
    $exitCode = $LASTEXITCODE

    if ($exitCode -ne 0) {
        exit $exitCode
    }

    if (-not (Test-Path -LiteralPath $rawOutputPath)) {
        throw "Beyond Compare exited successfully but did not create report: $rawOutputPath"
    }

    $conversionError = $false
    if ($effectiveOutputFormat -eq "Jsonl") {
        if ($Mode -eq "Folder") {
            $conversionError = Convert-FolderReportToJsonl -RawPath $rawOutputPath -OutputPath $Output -Left $Left -Right $Right
        }
        else {
            $conversionError = Convert-TextReportToJsonl -RawPath $rawOutputPath -OutputPath $Output -Mode $Mode -Left $Left -Right $Right -BComp $BComp
        }
    }
    else {
        $content = Get-Content -LiteralPath $Output -Raw
        $conversionError = Test-ReportConversionError -Content $content
    }

    if ($FailOnConversionError -and $conversionError) {
        throw "Beyond Compare report contains Conversion Error."
    }

    if (-not (Test-Path -LiteralPath $Output)) {
        throw "Report was not created: $Output"
    }

    Write-Output $Output
    exit 0
}
finally {
    if (Test-Path -LiteralPath $scriptPath) {
        Remove-Item -LiteralPath $scriptPath -Force
    }
    if ($effectiveOutputFormat -eq "Jsonl" -and (Test-Path -LiteralPath $rawOutputPath)) {
        Remove-Item -LiteralPath $rawOutputPath -Force
    }
}
