param(
    [ValidateSet("File", "Folder", "TextHtml")]
    [string] $Mode = "File",

    [Parameter(Mandatory = $true)]
    [string] $Left,

    [Parameter(Mandatory = $true)]
    [string] $Right,

    [Parameter(Mandatory = $true)]
    [string] $Output,

    [switch] $FailOnConversionError
)

$ErrorActionPreference = "Stop"

$BComp = "J:\Program Files\Beyond Compare 5\BComp.com"

if (-not (Test-Path -LiteralPath $BComp)) {
    throw "Beyond Compare console helper not found at hardcoded path: $BComp"
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

$scriptPath = [System.IO.Path]::ChangeExtension([System.IO.Path]::GetTempFileName(), ".bc")

try {
    switch ($Mode) {
        "Folder" {
            @'
load "%1" "%2"
expand all
folder-report layout:xml output-to:"%3"
'@ | Set-Content -LiteralPath $scriptPath -Encoding ASCII
        }
        "TextHtml" {
            @'
text-report layout:side-by-side options:display-mismatches,line-numbers output-to:"%3" output-options:html-color "%1" "%2"
'@ | Set-Content -LiteralPath $scriptPath -Encoding ASCII
        }
        default {
            @'
file-report layout:summary output-to:"%3" "%1" "%2"
'@ | Set-Content -LiteralPath $scriptPath -Encoding ASCII
        }
    }

    & $BComp /silent /closescript "@$scriptPath" $Left $Right $Output | Out-Null
    $exitCode = $LASTEXITCODE

    if ($exitCode -ne 0) {
        exit $exitCode
    }

    if (-not (Test-Path -LiteralPath $Output)) {
        throw "Beyond Compare exited successfully but did not create report: $Output"
    }

    if ($FailOnConversionError) {
        $content = Get-Content -LiteralPath $Output -Raw
        if ($content -match "Conversion Error") {
            throw "Beyond Compare report contains Conversion Error."
        }
    }

    Write-Output $Output
    exit 0
}
finally {
    if (Test-Path -LiteralPath $scriptPath) {
        Remove-Item -LiteralPath $scriptPath -Force
    }
}
