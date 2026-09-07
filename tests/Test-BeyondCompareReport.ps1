param([string] $BCompPath)

$ErrorActionPreference = "Stop"
$helper = Join-Path $PSScriptRoot "..\scripts\New-BeyondCompareReport.ps1"
$shell = (Get-Process -Id $PID).Path
$tempBase = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath()).TrimEnd('\')
$testRoot = Join-Path $tempBase ("document-compare-test-" + [guid]::NewGuid().ToString("N"))
$utf8 = New-Object System.Text.UTF8Encoding($false)
New-Item -ItemType Directory -Path $testRoot | Out-Null

function Write-Fixture {
    param([string] $Name, [string] $Text)
    $path = Join-Path $testRoot $Name
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $path) | Out-Null
    [System.IO.File]::WriteAllText($path, $Text, $utf8)
}

function Assert-True {
    param([bool] $Condition, [string] $Message)
    if (-not $Condition) { throw $Message }
}

function Invoke-Comparison {
    param(
        [string] $LeftName,
        [string] $RightName,
        [string] $Mode = "File",
        [string] $Format = "Jsonl",
        [switch] $Strict,
        [switch] $ExpectFailure
    )

    $report = Join-Path $testRoot ([guid]::NewGuid().ToString("N") + ".report")
    $arguments = @(
        "-NoProfile", "-File", $helper, "-Mode", $Mode, "-OutputFormat", $Format,
        "-Left", (Join-Path $testRoot $LeftName), "-Right", (Join-Path $testRoot $RightName),
        "-Output", $report
    )
    if ($BCompPath) { $arguments += @("-BCompPath", $BCompPath) }
    if ($Strict) { $arguments += "-FailOnConversionError" }
    try {
        # Windows PowerShell turns native stderr into error records, even for expected failures.
        $ErrorActionPreference = "Continue"
        $console = & $shell @arguments 2>&1
        $code = $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = "Stop"
    }
    Assert-True (($code -ne 0) -eq [bool]$ExpectFailure) "Unexpected exit $code for $LeftName / $RightName : $console"
    Assert-True (Test-Path -LiteralPath $report) "Report was not written."
    if ($Format -eq "Jsonl") {
        Get-Content -LiteralPath $report | ForEach-Object { $_ | ConvertFrom-Json }
    }
}

function Assert-Outcome {
    param([object[]] $Records, [string] $Expected)
    $summaries = @($Records | Where-Object { $_.type -eq "summary" })
    Assert-True ($summaries.Count -eq 1) "Expected exactly one summary."
    Assert-True ($Records[1].type -eq "summary") "Summary must immediately follow metadata."
    Assert-True ($summaries[0].outcome -eq $Expected) "Expected $Expected; got $($summaries[0].outcome)."
}

try {
    Write-Fixture "left.txt" "Start`nBudget: 100`nEnd`n"
    Write-Fixture "right.txt" "Start`nBudget: 120`nExtra`nEnd`n"
    Write-Fixture "same.txt" "Start`nBudget: 100`nEnd`n"
    Write-Fixture "crlf.txt" "Start`r`nBudget: 100`r`nEnd`r`n"
    Write-Fixture "phrase-left.txt" "Left error: Conversion Error`n"
    Write-Fixture "phrase-right.txt" "Left error: Conversion Error resolved`n"

    $changed = @(Invoke-Comparison "left.txt" "right.txt")
    Assert-Outcome $changed "different"
    Assert-True ($changed[0].schema_version -eq 2) "Expected schema version 2."
    Assert-True (@($changed | Where-Object { $_.type -eq "report_line" -and $_.text -match "Budget: 100" }).Count -gt 0) "Original report evidence was lost."
    Assert-Outcome @(Invoke-Comparison "left.txt" "same.txt") "same"
    Assert-Outcome @(Invoke-Comparison "left.txt" "crlf.txt") "same"

    $phrase = @(Invoke-Comparison "phrase-left.txt" "phrase-right.txt" -Strict)
    Assert-Outcome $phrase "different"
    Assert-True (@($phrase | Where-Object { $_.type -eq "conversion_error" }).Count -eq 0) "File content was mistaken for a diagnostic."
    Invoke-Comparison "phrase-left.txt" "phrase-right.txt" -Format Raw -Strict
    Invoke-Comparison "phrase-left.txt" "phrase-right.txt" -Mode TextHtml -Format Raw -Strict

    # A malformed OLE workbook exercises a real BC conversion diagnostic.
    $oleHeader = [byte[]](0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1)
    [System.IO.File]::WriteAllBytes((Join-Path $testRoot "bad-left.xls"), $oleHeader + [byte[]](0, 1, 2, 3))
    [System.IO.File]::WriteAllBytes((Join-Path $testRoot "bad-right.xls"), $oleHeader + [byte[]](0, 1, 2, 4))
    $failed = @(Invoke-Comparison "bad-left.xls" "bad-right.xls")
    Assert-Outcome $failed "error"
    Assert-True (@($failed | Where-Object { $_.type -eq "conversion_error" }).Count -eq 1) "Real conversion failure was not reported."
    Assert-Outcome @(Invoke-Comparison "bad-left.xls" "bad-right.xls" -Strict -ExpectFailure) "error"
    Invoke-Comparison "bad-left.xls" "bad-right.xls" -Format Raw -Strict -ExpectFailure

    Write-Fixture "old\nested\changed.txt" "before`n"
    Write-Fixture "new\nested\changed.txt" "after!`n"
    Write-Fixture "old\same.txt" "same`n"
    Write-Fixture "new\same.txt" "same`n"
    Write-Fixture "old\removed.txt" "removed`n"
    Write-Fixture "new\added.txt" "added`n"
    Get-ChildItem -LiteralPath (Join-Path $testRoot "old"), (Join-Path $testRoot "new") -Recurse -File |
        ForEach-Object { $_.LastWriteTime = [datetime]'2026-01-01T12:00:00' }
    $folder = @(Invoke-Comparison "old" "new" -Mode Folder)
    Assert-Outcome $folder "different"
    foreach ($expected in @(@("nested/changed.txt", "diff"), @("removed.txt", "ltonly"), @("added.txt", "rtonly"))) {
        $matching = @($folder | Where-Object { $_.type -eq "entry" -and $_.path -eq $expected[0] -and $_.status -eq $expected[1] })
        Assert-True ($matching.Count -eq 1) "Missing folder difference: $($expected -join ' / ')."
    }

    Write-Fixture "equal-left\same.txt" "same`n"
    Write-Fixture "equal-right\same.txt" "same`n"
    Assert-Outcome @(Invoke-Comparison "equal-left" "equal-right" -Mode Folder) "same"
    New-Item -ItemType Directory -Path (Join-Path $testRoot "equal-right\empty-added") | Out-Null
    $emptyFolder = @(Invoke-Comparison "equal-left" "equal-right" -Mode Folder)
    Assert-Outcome $emptyFolder "different"
    Assert-True (@($emptyFolder | Where-Object { $_.type -eq "entry" -and $_.path -eq "empty-added" -and $_.status -eq "rtonly" }).Count -eq 1) "Empty added folder was missed."

    Write-Output "Passed: file outcomes, report evidence, literal diagnostic text, real conversion failures, raw/HTML modes, and folder outcomes including empty added folders."
}
finally {
    $resolvedTestRoot = [System.IO.Path]::GetFullPath($testRoot)
    if ((Split-Path -Parent $resolvedTestRoot) -ne $tempBase -or (Split-Path -Leaf $resolvedTestRoot) -notlike "document-compare-test-*") {
        throw "Refusing to clean up outside the test temporary directory: $resolvedTestRoot"
    }
    Remove-Item -LiteralPath $resolvedTestRoot -Recurse -Force
}
