param(
    [string]$BaselineRef = 'c02d86fd7e5c99b648f9729d7a0c6160d492a642',
    [ValidateSet('Excel', 'Word')][string]$HostApplication = 'Word',
    [ValidateRange(3, 31)][int]$Samples = 7,
    [ValidateRange(0.01, 1.0)][double]$TargetSeconds = 0.05,
    [string]$OutputDir = (Join-Path $PSScriptRoot 'benchmarks\latest')
)

# The baseline is the pure-VBA stdRegex before the list/array changes, not the
# older VBScript wrapper. Both classes run in the same hidden Office process.
$ErrorActionPreference = 'Stop'
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$baselineCommit = (& git -C $repoRoot rev-parse --verify --end-of-options ($BaselineRef + '^{commit}')).Trim()
if ($LASTEXITCODE -ne 0) { throw 'Cannot resolve baseline commit.' }
$baselineLines = & git -C $repoRoot show ($baselineCommit + ':src/stdRegex.cls')
if ($LASTEXITCODE -ne 0) { throw 'Cannot read baseline class.' }
$baseline = [regex]::Replace(($baselineLines -join "`r`n"), '\bstdRegex\b', 'stdRegexBaseline') + "`r`n"
$tempRoot = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath())
$workDir = Join-Path $tempRoot ('stdRegex-benchmark-' + [guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $workDir
$baselinePath = Join-Path $workDir 'stdRegexBaseline.cls'
[System.IO.File]::WriteAllText($baselinePath, $baseline, [System.Text.UTF8Encoding]::new($false))
$office = $null
$document = $null
$rows = [System.Collections.Generic.List[object]]::new()
try {
    $office = New-Object -ComObject ($HostApplication + '.Application')
    $office.Visible = $false
    $office.DisplayAlerts = 0
    if ($HostApplication -eq 'Excel') { $document = $office.Workbooks.Add() }
    else { $document = $office.Documents.Add() }
    $project = $document.GetType().InvokeMember('VBProject', [System.Reflection.BindingFlags]::GetProperty, $null, $document, $null)
    foreach ($path in @($baselinePath, (Join-Path $repoRoot 'src\stdRegex.cls'), (Join-Path $PSScriptRoot 'Benchmark.bas'))) {
        $null = $project.VBComponents.Import($path)
    }
    $prefix = if ($HostApplication -eq 'Excel') { "'" + $document.Name + "'!" } else { '' }
    $officeVersion = $office.Version
    $vbaArchitecture = $office.Run($prefix + 'StdRegexBenchmark.HostDetails')
    for ($caseId = 0; $caseId -lt 6; $caseId++) {
        $values = $office.Run($prefix + 'StdRegexBenchmark.RunCase', $caseId, $Samples, $TargetSeconds)
        if ($values -is [string]) { throw $values }
        for ($i = 1; $i -le $values.GetLength(0); $i++) {
            $rows.Add([pscustomobject]@{
                case = $values[$i, 1]; matches = [int]$values[$i, 2]; captures = [int]$values[$i, 3]
                input_characters = [int]$values[$i, 4]; operation = $values[$i, 5]
                sample = [int]$values[$i, 6]; iterations = [int]$values[$i, 7]
                elapsed_ms = [double]$values[$i, 8]; ms_per_call = [double]$values[$i, 9]
            })
        }
        Write-Output ('Measured: ' + $values[1, 1])
    }
    $metadata = [ordered]@{
        measured_at_utc = [DateTime]::UtcNow.ToString('o')
        host_application = $HostApplication; office_version = $officeVersion; vba_architecture = $vbaArchitecture
        operating_system = [System.Environment]::OSVersion.VersionString
        baseline_commit = $baselineCommit
        current_class_sha256 = (Get-FileHash -LiteralPath (Join-Path $repoRoot 'src\stdRegex.cls') -Algorithm SHA256).Hash
        engine = (Get-Content -LiteralPath (Join-Path $PSScriptRoot 'engine.json') -Raw | ConvertFrom-Json)
        samples = $Samples; minimum_batch_seconds = $TargetSeconds
        methodology = 'QueryPerformanceCounter; compiled regexes reused; 2 warm-up calls; calibrated batches; rotated operation order; includes result allocation and release; output equality checked before timing.'
    }
    $null = New-Item -ItemType Directory -Path $OutputDir -Force
    $metadata | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $OutputDir 'environment.json') -Encoding utf8
    $rows | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $OutputDir 'samples.json') -Encoding utf8
    $rows | Export-Csv -LiteralPath (Join-Path $OutputDir 'samples.csv') -NoTypeInformation -Encoding utf8
    & python (Join-Path $PSScriptRoot 'summarize-benchmark.py') $OutputDir
    if ($LASTEXITCODE -ne 0) { throw 'Benchmark summarization failed.' }
} finally {
    try {
        if ($null -ne $document) {
            if ($HostApplication -eq 'Excel') { $document.Close($false) }
            else { $document.Close(0) }
        }
    } finally {
        try { if ($null -ne $office) { $office.Quit() } }
        finally {
            $resolvedWork = [System.IO.Path]::GetFullPath($workDir)
            $resolvedTemp = [System.IO.Path]::GetFullPath($tempRoot).TrimEnd('\', '/') + [System.IO.Path]::DirectorySeparatorChar
            if (-not $resolvedWork.StartsWith($resolvedTemp, [System.StringComparison]::OrdinalIgnoreCase) -or
                -not ([System.IO.Path]::GetFileName($resolvedWork)).StartsWith('stdRegex-benchmark-')) { throw 'Unsafe scratch cleanup path.' }
            Remove-Item -LiteralPath $resolvedWork -Recurse -Force
        }
    }
}
