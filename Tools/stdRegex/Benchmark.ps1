param(
    [switch]$TemporaryTrustAccess,
    [string]$AsfSource,
    [string]$OutputDirectory = (Join-Path $PSScriptRoot 'benchmarks/latest')
)
$ErrorActionPreference = 'Stop'
$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$sources = Join-Path ([IO.Path]::GetTempPath()) ('regex-benchmark-' + [guid]::NewGuid())
New-Item -ItemType Directory -Path $sources | Out-Null
$legacy = (git -C $root show abf14dd81e553cbcf438284975c80f3e9220196a:src/stdRegex.cls) -join "`r`n"
if ($LASTEXITCODE -ne 0) { throw 'Cannot read the pinned old stdRegex revision.' }
$legacy = [regex]::Replace($legacy, '\bstdRegex\b', 'legacyStdRegex')
[IO.File]::WriteAllText((Join-Path $sources 'legacyStdRegex.cls'), $legacy + "`r`n", [Text.UTF8Encoding]::new($false))
if ($AsfSource) {
    $asfText = [IO.File]::ReadAllText($AsfSource).Replace("`r`n", "`n").Replace("`n", "`r`n")
    [IO.File]::WriteAllText((Join-Path $sources 'ASF_RegexEngine.cls'), $asfText, [Text.UTF8Encoding]::new($false))
} else {
    Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/ECP-Solutions/ASF/290ad3eabff97fc16f34568d83139994f89b83ac/src/ASF_RegexEngine.cls' -OutFile (Join-Path $sources 'ASF_RegexEngine.cls')
}
$asfPath = Join-Path $sources 'ASF_RegexEngine.cls'
$asfText = [IO.File]::ReadAllText($asfPath).Replace("`r`n", "`n").Replace("`n", "`r`n")
[IO.File]::WriteAllText($asfPath, $asfText, [Text.UTF8Encoding]::new($false))
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$output = Join-Path ([IO.Path]::GetFullPath($OutputDirectory)) 'samples.csv'
& (Join-Path $PSScriptRoot 'Test.ps1') -TemporaryTrustAccess:$TemporaryTrustAccess -BenchmarkSources $sources -BenchmarkOutputPath $output
Write-Output "Benchmark samples saved to $output"
