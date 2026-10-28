param(
    [string]$EngineRoot = (Join-Path $PSScriptRoot '..\..\..\vba-regex'),
    [ValidateSet('Excel', 'Word')][string]$HostApplication = 'Excel'
)

# Requires an Office host with "Trust access to the VBA project object model".
# No security settings are changed by this runner.
$ErrorActionPreference = 'Stop'
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
& python (Join-Path $PSScriptRoot 'build.py') --engine-root $EngineRoot --check
if ($LASTEXITCODE -ne 0) { throw 'Generated stdRegex.cls is out of date.' }

$office = $null
$document = $null
try {
    $office = New-Object -ComObject ($HostApplication + '.Application')
    $office.Visible = $false
    $office.DisplayAlerts = 0
    if ($HostApplication -eq 'Excel') { $document = $office.Workbooks.Add() }
    else { $document = $office.Documents.Add() }
    $project = $document.GetType().InvokeMember('VBProject', [System.Reflection.BindingFlags]::GetProperty, $null, $document, $null)
    foreach ($file in @('src\stdRegex.cls', 'Tools\stdRegex-build\TestReporter.cls', 'Tools\stdRegex-build\TestRunner.bas', 'tests\stdRegexTests.bas')) {
        $null = $project.VBComponents.Import((Join-Path $repoRoot $file))
    }
    if ($HostApplication -eq 'Excel') {
        $result = $office.Run("'" + $document.Name + "'!StdRegexTestRunner.RunTests")
    } else {
        $result = $office.Run('StdRegexTestRunner.RunTests')
    }
    if ($result -is [string]) { throw $result }
    if ($result -le 0) { throw 'No stdRegex assertions ran.' }
    Write-Output "Passed $result stdRegex checks in $HostApplication VBA."
} finally {
    try {
        if ($null -ne $document) {
            if ($HostApplication -eq 'Excel') { $document.Close($false) }
            else { $document.Close(0) }
        }
    } finally {
        if ($null -ne $office) { $office.Quit() }
    }
}
