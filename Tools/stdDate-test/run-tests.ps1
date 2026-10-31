$HostApplication = 'Excel'
$ErrorActionPreference = 'Stop'
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$office = $null
$scratch = Join-Path ([IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString() + '.cls')
$source = Get-Content (Join-Path $repoRoot 'src/WIP/stdDate.cls') -Raw
$source = $source.Replace('  Set CName = CachedNames(key)', '  Set CName = CachedNames(key)' + "`r`n" + '  StdRegexTestRunner.regexCount = CachedRegex.Count' + "`r`n" + '  StdRegexTestRunner.namesCount = CachedNames.Count')
Set-Content $scratch $source
$document = $null
try {
    $office = New-Object -ComObject 'Excel.Application'
    $office.Visible = $false
    $office.DisplayAlerts = 0
    if ($HostApplication -eq 'Excel') { $document = $office.Workbooks.Add() }
    else { $document = $office.Documents.Add() }
    $project = $document.GetType().InvokeMember('VBProject', [System.Reflection.BindingFlags]::GetProperty, $null, $document, $null)
    foreach ($file in @('src\WIP\stdDate.cls', 'Tools\stdRegex-build\TestReporter.cls', 'Tools\stdDate-test\TestRunner.bas', 'tests\stdDateTests.bas')) {
        if ($file -eq 'src\WIP\stdDate.cls') { $null = $project.VBComponents.Import($scratch) }
        else { $null = $project.VBComponents.Import((Join-Path $repoRoot $file)) }
    }
    if ($HostApplication -eq 'Excel') {
        $result = $office.Run("'" + $document.Name + "'!StdRegexTestRunner.RunTests")
    } else {
        $result = $office.Run('StdRegexTestRunner.RunTests')
    }
    if ($result -is [string]) { throw $result }
    if ($result -le 0) { throw 'No stdDate assertions ran.' }
    Write-Output "Passed $result stdDate checks in $HostApplication VBA."
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
