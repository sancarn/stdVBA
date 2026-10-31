param([string]$PidFile)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path "$PSScriptRoot/../..").Path
Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class HostPid {
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
}
"@
$existingIds = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$excel = $null
$book = $null
try {
  $excel = New-Object -ComObject Excel.Application
  [uint32]$hostId = 0
  [HostPid]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$hostId) | Out-Null
  if ($existingIds -contains $hostId) {
    $excel = $null
    throw 'Office returned an existing session; refusing to use it.'
  }
  Set-Content -LiteralPath $PidFile -Value $hostId
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $book = $excel.Workbooks.Add()
  $project = $book.GetType().InvokeMember('VBProject', [System.Reflection.BindingFlags]::GetProperty, $null, $book, $null)
  if ($null -eq $project) { throw 'VBA project access unavailable; security settings were not changed.' }
  $vbeVersion = [version]$excel.VBE.Version
  $vbeKey = "HKCU:\Software\Microsoft\VBA\$($vbeVersion.Major).$($vbeVersion.Minor)\Common"
  $trapping = Get-ItemProperty -LiteralPath $vbeKey
  Write-Output "VBE error trapping: BreakOnAllErrors=$($trapping.BreakOnAllErrors), BreakOnServerErrors=$($trapping.BreakOnServerErrors)"
  if ($trapping.BreakOnAllErrors -ne 0 -or $trapping.BreakOnServerErrors -ne 0) {
    throw 'Tests require Break on Unhandled Errors. Debugging preferences were not changed.'
  }
  foreach ($file in @('src/stdICallable.cls','src/stdHTTP.cls','tests/stdHTTPLocalTests.bas')) {
    $null = $project.VBComponents.Import((Join-Path $root $file))
  }
  $result = $excel.Run("'" + $book.Name + "'!stdHTTPLocalTests.RunTests", $env:STDHTTP_URL, $env:STDHTTP_REFUSED_URL)
  Write-Output $result
  if ($result -ne 'Passed local HTTP regressions') { throw "Unexpected test result: $result" }
} finally {
  try { if ($book) { $book.Close($false) } }
  finally { if ($excel) { $excel.Quit() } }
}
