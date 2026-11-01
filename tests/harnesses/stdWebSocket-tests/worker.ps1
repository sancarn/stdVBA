param([string]$PidFile)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path "$PSScriptRoot/../../..").Path
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
$snapshots = @()
function Record-RegistryValue([string]$Path, [string]$Name) {
  $key = Get-Item -LiteralPath $Path -ErrorAction SilentlyContinue
  $exists = $key -and ($key.GetValueNames() -contains $Name)
  $value = $null
  $kind = $null
  if ($exists) { $value = $key.GetValue($Name); $kind = $key.GetValueKind($Name) }
  Write-Output "Original registry: $Path $Name exists=$exists type=$kind value=$value"
  $script:snapshots += [pscustomobject]@{ Path=$Path; Name=$Name; Exists=$exists; Value=$value; Kind=$kind }
}
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
  $securityKey = "HKCU:\Software\Microsoft\Office\$($excel.Version)\Excel\Security"
  Record-RegistryValue $securityKey 'AccessVBOM'
  $access = Get-Item -LiteralPath $securityKey
  Write-Output "Actual Excel $($excel.Version) AccessVBOM: $($access.GetValueKind('AccessVBOM'))=$($access.GetValue('AccessVBOM')) (unchanged)"
  $vbeVersion = [version]$excel.VBE.Version
  $vbeKey = "HKCU:\Software\Microsoft\VBA\$($vbeVersion.Major).$($vbeVersion.Minor)\Common"
  $trapping = Get-ItemProperty -LiteralPath $vbeKey
  Record-RegistryValue $vbeKey 'BreakOnAllErrors'
  Record-RegistryValue $vbeKey 'BreakOnServerErrors'
  Write-Output "VBE error trapping: BreakOnAllErrors=$($trapping.BreakOnAllErrors), BreakOnServerErrors=$($trapping.BreakOnServerErrors)"
  if ($trapping.BreakOnAllErrors -ne 0 -or $trapping.BreakOnServerErrors -ne 0) {
    throw 'Tests require Break on Unhandled Errors. Debugging preferences were not changed.'
  }
  $book = $excel.Workbooks.Add()
  $project = $book.GetType().InvokeMember('VBProject', [System.Reflection.BindingFlags]::GetProperty, $null, $book, $null)
  if ($null -eq $project) { throw 'VBA project access unavailable; security settings were not changed.' }
  $files = @('src/stdWebSocket.cls','tests/harnesses/stdWebSocket-tests/WebSocketRecorder.cls','tests/stdWebSocketLocalTests.bas','src/stdICallable.cls','tests/harnesses/stdWebSocket-tests/WebSocketAsyncRecorder.cls','tests/stdWebSocketAsyncTests.bas') | ForEach-Object { Join-Path $root $_ }
  if ($env:STDWEBSOCKET_WIP -eq '1') {
    $generated = Join-Path (Split-Path $PidFile) 'generated'
    $null = New-Item -ItemType Directory -Path $generated -Force
    $files[0] = Join-Path $root 'src/WIP/stdWebSocket/stdWebSocketNew.cls'
    for ($i = 1; $i -lt $files.Count; $i++) {
      $text = [IO.File]::ReadAllText($files[$i]) -replace '\bstdWebSocket\b', 'stdWebSocketNew'
      $target = Join-Path $generated (Split-Path $files[$i] -Leaf)
      [IO.File]::WriteAllText($target, $text, (New-Object Text.UTF8Encoding($false)))
      $files[$i] = $target
    }
  }
  foreach ($file in $files) {
    $bytes = [IO.File]::ReadAllBytes($file)
    if ([Text.Encoding]::UTF8.GetString($bytes) -match '(?<!\r)\n|\r(?!\n)') { throw "Non-CRLF bytes before import: $file" }
    $null = $project.VBComponents.Import($file)
  }
  $architecture = $excel.Run("'" + $book.Name + "'!stdWebSocketLocalTests.Architecture")
  Write-Output "Actual Office architecture: $architecture"
  $native = Join-Path $env:STDWEBSOCKET_NATIVE_ROOT "$architecture/stdWebSocketNative.dll"
  $result = $excel.Run("'" + $book.Name + "'!stdWebSocketLocalTests.RunTests", $env:STDWEBSOCKET_URL, $env:STDWEBSOCKET_REFUSED_URL, $native, $env:STDWEBSOCKET_TLS)
  Write-Output $result
  if ($result -notlike 'Passed * local WebSocket assertions') { throw "Unexpected test result: $result" }
  $result = $excel.Run("'" + $book.Name + "'!stdWebSocketAsyncTests.BeginIdle", $env:STDWEBSOCKET_URL, $native)
  Write-Output $result
  Start-Sleep -Milliseconds 600
  $result = $excel.Run("'" + $book.Name + "'!stdWebSocketAsyncTests.CheckIdle")
  Write-Output $result
  if ($result -ne 'Passed real idle delivery') { throw $result }
  $result = $excel.Run("'" + $book.Name + "'!stdWebSocketAsyncTests.RunTests", $env:STDWEBSOCKET_URL, $native)
  Write-Output $result
  if ($result -notlike 'Passed * asynchronous WebSocket assertions') { throw $result }
} finally {
  try { if ($book) { $book.Close($false) } }
  finally {
    try { if ($excel) { $excel.Quit() } }
    finally {
      foreach ($snapshot in $snapshots) {
        $key = Get-Item -LiteralPath $snapshot.Path -ErrorAction SilentlyContinue
        $exists = $key -and ($key.GetValueNames() -contains $snapshot.Name)
        if ($snapshot.Exists) {
          if (!$exists -or $key.GetValue($snapshot.Name) -ne $snapshot.Value -or $key.GetValueKind($snapshot.Name) -ne $snapshot.Kind) {
            if (!$key) { $key = New-Item -Path $snapshot.Path -Force }
            $key.SetValue($snapshot.Name, $snapshot.Value, $snapshot.Kind)
          }
        } elseif ($exists) { $key.DeleteValue($snapshot.Name) }
        Write-Output "Verified original registry state: $($snapshot.Name)"
      }
    }
  }
}
