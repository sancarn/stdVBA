param([string]$OutputDirectory = $PSScriptRoot, [switch]$WithTests)
$ErrorActionPreference = 'Stop'
$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
$installation = & $vswhere -latest -products '*' -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath
if (!$installation) { throw 'Install Visual Studio C++ build tools and Windows SDK.' }
$dev = Join-Path $installation 'Common7\Tools\VsDevCmd.bat'
$null = New-Item -ItemType Directory -Force -Path $OutputDirectory
$source = Join-Path $PSScriptRoot 'stdWebSocketNative.c'
foreach ($architecture in @('x86','x64')) {
  $out = Join-Path $OutputDirectory $architecture
  $null = New-Item -ItemType Directory -Force -Path $out
  $definition = Join-Path $out 'exports.def'
  if ($architecture -eq 'x86') {
    [IO.File]::WriteAllText($definition, "EXPORTS`r`nstdWebSocketOperation=_stdWebSocketOperation@32`r`nstdWebSocketConnect=_stdWebSocketConnect@24`r`nstdWebSocketDisconnect=_stdWebSocketDisconnect@4`r`n")
  } else {
    [IO.File]::WriteAllText($definition, "EXPORTS`r`nstdWebSocketOperation`r`nstdWebSocketConnect`r`nstdWebSocketDisconnect`r`n")
  }
  $exports = [IO.File]::ReadAllText($definition)
  foreach ($export in @(@('stdWebSocketListen',8),@('stdWebSocketListening',4),@('stdWebSocketTerminal',4),@('stdWebSocketNext',16),@('stdWebSocketUnlisten',8))) {
    $exports += $export[0]
    if ($architecture -eq 'x86') { $exports += "=_$($export[0])@$($export[1])" }
    $exports += "`r`n"
  }
  [IO.File]::WriteAllText($definition, $exports)
  $dll = Join-Path $out 'stdWebSocketNative.dll'
  $obj = Join-Path $out 'stdWebSocketNative.obj'
  $lib = Join-Path $out 'stdWebSocketNative.lib'
  $batch = Join-Path $out 'build.cmd'
  $installer = Split-Path $vswhere
  $commands = "@echo off`r`nset `"PATH=$installer;%PATH%`"`r`ncall `"$dev`" -no_logo -arch=$architecture`r`nif errorlevel 1 exit /b 1`r`ncl /nologo /W4 /WX /O2 /LD /MT /Fo`"$obj`" `"$source`" /link winhttp.lib user32.lib oleaut32.lib uuid.lib /DEF:`"$definition`" /OUT:`"$dll`" /IMPLIB:`"$lib`"`r`nif errorlevel 1 exit /b 1`r`n"
  if ($WithTests) {
    $test = (Resolve-Path (Join-Path $PSScriptRoot '../../../../Tools/stdWebSocket-tests/native_tests.c')).Path
    $testObj = Join-Path $out 'native_tests.obj'
    $testExe = Join-Path $out 'native_tests.exe'
    $commands += "cl /nologo /W4 /WX /O2 /MT /Fo`"$testObj`" /Fe`"$testExe`" `"$test`" /link winhttp.lib user32.lib oleaut32.lib uuid.lib`r`n"
  }
  [IO.File]::WriteAllText($batch, $commands)
  & cmd /d /c $batch
  if ($LASTEXITCODE) { throw "Native $architecture build failed" }
  Write-Output $dll
}
