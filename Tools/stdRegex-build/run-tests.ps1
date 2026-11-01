param(
 [string]$EngineRoot = (Join-Path $PSScriptRoot '..\..\..\vba-regex'),
 [ValidateSet('Excel','Word')][string]$HostApplication = 'Excel'
)
& python (Join-Path $PSScriptRoot 'build.py') --engine-root $EngineRoot --check
if ($LASTEXITCODE) { exit $LASTEXITCODE }
& python (Join-Path $PSScriptRoot '../Testing/run.py') --suite stdRegex --host $HostApplication
exit $LASTEXITCODE
