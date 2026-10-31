$ErrorActionPreference = 'Stop'
$root = (Resolve-Path "$PSScriptRoot/../..").Path
$fixture = Start-Process python -ArgumentList @("`"$PSScriptRoot/fixture.py`"") -WindowStyle Hidden -PassThru
$excel = $null
$book = $null
try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $book = $excel.Workbooks.Add()
  foreach ($file in @('src/stdICallable.cls','src/stdHTTP.cls','tests/stdHTTPLocalTests.bas')) {
    $null = $book.VBProject.VBComponents.Import((Join-Path $root $file))
  }
  $result = $excel.Run("'" + $book.Name + "'!stdHTTPLocalTests.RunTests")
  Write-Output $result
  if ($result -like 'FAIL:*') { throw $result }
} finally {
  if ($book) { $book.Close($false) }
  if ($excel) { $excel.Quit() }
  Stop-Process -Id $fixture.Id -ErrorAction SilentlyContinue
}
