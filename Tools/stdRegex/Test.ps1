param(
    [switch]$TemporaryTrustAccess,
    [string]$BenchmarkSources,
    [string]$BenchmarkOutputPath
)
$ErrorActionPreference = 'Stop'
$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$excel = $null
$book = $null
$process = $null
$seedPath = Join-Path ([IO.Path]::GetTempPath()) ('stdRegex-' + [guid]::NewGuid() + '.xlsx')
$securityKey = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
$previousTrust = Get-ItemProperty -Path $securityKey -Name AccessVBOM -ErrorAction SilentlyContinue
$debuggerKey = 'HKCU:\Software\Microsoft\VBA\7.1\Common'
$previousDebugger = Get-ItemProperty -Path $debuggerKey -ErrorAction SilentlyContinue
try {
    # /x starts a separate process, so existing user workbooks and cached security settings
    # cannot affect the test. Use Excel's native object model to attach to that process.
    $seedExcel = New-Object -ComObject Excel.Application
    $seedExcel.DisplayAlerts = $false
    $seedBook = $seedExcel.Workbooks.Add()
    $seedBook.SaveAs($seedPath, 51)
    $seedBook.Close($false)
    $excelExe = Join-Path $seedExcel.Path 'EXCEL.EXE'
    $seedExcel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($seedExcel)
    if ($TemporaryTrustAccess) {
        New-Item -Path $securityKey -Force | Out-Null
        Set-ItemProperty -Path $securityKey -Name AccessVBOM -Value 1 -Type DWord
    }
    # Tests intentionally raise errors. Let the caller's handler receive class errors.
    New-Item -Path $debuggerKey -Force | Out-Null
    Set-ItemProperty -Path $debuggerKey -Name BreakOnServerErrors -Value 0 -Type DWord
    Set-ItemProperty -Path $debuggerKey -Name BreakOnAllErrors -Value 0 -Type DWord
    Add-Type @'
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class RegexTestExcelWindow {
    public delegate bool EnumProc(IntPtr hwnd, IntPtr param);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent, EnumProc callback, IntPtr param);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassName(IntPtr hwnd, StringBuilder name, int count);
    [DllImport("oleacc.dll")] public static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint id, ref Guid iid, [MarshalAs(UnmanagedType.IDispatch)] out object obj);
    public static object GetNativeObject(IntPtr main) {
        object result = null;
        EnumChildWindows(main, (hwnd, param) => {
            var name = new StringBuilder(256);
            GetClassName(hwnd, name, 256);
            if (name.ToString() != "EXCEL7") return true;
            Guid iid = new Guid("00020400-0000-0000-C000-000000000046");
            AccessibleObjectFromWindow(hwnd, 0xFFFFFFF0, ref iid, out result);
            return result == null;
        }, IntPtr.Zero);
        return result;
    }
}
'@
    $process = Start-Process -FilePath $excelExe -ArgumentList @('/x', '/automation', ('"' + $seedPath + '"')) -WindowStyle Hidden -PassThru
    $deadline = [DateTime]::UtcNow.AddSeconds(30)
    do {
        $process.Refresh()
        $window = [RegexTestExcelWindow]::GetNativeObject($process.MainWindowHandle)
        if ($null -ne $window) { $excel = $window.Application; break }
        Start-Sleep -Milliseconds 200
    } while ([DateTime]::UtcNow -lt $deadline)
    if ($null -eq $excel) { throw 'Unable to attach to the isolated Excel test process.' }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add()
    if ($null -eq $book.VBProject) { throw 'Excel must trust access to the VBA project object model to import the test modules.' }
    $null = $book.VBProject.VBComponents.Import((Join-Path $root 'src/stdRegex.cls'))
    $null = $book.VBProject.VBComponents.Import((Join-Path $root 'tests/stdRegexTests.bas'))
    $assertions = $book.VBProject.VBComponents.Add(1)
    $assertions.Name = 'Test'
    $assertions.CodeModule.AddFromString(@'
Option Explicit
Private passed As Long
Private lastAssertion As String
Public Sub Topic(ByVal name As String)
End Sub
Public Sub Assert(ByVal name As String, ByVal result As Boolean)
    If Not result Then Err.Raise 513, "stdRegexTests", name
    passed = passed + 1
    lastAssertion = name
End Sub
Public Function RunAll() As String
    On Error GoTo Failed
    stdRegexTests.testAll
    RunAll = "PASS: " & passed
    Exit Function
Failed:
    RunAll = "FAIL after " & passed & " assertions (last: " & lastAssertion & "): " & Err.Number & " " & Err.Source & " " & Err.Description
End Function
Public Function PassedCount() As Long
    PassedCount = passed
End Function
'@)
    $result = $excel.Run("'" + $book.Name + "'!Test.RunAll")
    if (-not $result.StartsWith('PASS:')) { throw $result }
    $passed = $excel.Run("'" + $book.Name + "'!Test.PassedCount")
    Write-Output "$passed regex assertions passed in Excel $($excel.Version) ($($excel.OperatingSystem))."
    if ($BenchmarkSources) {
        foreach ($source in @('legacyStdRegex.cls', 'ASF_RegexEngine.cls')) {
            $imported = $book.VBProject.VBComponents.Import((Join-Path $BenchmarkSources $source))
            Write-Output "Imported $($imported.Name): $($imported.CodeModule.CountOfLines) lines."
        }
        $benchmarkModule = $book.VBProject.VBComponents.Import((Join-Path $PSScriptRoot 'Benchmark.bas'))
        Write-Output "Imported benchmark module $($benchmarkModule.Name): $($benchmarkModule.CodeModule.CountOfLines) lines."
        $header = 'engine,scenario,operation,input_chars,expected_matches,status,note,iterations,sample_1_us,sample_2_us,sample_3_us,sample_4_us,sample_5_us'
        [IO.File]::WriteAllText($BenchmarkOutputPath, $header + "`n", [Text.UTF8Encoding]::new($false))
        $scenarios = @('Literal scan', 'Short validation', 'Capture extraction', 'Alternation and ignore case', 'Multiline anchors', 'No-match scan', 'Word boundaries', 'Backreference')
        $operations = @('Compile + first Test', 'Test', 'MatchAll', 'Replace')
        $engines = @('Sihlfall', 'ASF', 'Old stdRegex')
        $cell = 0
        foreach ($scenario in $scenarios) {
            foreach ($operation in $operations) {
                # Rotate engine order to reduce a fixed warm-up / system-load bias.
                for ($offset = 0; $offset -lt 3; $offset++) {
                    $engineName = $engines[($cell + $offset) % 3]
                    $row = $excel.Run("'" + $book.Name + "'!RegexBenchmark.Measure", $engineName, $scenario, $operation)
                    [IO.File]::AppendAllText($BenchmarkOutputPath, $row + "`n", [Text.UTF8Encoding]::new($false))
                    Write-Output $row
                }
                $cell++
            }
        }
        foreach ($kind in @('Dictionary', 'Collection')) {
            $row = $excel.Run("'" + $book.Name + "'!RegexBenchmark.MeasureAllocations", $kind)
            [IO.File]::AppendAllText($BenchmarkOutputPath, $row + "`n", [Text.UTF8Encoding]::new($false))
            Write-Output $row
        }
        $metadata = @{
            measured_at_utc = [DateTimeOffset]::UtcNow.ToString('o'); excel_version = $excel.Version; excel_build = $excel.Build
            operating_system = $excel.OperatingSystem
            cpu = (Get-CimInstance Win32_Processor | Select-Object -First 1 -ExpandProperty Name)
            migrated_commit = (git -C $root rev-parse HEAD)
            migrated_sha256 = (Get-FileHash (Join-Path $root 'src/stdRegex.cls')).Hash
            benchmark_sha256 = (Get-FileHash (Join-Path $PSScriptRoot 'Benchmark.bas')).Hash
            legacy_commit = 'abf14dd81e553cbcf438284975c80f3e9220196a'
            sihlfall_commit = '1619cd75526903045b8c5cb8c6048f93bc4160be'
            asf_commit = '290ad3eabff97fc16f34568d83139994f89b83ac'
            samples = 5; target_batch_ms = 30; timer = 'QueryPerformanceCounter'
            engine_step_limits = 2147483647
            description = 'Public API timings: migrated and legacy stdRegex wrappers; standalone ASF. Fresh object + initialization + first Test for compile; reused preinitialized objects for other operations. Correctness checked before each timed cell.'
        }
        $metadata | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath ([IO.Path]::ChangeExtension($BenchmarkOutputPath, '.json')) -Encoding utf8
    }
} finally {
    try {
        if ($null -ne $book) { $book.Close($false) }
        if ($null -ne $excel) { $excel.Quit() }
        foreach ($obj in @($assertions, $book, $window, $excel)) {
            if ($null -ne $obj -and [Runtime.InteropServices.Marshal]::IsComObject($obj)) {
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($obj)
            }
        }
    } finally {
        # Stop only the disposable /x process if outstanding COM references keep it alive.
        if ($null -ne $process -and -not $process.WaitForExit(3000)) { Stop-Process -Id $process.Id }
        Remove-Item -LiteralPath $seedPath -ErrorAction SilentlyContinue
        if ($TemporaryTrustAccess) {
            if ($null -ne $previousTrust) {
                Set-ItemProperty -Path $securityKey -Name AccessVBOM -Value $previousTrust.AccessVBOM -Type DWord
            } else {
                Remove-ItemProperty -Path $securityKey -Name AccessVBOM -ErrorAction SilentlyContinue
            }
        }
        foreach ($setting in @('BreakOnServerErrors', 'BreakOnAllErrors')) {
            if ($null -ne $previousDebugger -and $null -ne $previousDebugger.$setting) {
                Set-ItemProperty -Path $debuggerKey -Name $setting -Value $previousDebugger.$setting -Type DWord
            } else {
                Remove-ItemProperty -Path $debuggerKey -Name $setting -ErrorAction SilentlyContinue
            }
        }
    }
}
