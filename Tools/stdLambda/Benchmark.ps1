param(
    [switch]$TemporaryTrustAccess,
    [string]$OutputDirectory = (Join-Path $PSScriptRoot 'benchmarks/latest'),
    [string[]]$Engines = @('stdLambda', 'ASF')
)
# stdLambda/ASF benchmark in Excel, with the runner of Tools/stdRegex. Imports ASF (src/) and
# stdLambda (stdVBA at the pinned commit) into a workbook of a separate Excel process, runs
# Benchmark.bas for every workload, operation and engine and writes samples.csv and samples.json
# to OutputDirectory.
$ErrorActionPreference = 'Stop'
$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$stdVbaCommit = '230aea3aef43cd4bb78a728e7a18f70e11b9a77b'
$stdVbaRaw = "https://raw.githubusercontent.com/sancarn/stdVBA/$stdVbaCommit/src"
$ansi = [Text.Encoding]::GetEncoding(1252)
$sources = Join-Path ([IO.Path]::GetTempPath()) ('lambda-benchmark-' + [guid]::NewGuid())
New-Item -ItemType Directory -Path $sources | Out-Null

# stdLambda and the interface it implements, with CRLF line ends for VBComponents.Import
foreach ($name in @('stdICallable', 'stdLambda')) {
    $path = Join-Path $sources ($name + '.cls')
    Invoke-WebRequest -Uri "$stdVbaRaw/$name.cls" -OutFile $path
    $text = [IO.File]::ReadAllText($path, $ansi).Replace("`r`n", "`n").Replace("`n", "`r`n")
    [IO.File]::WriteAllText($path, $text, $ansi)
}

New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$output = Join-Path ([IO.Path]::GetFullPath($OutputDirectory)) 'samples.csv'
$excel = $null
$book = $null
$process = $null
$seedPath = Join-Path ([IO.Path]::GetTempPath()) ('lambda-benchmark-' + [guid]::NewGuid() + '.xlsx')
$securityKey = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
$previousTrust = Get-ItemProperty -Path $securityKey -Name AccessVBOM -ErrorAction SilentlyContinue
$debuggerKey = 'HKCU:\Software\Microsoft\VBA\7.1\Common'
$previousDebugger = Get-ItemProperty -Path $debuggerKey -ErrorAction SilentlyContinue
try {
    # /x starts a separate process, so existing user workbooks and cached security settings
    # cannot affect the run. Use Excel's native object model to attach to that process.
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
    # Engines raise errors the benchmark handles; let them reach its handler.
    New-Item -Path $debuggerKey -Force | Out-Null
    Set-ItemProperty -Path $debuggerKey -Name BreakOnServerErrors -Value 0 -Type DWord
    Set-ItemProperty -Path $debuggerKey -Name BreakOnAllErrors -Value 0 -Type DWord
    Add-Type @'
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class LambdaBenchmarkExcelWindow {
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
        $window = [LambdaBenchmarkExcelWindow]::GetNativeObject($process.MainWindowHandle)
        if ($null -ne $window) { $excel = $window.Application; break }
        Start-Sleep -Milliseconds 200
    } while ([DateTime]::UtcNow -lt $deadline)
    if ($null -eq $excel) { throw 'Unable to attach to the isolated Excel process.' }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add()
    if ($null -eq $book.VBProject) { throw 'Excel must trust access to the VBA project object model to import the modules.' }
    $classes = @(Get-ChildItem -Path (Join-Path $root 'src') -Filter '*.cls' | ForEach-Object { $_.FullName })
    $classes += @((Join-Path $sources 'stdICallable.cls'), (Join-Path $sources 'stdLambda.cls'))
    foreach ($path in $classes) {
        $imported = $book.VBProject.VBComponents.Import($path)
        Write-Output "Imported $($imported.Name): $($imported.CodeModule.CountOfLines) lines."
    }
    $benchmarkModule = $book.VBProject.VBComponents.Import((Join-Path $PSScriptRoot 'Benchmark.bas'))
    Write-Output "Imported benchmark module $($benchmarkModule.Name): $($benchmarkModule.CodeModule.CountOfLines) lines."

    $header = 'engine,workload,operation,status,note,iterations,sample_1_us,sample_2_us,sample_3_us,sample_4_us,sample_5_us'
    [IO.File]::WriteAllText($output, $header + "`n", [Text.UTF8Encoding]::new($false))
    $workloads = @('Arithmetic expression', 'Predicate', 'String functions', 'Conditional', 'Local variables',
                   'Object members', 'Loop', 'While and branches', 'Recursion', 'Callback over 300 items')
    $operations = @('Compile', 'Run')
    $cell = 0
    foreach ($workload in $workloads) {
        foreach ($operation in $operations) {
            # Rotate engine order to reduce a fixed warm-up / system-load bias.
            for ($offset = 0; $offset -lt $Engines.Count; $offset++) {
                $engineName = $Engines[($cell + $offset) % $Engines.Count]
                $row = $excel.Run("'" + $book.Name + "'!LambdaBenchmark.Measure", $engineName, $workload, $operation)
                [IO.File]::AppendAllText($output, $row + "`n", [Text.UTF8Encoding]::new($false))
                Write-Output $row
            }
            $cell++
        }
    }
    $metadata = @{
        measured_at_utc = [DateTimeOffset]::UtcNow.ToString('o'); office = "Excel $($excel.Version) build $($excel.Build)"
        operating_system = $excel.OperatingSystem
        cpu = (Get-CimInstance Win32_Processor | Select-Object -First 1 -ExpandProperty Name)
        engines = $Engines
        asf_commit = (git -C $root rev-parse HEAD)
        stdvba_commit = $stdVbaCommit
        benchmark_sha256 = (Get-FileHash (Join-Path $PSScriptRoot 'Benchmark.bas')).Hash
        samples = 5; target_batch_ms = 30; timer = 'QueryPerformanceCounter'
        description = 'Compile: fresh object + compile (stdLambda protInit, ASF New + Compile). Run: one call of the compiled object with the workload''s arguments. Results checked before each timed cell.'
    }
    $metadata | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath ([IO.Path]::ChangeExtension($output, '.json')) -Encoding utf8
    Write-Output "Benchmark samples saved to $output"
} finally {
    try {
        if ($null -ne $book) { $book.Close($false) }
        if ($null -ne $excel) { $excel.Quit() }
        foreach ($obj in @($book, $window, $excel)) {
            if ($null -ne $obj -and [Runtime.InteropServices.Marshal]::IsComObject($obj)) {
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($obj)
            }
        }
    } finally {
        # Stop only the disposable /x process if outstanding COM references keep it alive.
        if ($null -ne $process -and -not $process.WaitForExit(3000)) { Stop-Process -Id $process.Id }
        Remove-Item -LiteralPath $seedPath -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $sources -Recurse -ErrorAction SilentlyContinue
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
