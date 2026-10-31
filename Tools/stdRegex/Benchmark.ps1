param(
    [switch]$TemporaryTrustAccess,
    [string]$OutputDirectory = (Join-Path $PSScriptRoot 'benchmarks/latest'),
    [string[]]$Engines = @('Sihlfall', 'ASF', 'Old stdRegex', 'ASF 290ad3e', 'ASF v3.2.0')
)
# Regex engine benchmark in Excel, adapted from stdVBA's Tools/stdRegex (Benchmark.ps1 and
# Test.ps1, branch codex/issue-141-regex-analysis). Imports the engines into a workbook of a
# separate Excel process, runs Benchmark.bas for every scenario, operation and engine and
# writes samples.csv and samples.json to OutputDirectory.
$ErrorActionPreference = 'Stop'
$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$stdVbaRaw = 'https://raw.githubusercontent.com/sancarn/stdVBA'
$sihlfallCommit = '04e7d048482b0cdf0a3eb3480fc8cd0d3578dc9c'   # stdVBA: stdRegex on sihlfall/vba-regex 1619cd7
$legacyCommit = 'abf14dd81e553cbcf438284975c80f3e9220196a'     # stdVBA: VBScript stdRegex
$ansi = [Text.Encoding]::GetEncoding(1252)
$sources = Join-Path ([IO.Path]::GetTempPath()) ('regex-benchmark-' + [guid]::NewGuid())
New-Item -ItemType Directory -Path $sources | Out-Null

function Get-GitBlob([string]$revision, [string]$path) {
    $info = New-Object Diagnostics.ProcessStartInfo 'git', ('-C "' + $root + '" show ' + $revision + ':' + $path)
    $info.RedirectStandardOutput = $true
    $info.UseShellExecute = $false
    $process = [Diagnostics.Process]::Start($info)
    $buffer = New-Object IO.MemoryStream
    $process.StandardOutput.BaseStream.CopyTo($buffer)
    $process.WaitForExit()
    if ($process.ExitCode -ne 0) { throw "git show $revision`:$path failed." }
    return $ansi.GetString($buffer.ToArray())
}
function Save-Class([string]$name, [string]$text) {
    $text = $text.Replace("`r`n", "`n").Replace("`n", "`r`n")
    [IO.File]::WriteAllText((Join-Path $sources ($name + '.cls')), $text, $ansi)
}

# Sources: stdRegex classes from stdVBA, ASF versions from this repository's history
Invoke-WebRequest -Uri "$stdVbaRaw/$sihlfallCommit/src/stdRegex.cls" -OutFile (Join-Path $sources 'download.cls')
Save-Class 'stdRegex' ([IO.File]::ReadAllText((Join-Path $sources 'download.cls'), $ansi))
Invoke-WebRequest -Uri "$stdVbaRaw/$legacyCommit/src/stdRegex.cls" -OutFile (Join-Path $sources 'download.cls')
Save-Class 'legacyStdRegex' ([regex]::Replace([IO.File]::ReadAllText((Join-Path $sources 'download.cls'), $ansi), '\bstdRegex\b', 'legacyStdRegex'))
Remove-Item (Join-Path $sources 'download.cls')
Save-Class 'ASF_RegexEngine' ([IO.File]::ReadAllText((Join-Path $root 'src/ASF_RegexEngine.cls'), $ansi))
Save-Class 'ASF_RegexEngine_v320' ((Get-GitBlob 'v3.2.0' 'src/ASF_RegexEngine.cls').Replace('ASF_RegexEngine', 'ASF_RegexEngine_v320'))
Save-Class 'ASF_RegexEngine_290ad3e' ((Get-GitBlob '290ad3e' 'src/ASF_RegexEngine.cls').Replace('ASF_RegexEngine', 'ASF_RegexEngine_290ad3e'))

New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$output = Join-Path ([IO.Path]::GetFullPath($OutputDirectory)) 'samples.csv'
$excel = $null
$book = $null
$process = $null
$seedPath = Join-Path ([IO.Path]::GetTempPath()) ('regex-benchmark-' + [guid]::NewGuid() + '.xlsx')
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
public static class RegexBenchmarkExcelWindow {
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
        $window = [RegexBenchmarkExcelWindow]::GetNativeObject($process.MainWindowHandle)
        if ($null -ne $window) { $excel = $window.Application; break }
        Start-Sleep -Milliseconds 200
    } while ([DateTime]::UtcNow -lt $deadline)
    if ($null -eq $excel) { throw 'Unable to attach to the isolated Excel process.' }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add()
    if ($null -eq $book.VBProject) { throw 'Excel must trust access to the VBA project object model to import the modules.' }
    foreach ($source in @('stdRegex', 'legacyStdRegex', 'ASF_RegexEngine', 'ASF_RegexEngine_v320', 'ASF_RegexEngine_290ad3e')) {
        $imported = $book.VBProject.VBComponents.Import((Join-Path $sources ($source + '.cls')))
        Write-Output "Imported $($imported.Name): $($imported.CodeModule.CountOfLines) lines."
    }
    $benchmarkModule = $book.VBProject.VBComponents.Import((Join-Path $PSScriptRoot 'Benchmark.bas'))
    Write-Output "Imported benchmark module $($benchmarkModule.Name): $($benchmarkModule.CodeModule.CountOfLines) lines."

    $header = 'engine,scenario,operation,input_chars,expected_matches,status,note,iterations,sample_1_us,sample_2_us,sample_3_us,sample_4_us,sample_5_us'
    [IO.File]::WriteAllText($output, $header + "`n", [Text.UTF8Encoding]::new($false))
    $scenarios = @('Literal scan', 'Short validation', 'Capture extraction', 'Alternation and ignore case', 'Multiline anchors', 'No-match scan', 'Word boundaries', 'Backreference')
    $operations = @('Compile + first Test', 'Test', 'MatchAll', 'Replace')
    $cell = 0
    foreach ($scenario in $scenarios) {
        foreach ($operation in $operations) {
            # Rotate engine order to reduce a fixed warm-up / system-load bias.
            for ($offset = 0; $offset -lt $Engines.Count; $offset++) {
                $engineName = $Engines[($cell + $offset) % $Engines.Count]
                $row = $excel.Run("'" + $book.Name + "'!RegexBenchmark.Measure", $engineName, $scenario, $operation)
                [IO.File]::AppendAllText($output, $row + "`n", [Text.UTF8Encoding]::new($false))
                Write-Output $row
            }
            $cell++
        }
    }
    foreach ($kind in @('Dictionary', 'Collection')) {
        $row = $excel.Run("'" + $book.Name + "'!RegexBenchmark.MeasureAllocations", $kind)
        [IO.File]::AppendAllText($output, $row + "`n", [Text.UTF8Encoding]::new($false))
        Write-Output $row
    }
    $metadata = @{
        measured_at_utc = [DateTimeOffset]::UtcNow.ToString('o'); office = "Excel $($excel.Version) build $($excel.Build)"
        operating_system = $excel.OperatingSystem
        cpu = (Get-CimInstance Win32_Processor | Select-Object -First 1 -ExpandProperty Name)
        engines = $Engines
        asf_commit = (git -C $root rev-parse HEAD)
        asf_sha256 = (Get-FileHash (Join-Path $root 'src/ASF_RegexEngine.cls')).Hash
        asf_v320_commit = (git -C $root rev-parse 'v3.2.0^{commit}')
        asf_290ad3e_commit = (git -C $root rev-parse '290ad3e^{commit}')
        sihlfall_stdregex_commit = $sihlfallCommit; legacy_stdregex_commit = $legacyCommit
        benchmark_sha256 = (Get-FileHash (Join-Path $PSScriptRoot 'Benchmark.bas')).Hash
        samples = 5; target_batch_ms = 30; timer = 'QueryPerformanceCounter'
        engine_step_limits = 2147483647
        description = 'Public API timings: stdRegex (sihlfall and VBScript) and ASF regex engines. Fresh object + initialization + first Test for compile; reused preinitialized objects for other operations. Correctness checked before each timed cell.'
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
