param(
    [string]$Output = (Join-Path $PSScriptRoot '../ASF.xlam'),
    [switch]$TemporaryTrustAccess
)
# Builds the ASF add-in for Excel (excel/ASF.xlam) from the sources, with Excel for Windows:
#   - the engine (src/*.cls), its classes made visible to the workbooks that reference the add-in
#     (VB_Exposed = True: PublicNotCreatable; NewEngine makes them);
#   - the modules of excel/addin and the regex worksheet functions of the use cases;
#   - the descriptions of the functions (ASFRegister.RegisterFunctions);
#   - project name ASFAddIn, saved as an Excel add-in (.xlam).
# Excel must trust access to the VBA project object model (Trust Center > Macro Settings), or run
# with -TemporaryTrustAccess, which turns it on for the build and back as it was.
$ErrorActionPreference = 'Stop'
$Output = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Output)
$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$ansi = [Text.Encoding]::GetEncoding(1252)
$staging = Join-Path ([IO.Path]::GetTempPath()) ('asf-addin-' + [guid]::NewGuid())
New-Item -ItemType Directory -Path $staging | Out-Null

# the files to import, with CRLF line ends and the engine's classes exposed
$modules = @()
foreach ($file in Get-ChildItem -Path (Join-Path $root 'src') -Filter '*.cls') {
    $text = [IO.File]::ReadAllText($file.FullName, $ansi).Replace("`r`n", "`n").Replace("`n", "`r`n")
    $text = $text.Replace('Attribute VB_Exposed = False', 'Attribute VB_Exposed = True')
    $path = Join-Path $staging $file.Name
    [IO.File]::WriteAllText($path, $text, $ansi)
    $modules += $path
}
$basFiles = @(Get-ChildItem -Path (Join-Path $root 'excel/addin') -Filter '*.bas' | ForEach-Object { $_.FullName })
$basFiles += (Join-Path $root 'examples/use-cases/regex/RegexFunctions.bas')
foreach ($file in $basFiles) {
    $text = [IO.File]::ReadAllText($file, $ansi).Replace("`r`n", "`n").Replace("`n", "`r`n")
    $path = Join-Path $staging ([IO.Path]::GetFileName($file))
    [IO.File]::WriteAllText($path, $text, $ansi)
    $modules += $path
}

$securityKey = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
$previousTrust = Get-ItemProperty -Path $securityKey -Name AccessVBOM -ErrorAction SilentlyContinue
$excel = $null
$book = $null
try {
    if ($TemporaryTrustAccess) {
        New-Item -Path $securityKey -Force | Out-Null
        Set-ItemProperty -Path $securityKey -Name AccessVBOM -Value 1 -Type DWord
    }
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add()
    if ($null -eq $book.VBProject) { throw 'Excel must trust access to the VBA project object model (or run with -TemporaryTrustAccess).' }
    $book.VBProject.Name = 'ASFAddIn'
    $book.VBProject.Description = 'ASF: scripts in worksheet functions (ASF_RUN, ASF_CALL) and the ASF engine'
    foreach ($path in $modules) {
        $imported = $book.VBProject.VBComponents.Import($path)
        Write-Output "Imported $($imported.Name): $($imported.CodeModule.CountOfLines) lines."
    }
    $excel.Run("'" + $book.Name + "'!ASFRegister.RegisterFunctions")
    $book.Title = 'ASF'
    $book.Subject = 'ASF scripts in Excel: ASF_RUN, ASF_CALL and the RX_ regex functions'
    if (Test-Path $Output) { Remove-Item $Output }
    $book.SaveAs($Output, 55)   # xlOpenXMLAddIn
    Write-Output "Saved $Output"
} finally {
    try {
        if ($null -ne $book) { $book.Close($false) }
        if ($null -ne $excel) { $excel.Quit() }
        foreach ($obj in @($book, $excel)) {
            if ($null -ne $obj -and [Runtime.InteropServices.Marshal]::IsComObject($obj)) {
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($obj)
            }
        }
    } finally {
        if ($TemporaryTrustAccess) {
            if ($null -eq $previousTrust) {
                Remove-ItemProperty -Path $securityKey -Name AccessVBOM -ErrorAction SilentlyContinue
            } else {
                Set-ItemProperty -Path $securityKey -Name AccessVBOM -Value $previousTrust.AccessVBOM -Type DWord
            }
        }
        Remove-Item -Recurse -Force $staging -ErrorAction SilentlyContinue
    }
}
