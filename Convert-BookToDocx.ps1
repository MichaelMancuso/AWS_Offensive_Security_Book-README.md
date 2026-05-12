# Convert-BookToDocx.ps1
# Converts AWS_Offensive_Security.html to AWS_Offensive_Security.docx
# using Microsoft Word's own HTML importer (requires Word to be installed).
#
# Run:  powershell -ExecutionPolicy Bypass -File .\Convert-BookToDocx.ps1
#
# The script searches for the HTML in (1) the folder it lives in, then
# (2) both Cowork data roots (classic and MSIX / packaged app).

$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$fileName  = 'AWS_Offensive_Security.html'

$candidates = @(
    (Join-Path $scriptDir $fileName),
    'C:\Users\M2\AppData\Roaming\Claude\local-agent-mode-sessions\38a063fa-b67a-4d5f-b927-0c4121abcf17\b4fb3676-2204-418d-a217-3d3c163ce4a8\local_a9e42691-adaa-44ca-8428-df1ce4fdd8d8\outputs\AWS_Offensive_Security.html'
)

$htmlPath = $null
foreach ($c in $candidates) {
    if (Test-Path $c) { $htmlPath = $c; break }
}

if (-not $htmlPath) {
    Write-Host "Searching Cowork session folders for $fileName..."

    $searchRoots = @(
        "$env:AppData\Claude\local-agent-mode-sessions"
    )
    $pkgParent = "$env:LocalAppData\Packages"
    if (Test-Path $pkgParent) {
        Get-ChildItem -Path $pkgParent -Directory -Filter 'Claude_*' -ErrorAction SilentlyContinue |
            ForEach-Object {
                $searchRoots += (Join-Path $_.FullName 'LocalCache\Roaming\Claude\local-agent-mode-sessions')
            }
    }

    foreach ($root in $searchRoots) {
        if (-not (Test-Path $root)) { continue }
        $hit = Get-ChildItem -Path $root -Recurse -Filter $fileName -ErrorAction SilentlyContinue |
               Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($hit) { $htmlPath = $hit.FullName; break }
    }
}

if (-not $htmlPath) {
    Write-Error "Could not find $fileName. Place it next to this script and re-run."
    exit 1
}

$docxPath = Join-Path $scriptDir 'AWS_Offensive_Security.docx'

Write-Host "Source: $htmlPath"
Write-Host "Target: $docxPath"
Write-Host "Launching Microsoft Word (will run hidden)..."

$word = $null
try {
    $word = New-Object -ComObject Word.Application
} catch {
    Write-Error "Failed to start Word COM. Is Microsoft Word installed? ($_)"
    exit 1
}
if (-not $word) {
    Write-Error "Word COM object is null. Word may not be installed (Office Store/UWP installs do NOT expose COM)."
    exit 1
}
$word.Visible = $false
$word.DisplayAlerts = 0

try {
    Write-Host "Opening HTML. For a 4,000+ line book this takes 1-2 minutes..."
    # Minimal 3-arg overload: path, ConfirmConversions=false, ReadOnly=true.
    # The 15-arg form crashes PowerShell's COM bridge with NullReferenceException.
    $doc = $word.Documents.Open($htmlPath, $false, $true)
    if (-not $doc) {
        throw "Documents.Open returned null for $htmlPath"
    }

    Write-Host "Saving as .docx..."
    # SaveAs via COM is picky in PowerShell 7 — strings get wrapped as PSObject
    # and the COM bridge can't unbox them through [ref]. Unwrap to native types
    # and use SaveAs2, which accepts positional non-ref args.
    $pathStr = [string]$docxPath
    [int]$wdFormatDocumentDefault = 16
    try {
        $doc.SaveAs2($pathStr, $wdFormatDocumentDefault)
    } catch {
        # Fallback: older Word versions lack SaveAs2
        $doc.SaveAs($pathStr, $wdFormatDocumentDefault)
    }
    $doc.Close($false)

    if (Test-Path $docxPath) {
        $size = (Get-Item $docxPath).Length
        Write-Host ("Done. Output: {0} ({1:N0} bytes)" -f $docxPath, $size)
    } else {
        Write-Error "Conversion finished but output file not found."
    }
}
finally {
    if ($word) {
        try { $word.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host ""
Write-Host "You can now open AWS_Offensive_Security.docx directly."
