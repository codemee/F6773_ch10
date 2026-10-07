param(
    [Parameter(Mandatory = $true)][string]$InputFolder,
    [Parameter(Mandatory = $true)][string]$OutputFolder,
    [string]$DebugLog = ''
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$files = @(Get-ChildItem -LiteralPath $InputFolder -Filter '*.docx' -File | Where-Object { -not $_.Name.StartsWith('~$') } | Sort-Object Name)
$results = @()
$word = $null

function Write-DebugLog([string]$Message) {
    if ($DebugLog) { Add-Content -LiteralPath $DebugLog -Value ((Get-Date -Format o) + ' ' + $Message) -Encoding utf8 }
}

try {
    Write-DebugLog 'Creating Word COM application.'
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    foreach ($file in $files) {
        $document = $null
        $pdfPath = Join-Path $OutputFolder ($file.BaseName + '.pdf')
        try {
            Write-DebugLog ('Opening ' + $file.Name)
            # Open read-only and export; the source DOCX is never saved or altered.
            $document = $word.Documents.Open($file.FullName, $false, $true)
            $document.ExportAsFixedFormat($pdfPath, 17)
            Write-DebugLog ('Exported ' + $file.Name)
            $results += [PSCustomObject]@{ source = $file.Name; output = [IO.Path]::GetFileName($pdfPath); status = 'success'; error = '' }
        }
        catch {
            $results += [PSCustomObject]@{ source = $file.Name; output = [IO.Path]::GetFileName($pdfPath); status = 'failed'; error = $_.Exception.Message }
        }
        finally {
            if ($null -ne $document) { $document.Close($false) | Out-Null }
        }
    }
}
catch {
    Write-DebugLog ('Startup error: ' + $_.Exception.Message)
    # Return a per-file result so the Python command can write its normal report.
    # COM startup can fail on machines where Word has not completed first-run setup.
    foreach ($file in $files) {
        $results += [PSCustomObject]@{ source = $file.Name; output = ($file.BaseName + '.pdf'); status = 'failed'; error = "無法啟動 Microsoft Word：$($_.Exception.Message)" }
    }
}
finally {
    if ($null -ne $word) {
        $word.Quit() | Out-Null
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
        Write-DebugLog 'Closed Word COM application.'
    }
}

@($results) | ConvertTo-Json -Compress
