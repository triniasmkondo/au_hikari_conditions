param([Parameter(Mandatory=$true)][string]$PlanPath)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$plan = Get-Content -LiteralPath $PlanPath -Raw -Encoding UTF8 | ConvertFrom-Json
$resultPath = Join-Path (Split-Path -Parent $PlanPath) 'conversion_results.json'
$results = @()
if (Test-Path -LiteralPath $resultPath) { $loaded = Get-Content -LiteralPath $resultPath -Raw -Encoding UTF8 | ConvertFrom-Json; foreach ($entry in $loaded) { $results += $entry } }
$word = $null
try {
    foreach ($item in $plan.documents) {
        $sourceHash = (Get-FileHash -LiteralPath $item.source -Algorithm SHA256).Hash.ToLowerInvariant()
        if ($sourceHash -ne $item.source_sha256) { throw 'LOCK source fingerprint mismatch' }
        $existing = @($results | Where-Object { $_.source -eq $item.source })
        if (Test-Path -LiteralPath $item.destination) {
            if ($existing.Count -ne 1) { throw 'Existing PDF has no conversion receipt' }
            $pdfHash = (Get-FileHash -LiteralPath $item.destination -Algorithm SHA256).Hash.ToLowerInvariant()
            if ($existing[0].pdf_sha256 -ne $pdfHash -or $existing[0].source_sha256 -ne $sourceHash) { throw 'Existing PDF fingerprint mismatch' }
            continue
        }
        [void][System.IO.Directory]::CreateDirectory((Split-Path -Parent $item.destination))
        if ($null -eq $word) {
            $word = New-Object -ComObject Word.Application
            $word.Visible = $false
            $word.DisplayAlerts = 0
            $word.AutomationSecurity = 3
        }
        $document = $null
        try {
            $document = $word.Documents.Open([string]$item.source, $false, $true, $false)
            $document.ExportAsFixedFormat([string]$item.destination, 17)
        } finally {
            if ($null -ne $document) { $document.Close(0); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($document) }
        }
        if ((Get-FileHash -LiteralPath $item.source -Algorithm SHA256).Hash.ToLowerInvariant() -ne $sourceHash) { throw 'LOCK source changed during conversion' }
        $results += [PSCustomObject]@{
            agency=$item.agency; role=$item.role; source=$item.source; destination=$item.destination;
            source_sha256=$sourceHash; pdf_sha256=(Get-FileHash -LiteralPath $item.destination -Algorithm SHA256).Hash.ToLowerInvariant();
            converter='Microsoft Word ExportAsFixedFormat'; converted_at=[DateTimeOffset]::UtcNow.ToString('o')
        }
        $results | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $resultPath -Encoding UTF8
        Write-Output ('PDF_COMPLETE ' + $results.Count + '/' + $plan.documents.Count)
    }
} finally {
    if ($null -ne $word) { $saveChanges = 0; $word.Quit([ref]$saveChanges); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
}
