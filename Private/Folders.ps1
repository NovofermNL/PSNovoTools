function New-PSN-EnsureFolders {
    param([string]$Root = 'C:\PSNovoTools')

    $folders = @(
        $Root
        (Join-Path $Root 'WIMFiles')
        (Join-Path $Root 'Updates')
        (Join-Path $Root 'Mount')
        (Join-Path $Root 'Output')
        (Join-Path $Root 'Logs')
    )

    foreach ($f in $folders) {
        if (-not (Test-Path $f)) {
            New-Item -ItemType Directory -Path $f -Force | Out-Null
        }
    }

    [pscustomobject]@{
        Root    = $Root
        WIM     = (Join-Path $Root 'WIMFiles')
        Updates = (Join-Path $Root 'Updates')
        Mount   = (Join-Path $Root 'Mount')
        Output  = (Join-Path $Root 'Output')
        Logs    = (Join-Path $Root 'Logs')
    }
}
