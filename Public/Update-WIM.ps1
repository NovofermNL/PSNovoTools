function Update-WIM {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$Root='C:\PSNovoTools',
        [string]$WimPath,
        [int]$Index,
        [string]$UpdatesPath,
        [switch]$NoGrid,
        [switch]$PassThru
    )

    Test-PSN-Admin
    $paths = New-PSN-EnsureFolders -Root $Root
    $null  = New-PSN-Transcript -LogsDir $paths.Logs -Prefix 'Update-WIM'

    try {
        if (-not $WimPath) {
            if ($NoGrid) {
                $WimPath = Get-ChildItem -Path $paths.WIM -Filter *.wim -File |
                           Select-Object -First 1 -ExpandProperty FullName
                if (-not $WimPath) { throw "Geen .wim gevonden in $($paths.WIM)" }
            } else {
                $sel = Get-ChildItem -Path $paths.WIM -Filter *.wim -File |
                       Sort-Object LastWriteTime -Descending |
                       Out-GridView -Title "Kies WIM" -PassThru
                if (-not $sel) { throw "Geen WIM geselecteerd." }
                $WimPath = $sel.FullName
            }
        }
        if (-not (Test-Path $WimPath)) { throw "WIM niet gevonden: $WimPath" }

        $idxInfo = Get-PSN-WimIndexInfo -WimPath $WimPath
        if (-not $idxInfo) { throw "Kon geen indexinformatie lezen uit: $WimPath" }

        if (-not $Index) {
            if ($NoGrid) {
                $Index = ($idxInfo | Select-Object -First 1).ImageIndex
            } else {
                $ix = $idxInfo | Out-GridView -Title "Kies WIM-index" -PassThru
                if (-not $ix) { throw "Geen index geselecteerd." }
                $Index = $ix.ImageIndex
            }
        }
        $chosen = $idxInfo | Where-Object { $_.ImageIndex -eq $Index }
        if (-not $chosen) { throw "Index $Index niet gevonden in $WimPath" }
        $arch = Get-PSN-ArchString -Architecture $chosen.Architecture

        if (-not $UpdatesPath) { $UpdatesPath = $paths.Updates }
        if (-not (Test-Path $UpdatesPath)) { throw "Updates-pad niet gevonden: $UpdatesPath" }
        $updates = Get-PSN-UpdateFiles -UpdatesRoot $UpdatesPath
        if (-not $updates) { throw "Geen .msu-updates gevonden in $UpdatesPath (recursief)." }

        if ($NoGrid) { $chosenUpdates = $updates }
        else { $chosenUpdates = $updates | Out-GridView -Title "Selecteer updates (.msu) - multi-select" -PassThru }
        if (-not $chosenUpdates) { throw "Geen updates geselecteerd." }

        $mount = New-PSN-TempMount -MountRoot $paths.Mount
        Mount-WindowsImage -ImagePath $WimPath -Index $Index -Path $mount -ErrorAction Stop | Out-Null
        $mounted = $true

        try {
            foreach ($u in $chosenUpdates) {
                Add-WindowsPackage -Path $mount -PackagePath $u.FullName -IgnoreCheck -PreventPending -ErrorAction Stop | Out-Null
            }

            $bi    = Get-PSN-MountedBuildInfo -MountPath $mount
            if (-not $bi) { $bi = [pscustomobject]@{ ProductName=$chosen.ImageName; EditionID=$chosen.EditionId; Build='UnknownBuild' } }
            $token = Get-PSN-ShortOsToken -ProductName $bi.ProductName -EditionId $bi.EditionID -Arch $arch
            $build = $bi.Build
            $name  = New-PSN-OutputName -ProductToken $token -Build $build
            $out   = Join-Path $paths.Output $name

            Dismount-WindowsImage -Path $mount -Save -ErrorAction Stop | Out-Null
            $mounted = $false

            Export-WindowsImage -SourceImagePath $WimPath -SourceIndex $Index `
                -DestinationImagePath $out -CheckIntegrity -CompressionType Max -ErrorAction Stop | Out-Null

            Write-Host ("Nieuwe WIM: {0}" -f $out)
            if ($PassThru) {
                [pscustomobject]@{
                    OutputPath=$out; ProductToken=$token; Build=$build
                    DateStamp=(Get-Date -Format 'yyyyddMM'); Index=$Index
                    SourceWim=$WimPath; Updates=$chosenUpdates.FullName
                }
            }
        }
        finally {
            if ($mounted) {
                try { Dismount-WindowsImage -Path $mount -Discard -ErrorAction SilentlyContinue | Out-Null } catch {}
            }
            if (Test-Path $mount) { try { Remove-Item -Recurse -Force $mount } catch {} }
        }
    }
    finally { Stop-PSN-Transcript }
}
Export-ModuleMember -Function Update-WIM
