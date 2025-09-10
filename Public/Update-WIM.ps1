function Update-WIM {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$Root='C:\PSNovoTools',
        [string]$WimPath,
        [int]$Index,
        [string]$UpdatesPath,
        [switch]$NoGrid,
        [switch]$PassThru,
        [switch]$SkipCleanup         # sla DISM component cleanup over
    )

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Test-PSN-Admin

    $paths = New-PSN-EnsureFolders -Root $Root
    $null  = New-PSN-Transcript -LogsDir $paths.Logs -Prefix 'Update-WIM'

    # Scratch dir
    $scratchDir = Join-Path $env:WINDIR 'Temp\DISM_Scratch'
    if (-not (Test-Path $scratchDir)) { New-Item -ItemType Directory -Path $scratchDir -Force | Out-Null }

    $wsearchWasRunning = $false
    $defenderExclAdded = $false
    $mounted = $false
    $mount   = $null

    $totalTimer = [System.Diagnostics.Stopwatch]::StartNew()

    try {
        # ===== 1) WIM kiezen =====
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

        # ===== 2) Index kiezen =====
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

        # ===== 3) Updates selecteren =====
        if (-not $UpdatesPath) { $UpdatesPath = $paths.Updates }
        if (-not (Test-Path $UpdatesPath)) { throw "Updates-pad niet gevonden: $UpdatesPath" }
        $updates = Get-PSN-UpdateFiles -UpdatesRoot $UpdatesPath
        if (-not $updates) { throw "Geen .msu-updates gevonden in $UpdatesPath (recursief)." }

        if ($NoGrid) {
            $chosenUpdates = $updates
        } else {
            $chosenUpdates = $updates | Out-GridView -Title "Selecteer updates (.msu) - multi-select" -PassThru
        }
        if (-not $chosenUpdates) { throw "Geen updates geselecteerd." }

        # Sorteer: eerst CUs (niet-ndp), dan .NET (ndp)
        $cuUpdates  = @($chosenUpdates | Where-Object { $_.Name -notlike "*ndp*" } | Sort-Object Name)
        $ndpUpdates = @($chosenUpdates | Where-Object { $_.Name -like "*ndp*" }   | Sort-Object Name)
        $ordered    = @($cuUpdates + $ndpUpdates)

        Write-Host ("Totaal {0} updates (CUs: {1}, .NET: {2})." -f $ordered.Count, $cuUpdates.Count, $ndpUpdates.Count)

        # ===== Voorzorg: services & AV =====
        try {
            $svc = Get-Service WSearch -ErrorAction Stop
            if ($svc.Status -eq 'Running') {
                $wsearchWasRunning = $true
                Write-Host "Windows Search (WSearch) stoppen..."
                Stop-Service WSearch -Force -ErrorAction Stop
            }
        } catch { Write-Host "Kon WSearch niet stoppen: $($_.Exception.Message)" }

        if (Get-Command Add-MpPreference -ErrorAction SilentlyContinue) {
            try {
                Add-MpPreference -ExclusionPath $paths.Mount -ErrorAction Stop
                $defenderExclAdded = $true
                Write-Host "Defender-exclusie voorbereid (map zal bestaan na mount): $($paths.Mount)"
            } catch { Write-Host "Kon Defender-exclusie niet toevoegen: $($_.Exception.Message)" }
        }

        # ===== 4) Mount =====
        $mount = New-PSN-TempMount -MountRoot $paths.Mount
        Write-Host "Mount image: $WimPath (Index $Index) → $mount"
        Mount-WindowsImage -ImagePath $WimPath -Index $Index -Path $mount -CheckIntegrity -ScratchDirectory $scratchDir -ErrorAction Stop | Out-Null
        $mounted = $true

        # Defender-exclusie opnieuw, nu met werkelijk mountpad
        if (-not $defenderExclAdded -and (Get-Command Add-MpPreference -ErrorAction SilentlyContinue)) {
            try { Add-MpPreference -ExclusionPath $mount -ErrorAction Stop; $defenderExclAdded = $true } catch {}
        }

        # ===== 5) Inject updates met voortgang =====
        if ($ordered.Count -gt 0) {
            $total = $ordered.Count
            $count = 0
            foreach ($msu in $ordered) {
                $count++
                $percent = [math]::Round(($count / $total) * 100, 2)
                $tick = [System.Diagnostics.Stopwatch]::StartNew()

                Write-Progress -Activity "Integratie van updates" -Status "Bezig met $($msu.Name) ($count van $total)" -PercentComplete $percent
                Write-Host ("[{0}/{1}] Toevoegen: {2}" -f $count,$total,$msu.Name)

                try {
                    Add-WindowsPackage -Path $mount -PackagePath $msu.FullName -IgnoreCheck -PreventPending `
                        -ScratchDirectory $scratchDir -ErrorAction Stop | Out-Null
                } catch {
                    Write-Warning ("Fout bij toevoegen van {0}: {1}" -f $msu.Name, $_.Exception.Message)
                } finally {
                    $tick.Stop()
                    Write-Host ("Duur: {0}" -f $tick.Elapsed.ToString())
                }
            }
            Write-Progress -Activity "Integratie van updates" -Completed
        }

        # ===== 6) Optionele component cleanup =====
        if (-not $SkipCleanup.IsPresent) {
            Write-Host "DISM component cleanup uitvoeren..."
            try {
                Start-Process -FilePath dism.exe -ArgumentList "/Image:$mount","/Cleanup-Image","/StartComponentCleanup","/ScratchDir:$scratchDir" -Wait -NoNewWindow
            } catch { Write-Warning ("Cleanup mislukt: {0}" -f $_.Exception.Message) }
        } else {
            Write-Host "Component cleanup overgeslagen (-SkipCleanup)."
        }

        # Wacht even tot servicing rustig is
        if (-not (Wait-PSN-ServicingIdle -TimeoutSec 180)) {
            Write-Warning "Servicing-proces bleef actief (dism/dismhost/tiworker/trustedinstaller)."
        }

        # ===== 7) Build-info, commit, dismount (retry) =====
        $bi    = Get-PSN-MountedBuildInfo -MountPath $mount
        if (-not $bi) { $bi = [pscustomobject]@{ ProductName=$chosen.ImageName; EditionID=$chosen.EditionId; Build='UnknownBuild' } }
        $token = Get-PSN-ShortOsToken -ProductName $bi.ProductName -EditionId $bi.EditionID -Arch $arch
        $build = $bi.Build
        $name  = New-PSN-OutputName -ProductToken $token -Build $build
        $out   = Join-Path $paths.Output $name

        Write-Host "Committen en dismounten..."
        $ok = Invoke-PSN-RetryDismount -MountDir $mount -MaxTries 3
        $mounted = -not $ok
        if (-not $ok) { throw "Kon mount niet netjes ontkoppelen: $mount" }

        # ===== 8) Export naar nieuwe WIM =====
        Write-Host "Exporteren naar: $out"
        Export-WindowsImage -SourceImagePath $WimPath -SourceIndex $Index `
            -DestinationImagePath $out -CheckIntegrity -CompressionType Max -ErrorAction Stop | Out-Null

        Write-Host ("Nieuwe WIM: {0}" -f $out)
        if ($PassThru) {
            [pscustomobject]@{
                OutputPath=$out; ProductToken=$token; Build=$build
                DateStamp=(Get-Date -Format 'yyyyddMM'); Index=$Index
                SourceWim=$WimPath; Updates=$ordered.FullName
            }
        }
    }
    catch {
        Write-Error $_.Exception.Message
        Write-Host "Incident-handling: cleanup mountpoints & discard indien nodig."
        try {
            Start-Process -FilePath dism.exe -ArgumentList "/Cleanup-Mountpoints" -Wait -NoNewWindow
            if ($mounted -and $mount) { Dismount-WindowsImage -Path $mount -Discard -ErrorAction SilentlyContinue }
        } catch {}
        throw
    }
    finally {
        # Services herstellen
        if ($wsearchWasRunning) {
            try { Start-Service WSearch -ErrorAction Stop } catch { Write-Host "Kon WSearch niet starten: $($_.Exception.Message)" }
        }
        if ($defenderExclAdded -and (Get-Command Remove-MpPreference -ErrorAction SilentlyContinue)) {
            try { Remove-MpPreference -ExclusionPath ($mount ?? $paths.Mount) -ErrorAction Stop } catch {}
        }

        $totalTimer.Stop()
        Write-Host ("Totale duur: {0}" -f $totalTimer.Elapsed.ToString())

        # Opruimen mountmap als die nog bestaat
        if ($mount -and (Test-Path $mount)) {
            try { Remove-Item -Recurse -Force $mount } catch {}
        }
        Stop-PSN-Transcript
    }
}
Export-ModuleMember -Function Update-WIM

