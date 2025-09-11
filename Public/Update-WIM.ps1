function Update-WIM {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$Root = 'C:\PSNovoTools',
        [string]$WimPath,
        [int]$Index,
        [string]$UpdatesPath,
        [switch]$NoGrid,
        [switch]$PassThru,
        [switch]$SkipCleanup
    )

    # Basis
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Test-PSN-Admin

    $paths = New-PSN-EnsureFolders -Root $Root
    $null = New-PSN-Transcript -LogsDir $paths.Logs -Prefix 'Update-WIM'

    # DISM scratch
    $scratchDir = Join-Path $env:WINDIR 'Temp\DISM_Scratch'
    if (-not (Test-Path $scratchDir)) { New-Item -ItemType Directory -Path $scratchDir -Force | Out-Null }

    # State
    $wsearchWasRunning = $false
    $defenderExclAdded = $false
    $mounted = $false
    $mount = $null
    $newToken = $null

    $totalTimer = [System.Diagnostics.Stopwatch]::StartNew()

    try {
        # ===== 1) WIM kiezen =====
        if (-not $WimPath) {
            if ($NoGrid) {
                $WimPath = Get-ChildItem -Path $paths.WIM -Filter *.wim -File |
                Select-Object -First 1 -ExpandProperty FullName
                if (-not $WimPath) { throw "Geen .wim gevonden in $($paths.WIM)" }
            }
            else {
                $sel = Get-ChildItem -Path $paths.WIM -Filter *.wim -File |
                Sort-Object LastWriteTime -Descending |
                Out-GridView -Title "Kies WIM" -PassThru
                if (-not $sel) { throw "Geen WIM geselecteerd." }
                $WimPath = $sel.FullName
            }
        }
        if (-not (Test-Path $WimPath)) { throw "WIM niet gevonden: $WimPath" }

        # ===== 2) Index kiezen (met betrouwbare Arch uit DISM) =====
        $idxInfo = Get-PSN-WimIndexInfo -WimPath $WimPath

        $idxDisplay = foreach ($i in $idxInfo) {
            # Altijd eerst via DISM (levert tekst of cijfers terug, beide afgevangen)
            $arch = Get-PSN-IndexArchFromDism -WimPath $WimPath -Index $i.ImageIndex
            if (-not $arch -and $i.PSObject.Properties.Name -contains 'Architecture' -and $null -ne $i.Architecture) {
                $arch = Get-PSN-ArchString -Architecture $i.Architecture
            }
            [pscustomobject]@{
                ImageIndex = $i.ImageIndex
                ImageName  = $i.ImageName
                EditionId  = $i.EditionId
                Arch       = $arch
            }
        }

        if (-not $Index) {
            if ($NoGrid) {
                $Index = ($idxDisplay | Select-Object -First 1).ImageIndex
            }
            else {
                $ix = $idxDisplay | Out-GridView -Title "Kies WIM-index" -PassThru
                if (-not $ix) { throw "Geen index geselecteerd." }
                $Index = $ix.ImageIndex
            }
        }
        $chosen = $idxInfo    | Where-Object { $_.ImageIndex -eq $Index }
        $archText = $idxDisplay | Where-Object { $_.ImageIndex -eq $Index } |
        Select-Object -First 1 -ExpandProperty Arch
        if (-not $chosen) { throw "Index $Index niet gevonden in $WimPath" }

        # ===== 3) Updates selecteren =====
        if (-not $UpdatesPath) { $UpdatesPath = $paths.Updates }
        if (-not (Test-Path $UpdatesPath)) { throw "Updates-pad niet gevonden: $UpdatesPath" }
        $updates = Get-PSN-UpdateFiles -UpdatesRoot $UpdatesPath
        if (-not $updates) { throw "Geen .msu-updates gevonden in $UpdatesPath (recursief)." }

        if ($NoGrid) { $chosenUpdates = $updates }
        else { $chosenUpdates = $updates | Out-GridView -Title "Selecteer updates (.msu) - multi-select" -PassThru }
        if (-not $chosenUpdates) { throw "Geen updates geselecteerd." }

        # Sorteer: CUs eerst, dan .NET (ndp)
        $cu = @($chosenUpdates | Where-Object { $_.Name -notlike "*ndp*" } | Sort-Object Name)
        $ndp = @($chosenUpdates | Where-Object { $_.Name -like "*ndp*" }   | Sort-Object Name)
        $ordered = @($cu + $ndp)
        Write-Host ("Totaal {0} updates (CUs: {1}, .NET: {2})." -f $ordered.Count, $cu.Count, $ndp.Count)

        # ===== Voorzorg: services & AV =====
        try {
            $svc = Get-Service WSearch -ErrorAction Stop
            if ($svc.Status -eq 'Running') {
                $wsearchWasRunning = $true
                Write-Verbose "Windows Search (WSearch) stoppen..."
                Stop-Service WSearch -Force -ErrorAction Stop
            }
        }
        catch {
            Write-Verbose ("Kon WSearch niet stoppen: {0}" -f $_.Exception.Message)
        }

        # ===== 4) Mount =====
        $mount = New-PSN-TempMount -MountRoot $paths.Mount
        Write-Host "Mount image: $WimPath (Index $Index) -> $mount"
        Mount-WindowsImage -ImagePath $WimPath -Index $Index -Path $mount -CheckIntegrity -ScratchDirectory $scratchDir -ErrorAction Stop | Out-Null
        $mounted = $true

        # Defender-exclusie met échte mountmap (als de eerste poging faalde)
        if (-not $defenderExclAdded -and (Get-Command Add-MpPreference -ErrorAction SilentlyContinue)) {
            try { Add-MpPreference -ExclusionPath $mount -ErrorAction Stop; $defenderExclAdded = $true } catch { }
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
                Write-Host ("[{0}/{1}] Toevoegen: {2}" -f $count, $total, $msu.Name)

                try {
                    Add-WindowsPackage -Path $mount -PackagePath $msu.FullName -IgnoreCheck -PreventPending `
                        -ScratchDirectory $scratchDir -ErrorAction Stop | Out-Null
                }
                catch {
                    Write-Warning ("Fout bij toevoegen van {0}: {1}" -f $msu.Name, $_.Exception.Message)
                }
                finally {
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
                Start-Process -FilePath dism.exe -ArgumentList "/Image:$mount", "/Cleanup-Image", "/StartComponentCleanup", "/ScratchDir:$scratchDir" -Wait -NoNewWindow
            }
            catch { Write-Warning ("Cleanup mislukt: {0}" -f $_.Exception.Message) }
        }
        else {
            Write-Host "Component cleanup overgeslagen (-SkipCleanup)."
        }

        # ===== 7) Build/Arch uit offline registry (of fallback naar GridView-Arch) =====
        $bi = Get-PSN-MountedBuildInfo -MountPath $mount
        $archFromReg = $null
        if ($bi -and $bi.PSObject.Properties.Name -contains 'ArchText' -and $bi.ArchText) { $archFromReg = $bi.ArchText }
        if (-not $archFromReg) { $archFromReg = $archText }

        if (-not $bi) {
            $bi = [pscustomobject]@{
                ProductName = $chosen.ImageName
                EditionID   = $chosen.EditionId
                Build       = 'UnknownBuild'
            }
        }

        $token = Get-PSN-ShortOsToken -ProductName $bi.ProductName -EditionId $bi.EditionID -Arch $archFromReg
        $build = $bi.Build
        $name = New-PSN-OutputName -ProductToken $token -Build $build
        $out = Join-Path $paths.Output $name

        # ===== 8) Commit + dismount (met retry helper) =====
        Write-Host "Committen en dismounten..."
        $ok = Invoke-PSN-RetryDismount -MountDir $mount -MaxTries 3
        $mounted = -not $ok
        if (-not $ok) { throw "Kon mount niet netjes ontkoppelen: $mount" }

        # ===== 9) Export naar nieuwe WIM =====
        Write-Host "Exporteren naar: $out"
        Export-WindowsImage -SourceImagePath $WimPath -SourceIndex $Index `
            -DestinationImagePath $out -CheckIntegrity -CompressionType Max -ErrorAction Stop | Out-Null

        # ===== 10) Safety: verifieer architectuur van output en hernoem indien nodig =====
        try {
            $archVerify = $null
            try {
                $ver = Get-WindowsImage -ImagePath $out | Select-Object -First 1
                if ($ver -and $ver.PSObject.Properties.Name -contains 'Architecture' -and $null -ne $ver.Architecture) {
                    $archVerify = Get-PSN-ArchString -Architecture $ver.Architecture
                }
            }
            catch { }
            if (-not $archVerify) {
                $archVerify = Get-PSN-ArchFromWim -WimPath $out -Index 1
            }

            if ($archVerify -and $archVerify -ne $archFromReg) {
                $newToken = ($token -replace '_(x86|x64|arm64)$', ("_{0}" -f $archVerify))
                $newName = New-PSN-OutputName -ProductToken $newToken -Build $build
                $newPath = Join-Path $paths.Output $newName
                Rename-Item -Path $out -NewName $newName
                $out = $newPath
                Write-Host ("Bestandsnaam gecorrigeerd naar: {0}" -f $out)
            }
        }
        catch { }

        Write-Host ("Nieuwe WIM: {0}" -f $out)

        if ($PassThru) {
            $productTokenOut = $token
            if ($null -ne $newToken -and $newToken -ne '') { $productTokenOut = $newToken }
            [pscustomobject]@{
                OutputPath   = $out
                ProductToken = $productTokenOut
                Build        = $build
                DateStamp    = (Get-Date -Format 'yyyyddMM')
                Index        = $Index
                SourceWim    = $WimPath
            }
        }
    }
    catch {
        Write-Error $_.Exception.Message
        try {
            Start-Process -FilePath dism.exe -ArgumentList "/Cleanup-Mountpoints" -Wait -NoNewWindow | Out-Null
            if ($mounted -and $mount) {
                Dismount-WindowsImage -Path $mount -Discard -ErrorAction SilentlyContinue | Out-Null
            }
        }
        catch { }
        throw
    }
    finally {
        if ($wsearchWasRunning) {
            try { Start-Service WSearch -ErrorAction Stop | Out-Null } catch { Write-Verbose ("Kon WSearch niet starten: {0}" -f $_.Exception.Message) }
        }
        if ($defenderExclAdded -and (Get-Command Remove-MpPreference -ErrorAction SilentlyContinue)) {
            $exclPath = $paths.Mount
            if ($mount) { $exclPath = $mount }
            try { Remove-MpPreference -ExclusionPath $exclPath -ErrorAction Stop } catch { }
        }
        $totalTimer.Stop()
        Write-Host ("Totale duur: {0}" -f $totalTimer.Elapsed.ToString())
        if ($mount -and (Test-Path $mount)) {
            try { Remove-Item -Recurse -Force $mount } catch { }
        }
        Stop-PSN-Transcript
    }
}
Export-ModuleMember -Function Update-WIM
