function Wait-PSN-ServicingIdle {
    param([int]$TimeoutSec = 180)
    $stopAt = (Get-Date).AddSeconds($TimeoutSec)
    do {
        $p = Get-Process dism, dismhost, tiworker, trustedinstaller -ErrorAction SilentlyContinue
        if (-not $p) { return $true }
        Start-Sleep -Seconds 2
    } while ((Get-Date) -lt $stopAt)
    return $false
}

function Invoke-PSN-RetryDismount {
    param(
        [Parameter(Mandatory)][string]$MountDir,
        [int]$MaxTries = 3
    )

    [GC]::Collect(); [GC]::WaitForPendingFinalizers()

    for ($i=1; $i -le $MaxTries; $i++) {
        try {
            Write-Host ("Poging {0}/{1}: Dismount-WindowsImage -Save..." -f $i,$MaxTries)
            Dismount-WindowsImage -Path $MountDir -Save -ErrorAction Stop
            return $true
        } catch {
            Write-Warning $_.Exception.Message
            Write-Host "Fallback: dism.exe /Unmount-Image /MountDir:$MountDir /Commit"
            Start-Process -FilePath dism.exe -ArgumentList "/Unmount-Image","/MountDir:$MountDir","/Commit" -Wait -NoNewWindow
            Start-Sleep -Seconds (5 * $i)

            $stillMounted = $false
            try {
                $m = Get-WindowsImage -Mounted | Where-Object { $_.MountPath -ieq $MountDir -and $_.MountStatus -eq 'Mounted' }
                if ($m) { $stillMounted = $true }
            } catch { $stillMounted = Test-Path $MountDir }

            if (-not $stillMounted) { return $true }

            if ($i -eq $MaxTries) {
                Write-Host "Forceer cleanup van mountpoints..."
                Start-Process -FilePath dism.exe -ArgumentList "/Cleanup-Mountpoints" -Wait -NoNewWindow
                try { Dismount-WindowsImage -Path $MountDir -Discard -ErrorAction SilentlyContinue } catch {}
                return -not (Test-Path $MountDir)
            }
        }
    }
    return $false
}
