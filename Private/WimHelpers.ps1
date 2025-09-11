function Get-PSN-WimIndexInfo {
    param([Parameter(Mandatory)][string]$WimPath)
    Get-WindowsImage -ImagePath $WimPath |
    Select-Object ImageIndex, ImageName, ImageDescription, Architecture, EditionId, InstallationType
}

function Get-PSN-ArchString {
    param([int]$Architecture)
    switch ($Architecture) { 0 { 'x86' } 9 { 'x64' } 12 { 'arm64' } default { 'x64' } }
}

function New-PSN-TempMount {
    param([string]$MountRoot)
    $p = Join-Path $MountRoot ([guid]::NewGuid().Guid)
    New-Item -ItemType Directory -Path $p -Force | Out-Null
    $p
}

function Get-PSN-UpdateFiles {
    param([Parameter(Mandatory)][string]$UpdatesRoot)
    Get-ChildItem -Path $UpdatesRoot -Filter *.msu -Recurse -File | Sort-Object FullName
}

function Get-PSN-MountedBuildInfo {
    param([Parameter(Mandatory)][string]$MountPath)
    $soft = Join-Path $MountPath 'Windows\System32\config\SOFTWARE'
    if (-not (Test-Path $soft)) { return $null }
    reg.exe load HKLM\PSN_SOFT $soft | Out-Null
    try {
        $cv = 'HKLM:\PSN_SOFT\Microsoft\Windows NT\CurrentVersion'
        $p = Get-ItemProperty $cv -ErrorAction Stop
        $build = if ($p.UBR) { "$($p.CurrentBuild).$($p.UBR)" } else { $p.CurrentBuild }

        # Arch via BuildLabEx
        $archText = 'x86'
        if ($p.BuildLabEx) {
            $lab = $p.BuildLabEx.ToLower()
            if ($lab -match 'arm64') { $archText = 'arm64' }
            elseif ($lab -match 'amd64') { $archText = 'x64' }
            else { $archText = 'x86' }
        }

        [pscustomobject]@{
            ProductName = $p.ProductName
            EditionID   = $p.EditionID
            Build       = $build
            ArchText    = $archText
        }
    }
    catch { $null } finally { reg.exe unload HKLM\PSN_SOFT | Out-Null }
}

function Get-PSN-ShortOsToken {
    param([string]$ProductName, [string]$EditionId, [ValidateSet('x86', 'x64', 'arm64', 'arm')][string]$Arch)
    $os = if ($ProductName -match 'Windows 11') { 'W11' }
    elseif ($ProductName -match 'Windows 10') { 'W10' }
    else { 'WIN' }
    if ([string]::IsNullOrWhiteSpace($EditionId)) { $EditionId = 'Unknown' }
    $ed = $EditionId -replace 'Professional', 'Pro' -replace 'Enterprise', 'Ent' -replace 'Education', 'Edu'
    "$os`_${ed}`_${Arch}"
}

function New-PSN-OutputName {
    param([string]$ProductToken, [string]$Build, [string]$Ext = '.wim')
    "${ProductToken}_$Build_$(Get-Date -Format 'yyyyddMM')$Ext"
}

# --- Architectuur uit DISM.exe (werkt op oude omgevingen; cijfers óf tekst) ---
function Get-PSN-ArchFromWim {
    param([Parameter(Mandatory)][string]$WimPath, [int]$Index = 1)
    $out = & dism.exe /English /Get-WimInfo /WimFile:$WimPath /Index:$Index 2>&1
    $m = $out | Select-String -Pattern 'Architecture\s*:\s*(\S+)' | Select-Object -First 1
    if (-not $m) { return $null }
    $val = $m.Matches[0].Groups[1].Value.Trim().ToLower()
    switch ($val) {
        '0' { 'x86' }
        '9' { 'x64' }
        '12' { 'arm64' }
        'x86' { 'x86' }
        'x64' { 'x64' }
        'amd64' { 'x64' }
        'arm64' { 'arm64' }
        default { $null }
    }
}

function Get-PSN-IndexArchFromDism {
    param([Parameter(Mandatory)][string]$WimPath,
        [Parameter(Mandatory)][int]$Index)
    Get-PSN-ArchFromWim -WimPath $WimPath -Index $Index
}

# --- Servicing: wachten tot DISM/TrustedInstaller klaar is ---
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

# --- Robuuste dismount met retry + /Cleanup-Mountpoints fallback ---
function Invoke-PSN-RetryDismount {
    param([Parameter(Mandatory)][string]$MountDir, [int]$MaxTries = 3)

    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    for ($i = 1; $i -le $MaxTries; $i++) {
        try {
            Write-Host ("Poging {0}/{1}: Dismount-WindowsImage -Save..." -f $i, $MaxTries)
            Dismount-WindowsImage -Path $MountDir -Save -ErrorAction Stop
            return $true
        }
        catch {
            Write-Warning $_.Exception.Message
            Write-Host "Fallback: dism.exe /Unmount-Image /MountDir:$MountDir /Commit"
            Start-Process -FilePath dism.exe -ArgumentList "/Unmount-Image", "/MountDir:$MountDir", "/Commit" -Wait -NoNewWindow
            Start-Sleep -Seconds (5 * $i)

            $stillMounted = $false
            try {
                $m = Get-WindowsImage -Mounted | Where-Object { $_.MountPath -ieq $MountDir -and $_.MountStatus -eq 'Mounted' }
                if ($m) { $stillMounted = $true }
            }
            catch { $stillMounted = Test-Path $MountDir }

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
