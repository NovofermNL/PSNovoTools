function Get-PSN-WimIndexInfo {
    param([Parameter(Mandatory)][string]$WimPath)
    Get-WindowsImage -ImagePath $WimPath |
        Select-Object ImageIndex,ImageName,ImageDescription,Architecture,EditionId,InstallationType
}
function Get-PSN-ArchString {
    param([int]$Architecture)
    switch ($Architecture) { 0{'x86'} 9{'x64'} 12{'arm64'} default{'x64'} }
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
        $p  = Get-ItemProperty $cv -ErrorAction Stop
        $build = if ($p.UBR) { "$($p.CurrentBuild).$($p.UBR)" } else { $p.CurrentBuild }
        [pscustomobject]@{ ProductName=$p.ProductName; EditionID=$p.EditionID; Build=$build }
    } catch { $null } finally { reg.exe unload HKLM\PSN_SOFT | Out-Null }
}
function Get-PSN-ShortOsToken {
    param([string]$ProductName,[string]$EditionId,[ValidateSet('x86','x64','arm64','arm')][string]$Arch)
    $os = if ($ProductName -match 'Windows 11') {'W11'} elseif ($ProductName -match 'Windows 10') {'W10'} else {'WIN'}
    if ([string]::IsNullOrWhiteSpace($EditionId)) { $EditionId = 'Unknown' }
    $ed = $EditionId -replace 'Professional','Pro' -replace 'Enterprise','Ent' -replace 'Education','Edu'
    "$os`_${ed}`_${Arch}"
}
function New-PSN-OutputName { param([string]$ProductToken,[string]$Build,[string]$Ext='.wim')
    "${ProductToken}_$Build_$(Get-Date -Format 'yyyyddMM')$Ext"
}
