function New-PSN-Transcript {
    param([string]$LogsDir, [string]$Prefix = 'Update-WIM')
    $stamp = Get-Date -Format 'yyyyddMM-HHmmss'
    $file = Join-Path $LogsDir "$Prefix-$stamp.log"
    try { Start-Transcript -Path $file -Append -ErrorAction Stop | Out-Null } catch {}
    $file
}
function Stop-PSN-Transcript { try { Stop-Transcript | Out-Null } catch {} }
