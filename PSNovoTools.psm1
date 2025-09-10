# PSNovoTools.psm1 - loader
$privatePath = Join-Path $PSScriptRoot 'Private'
$publicPath  = Join-Path $PSScriptRoot 'Public'

if (Test-Path $privatePath) {
    Get-ChildItem -Path $privatePath -Filter *.ps1 -File | ForEach-Object { . $_.FullName }
}
if (Test-Path $publicPath) {
    Get-ChildItem -Path $publicPath  -Filter *.ps1 -File | ForEach-Object { . $_.FullName }
}

# Fallback export (voor het geval Public-scripts Export-ModuleMember missen)
Export-ModuleMember -Function Import-PSNovoTools, Update-WIM -ErrorAction SilentlyContinue
