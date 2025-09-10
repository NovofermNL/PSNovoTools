function Import-PSNovoTools {
    [CmdletBinding()]
    param([string]$Root = 'C:\PSNovoTools')
    Test-PSN-Admin
    $paths = New-PSN-EnsureFolders -Root $Root
    "PSNovoTools klaar. Root: $($paths.Root)" | Write-Output
}
Export-ModuleMember -Function Import-PSNovoTools
