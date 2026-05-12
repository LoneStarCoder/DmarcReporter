$script:ModuleRoot = $PSScriptRoot

foreach ($folder in @('Private', 'Public')) {
    $path = Join-Path -Path $PSScriptRoot -ChildPath $folder
    if (-not (Test-Path -LiteralPath $path -PathType Container)) {
        continue
    }

    Get-ChildItem -LiteralPath $path -Filter '*.ps1' -File |
        Sort-Object -Property Name |
        ForEach-Object {
            . $_.FullName
        }
}

Export-ModuleMember -Function @(
    'Invoke-DmarcDashboard',
    'New-DmarcDashboardConfig'
)
