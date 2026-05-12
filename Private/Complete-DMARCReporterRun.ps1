function Complete-DMARCReporterRun {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ManifestPath,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Completed', 'Failed')]
        [string]$Status,

        [Parameter()]
        [object[]]$Stages = @(),

        [Parameter()]
        [object[]]$Errors = @()
    )

    $manifest = Get-Content -LiteralPath $ManifestPath -Raw | ConvertFrom-Json
    $manifest.Status = $Status
    $manifest.CompletedAt = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    $manifest.Stages = @($Stages)
    $manifest.Errors = @($Errors)
    $manifest | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $ManifestPath -Encoding UTF8

    return $manifest
}
