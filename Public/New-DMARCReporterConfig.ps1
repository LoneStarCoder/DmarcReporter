function New-DMARCReporterConfig {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter()]
        [string]$Path = '.\config\dmarc.config.json',

        [Parameter()]
        [string]$MailboxFolder = 'Inbox\Ignore\dmarcreports',

        [Parameter()]
        [string]$OutputRoot = '.\Runs',

        [Parameter()]
        [switch]$Force
    )

    $resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    $parent = Split-Path -Path $resolvedPath -Parent
    if (-not [string]::IsNullOrWhiteSpace($parent) -and -not (Test-Path -LiteralPath $parent -PathType Container)) {
        New-Item -Path $parent -ItemType Directory -Force | Out-Null
    }

    if ((Test-Path -LiteralPath $resolvedPath -PathType Leaf) -and -not $Force) {
        throw "Config file already exists: $resolvedPath. Use -Force to overwrite."
    }

    $config = Get-DMARCReporterDefaultConfig
    $config.MailboxFolder = $MailboxFolder
    $config.OutputRoot = $OutputRoot

    if ($PSCmdlet.ShouldProcess($resolvedPath, 'Create DMARCReporter config')) {
        $config | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $resolvedPath -Encoding UTF8
    }

    return Get-Item -LiteralPath $resolvedPath
}
