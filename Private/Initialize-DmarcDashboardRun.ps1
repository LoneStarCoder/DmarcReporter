function Initialize-DmarcDashboardRun {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Config
    )

    $outputRoot = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath([string]$Config.OutputRoot)
    if (-not (Test-Path -LiteralPath $outputRoot -PathType Container)) {
        New-Item -Path $outputRoot -ItemType Directory -Force | Out-Null
    }

    $runId = '{0}_{1}' -f (Get-Date -Format 'yyyy-MM-dd_HHmmss'), ([guid]::NewGuid().ToString('N').Substring(0, 4))
    $runRoot = Join-Path -Path $outputRoot -ChildPath $runId

    $paths = [ordered]@{
        Root       = $runRoot
        Input      = Join-Path -Path $runRoot -ChildPath 'input'
        Extracted  = Join-Path -Path $runRoot -ChildPath 'extracted'
        Normalized = Join-Path -Path $runRoot -ChildPath 'normalized'
        Reports    = Join-Path -Path $runRoot -ChildPath 'reports'
        Dashboard  = Join-Path -Path $runRoot -ChildPath 'dashboard'
        Logs       = Join-Path -Path $runRoot -ChildPath 'logs'
        Manifest   = Join-Path -Path $runRoot -ChildPath 'run.json'
    }

    foreach ($path in @($paths.Root, $paths.Input, $paths.Extracted, $paths.Normalized, $paths.Reports, $paths.Dashboard, $paths.Logs)) {
        New-Item -Path $path -ItemType Directory -Force | Out-Null
    }

    $manifest = [ordered]@{
        RunId       = $runId
        StartedAt   = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        CompletedAt = $null
        Status      = 'Running'
        ConfigUsed  = $Config
        OutputPaths = [pscustomobject]$paths
        Stages      = @()
        Errors      = @()
    }

    $manifest | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $paths.Manifest -Encoding UTF8

    return [pscustomobject]@{
        RunId    = $runId
        Paths    = [pscustomobject]$paths
        Manifest = [pscustomobject]$manifest
    }
}
