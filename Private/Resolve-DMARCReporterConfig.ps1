function Resolve-DMARCReporterConfig {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$ConfigPath,

        [Parameter()]
        [hashtable]$Overrides
    )

    $resolved = Get-DMARCReporterDefaultConfig

    if (-not [string]::IsNullOrWhiteSpace($ConfigPath)) {
        $configFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ConfigPath)
        if (-not (Test-Path -LiteralPath $configFile -PathType Leaf)) {
            throw "Config file not found: $configFile"
        }

        $rawConfig = Get-Content -LiteralPath $configFile -Raw
        if (-not [string]::IsNullOrWhiteSpace($rawConfig)) {
            $jsonConfig = $rawConfig | ConvertFrom-Json
            foreach ($property in $jsonConfig.PSObject.Properties) {
                $resolved[$property.Name] = $property.Value
            }
        }
    }

    if ($null -ne $Overrides) {
        foreach ($key in $Overrides.Keys) {
            $resolved[$key] = $Overrides[$key]
        }
    }

    if ([string]::IsNullOrWhiteSpace([string]$resolved.MailboxFolder)) {
        throw "MailboxFolder is required. Pass -MailboxFolder or set MailboxFolder in the config file."
    }

    if ([string]::IsNullOrWhiteSpace([string]$resolved.OutputRoot)) {
        throw "OutputRoot is required."
    }

    $resolved.Days = [int]$resolved.Days
    if ($resolved.Days -lt 1 -or $resolved.Days -gt 365) {
        throw "Days must be between 1 and 365. Value received: $($resolved.Days)"
    }

    if ($resolved.MessageFilter -notin @('UnreadOnly', 'All')) {
        throw "MessageFilter must be UnreadOnly or All. Value received: $($resolved.MessageFilter)"
    }

    if ($resolved.ExistingFileAction -notin @('Skip', 'Overwrite')) {
        throw "ExistingFileAction must be Skip or Overwrite. Value received: $($resolved.ExistingFileAction)"
    }

    return [pscustomobject]$resolved
}
