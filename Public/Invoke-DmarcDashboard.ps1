function Invoke-DmarcDashboard {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$ConfigPath,

        [Parameter()]
        [string]$MailboxFolder,

        [Parameter()]
        [string]$OutputRoot,

        [Parameter()]
        [ValidateRange(1, 365)]
        [int]$Days,

        [Parameter()]
        [ValidateSet('UnreadOnly', 'All')]
        [string]$MessageFilter,

        [Parameter()]
        [bool]$MarkAsRead,

        [Parameter()]
        [ValidateSet('Skip', 'Overwrite')]
        [string]$ExistingFileAction,

        [Parameter()]
        [switch]$EnableGeoLookup,

        [Parameter()]
        [switch]$DisableGeoLookup,

        [Parameter()]
        [bool]$UseGeoCache,

        [Parameter()]
        [string]$GeoCachePath,

        [Parameter()]
        [string]$GeoApiToken,

        [Parameter()]
        [string]$GeoApiTokenEnvName,

        [Parameter()]
        [string]$GeoApiTokenSecretName,

        [Parameter()]
        [string]$GeoApiBaseUri,

        [Parameter()]
        [switch]$ForceGeoRefresh,

        [Parameter()]
        [switch]$SkipReports,

        [Parameter()]
        [switch]$SkipDashboard
    )

    $overrides = @{}
    foreach ($key in @('MailboxFolder', 'OutputRoot', 'Days', 'MessageFilter', 'MarkAsRead', 'ExistingFileAction', 'UseGeoCache', 'GeoCachePath', 'GeoApiToken', 'GeoApiTokenEnvName', 'GeoApiTokenSecretName', 'GeoApiBaseUri')) {
        if ($PSBoundParameters.ContainsKey($key)) {
            $overrides[$key] = $PSBoundParameters[$key]
        }
    }

    if ($EnableGeoLookup.IsPresent) {
        $overrides.EnableGeoLookup = $true
    }

    if ($DisableGeoLookup.IsPresent) {
        $overrides.EnableGeoLookup = $false
    }

    if ($ForceGeoRefresh.IsPresent) {
        $overrides.ForceGeoRefresh = $true
    }

    if ($SkipReports.IsPresent) {
        $overrides.CreateCsvReports = $false
        $overrides.CreateHtmlReport = $false
    }

    if ($SkipDashboard.IsPresent) {
        $overrides.CreateDashboard = $false
    }

    $config = Resolve-DmarcDashboardConfig -ConfigPath $ConfigPath -Overrides $overrides
    $run = Initialize-DmarcDashboardRun -Config $config
    $stages = New-Object System.Collections.Generic.List[object]
    $errors = New-Object System.Collections.Generic.List[object]

    $masterTablePath = Join-Path -Path $run.Paths.Normalized -ChildPath 'mastertable.json'
    $geoPath = Join-Path -Path $run.Paths.Normalized -ChildPath 'GEOIP.json'
    $enrichedPath = Join-Path -Path $run.Paths.Normalized -ChildPath 'mastertable.enriched.json'
    $dashboardPath = Join-Path -Path $run.Paths.Dashboard -ChildPath 'dashboard.html'

    try {
        Write-Information "DMARC run $($run.RunId) started at $($run.Paths.Root)"

        $emailResult = Invoke-DmarcScript -ScriptName 'Get-Dmarc_emails.ps1' -Parameters @{
            Days               = [int]$config.Days
            EmailFolderPath    = [string]$config.MailboxFolder
            DestinationFolder  = [string]$run.Paths.Input
            MessageFilter      = [string]$config.MessageFilter
            MarkAsRead         = [bool]$config.MarkAsRead
            ExistingFileAction = [string]$config.ExistingFileAction
        }
        $stages.Add([pscustomobject]@{ Name = 'CollectOutlookAttachments'; Result = $emailResult }) | Out-Null

        $inputFiles = @(Get-ChildItem -LiteralPath $run.Paths.Input -File -ErrorAction SilentlyContinue)
        if ($inputFiles.Count -eq 0) {
            $manifest = Complete-DmarcDashboardRun -ManifestPath $run.Paths.Manifest -Status Completed -Stages $stages.ToArray() -Errors $errors.ToArray()

            return [pscustomobject]@{
                RunId        = $run.RunId
                Status       = $manifest.Status
                RunRoot      = $run.Paths.Root
                ManifestPath = $run.Paths.Manifest
                MasterTable  = $null
                GeoTable     = $null
                Enriched     = $null
                Reports      = $null
                Dashboard    = $null
                Stages       = $stages.ToArray()
            }
        }

        $extractResult = Invoke-DmarcScript -ScriptName 'Extract-Dmarc_reports.ps1' -Parameters @{
            SourceFolder      = [string]$run.Paths.Input
            DestinationFolder = [string]$run.Paths.Extracted
            Overwrite         = ([string]$config.ExistingFileAction -eq 'Overwrite')
        }
        $stages.Add([pscustomobject]@{ Name = 'ExtractReports'; Result = $extractResult }) | Out-Null

        $processResult = Invoke-DmarcScript -ScriptName 'Process-Dmarc_xml.ps1' -Parameters @{
            SourceFolder    = [string]$run.Paths.Extracted
            OutputJsonPath  = $masterTablePath
        }
        $stages.Add([pscustomobject]@{ Name = 'NormalizeXml'; Result = $processResult }) | Out-Null

        $reportInputPath = $masterTablePath
        $geoCacheCopied = $false

        if ([bool]$config.UseGeoCache -and -not [string]::IsNullOrWhiteSpace([string]$config.GeoCachePath)) {
            $sourceGeoCachePath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath([string]$config.GeoCachePath)
            if (Test-Path -LiteralPath $sourceGeoCachePath -PathType Leaf) {
                Copy-Item -LiteralPath $sourceGeoCachePath -Destination $geoPath -Force
                $geoCacheCopied = $true
                $stages.Add([pscustomobject]@{
                    Name   = 'LoadGeoCache'
                    Result = [pscustomobject]@{
                        SourceGeoCachePath = $sourceGeoCachePath
                        OutputJsonPath     = $geoPath
                    }
                }) | Out-Null
            }
        }

        if ([bool]$config.EnableGeoLookup) {
            $token = Get-DmarcGeoApiToken `
                -ExplicitToken $config.GeoApiToken `
                -EnvironmentVariableName $config.GeoApiTokenEnvName `
                -SecretName $config.GeoApiTokenSecretName

            if ([string]::IsNullOrWhiteSpace($token)) {
                throw "GEO lookup is enabled, but no token was found. Set GeoApiToken, $($config.GeoApiTokenEnvName), or SecretManagement secret $($config.GeoApiTokenSecretName)."
            }

            $geoLookupOutputPath = $geoPath
            if (-not [string]::IsNullOrWhiteSpace([string]$config.GeoCachePath)) {
                $geoLookupOutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath([string]$config.GeoCachePath)
                $geoLookupOutputDirectory = Split-Path -Path $geoLookupOutputPath -Parent
                if (-not [string]::IsNullOrWhiteSpace($geoLookupOutputDirectory) -and -not (Test-Path -LiteralPath $geoLookupOutputDirectory -PathType Container)) {
                    New-Item -Path $geoLookupOutputDirectory -ItemType Directory -Force | Out-Null
                }
            }

            $geoParams = @{
                InputJsonPath      = $masterTablePath
                OutputJsonPath     = $geoLookupOutputPath
                ApiBaseUri         = [string]$config.GeoApiBaseUri
                ApiToken           = $token
                DelayMilliseconds  = [int]$config.GeoDelayMilliseconds
            }

            if ([bool]$config.ForceGeoRefresh) {
                $geoParams.ForceRefresh = $true
            }

            $geoResult = Invoke-DmarcScript -ScriptName 'Invoke-GEO_IP_Lookup.ps1' -Parameters $geoParams
            $stages.Add([pscustomobject]@{ Name = 'GeoLookup'; Result = $geoResult }) | Out-Null
            if ($geoLookupOutputPath -ne $geoPath -and (Test-Path -LiteralPath $geoLookupOutputPath -PathType Leaf)) {
                Copy-Item -LiteralPath $geoLookupOutputPath -Destination $geoPath -Force
            }
            $geoCacheCopied = $true
        }

        if ($geoCacheCopied -and (Test-Path -LiteralPath $geoPath -PathType Leaf)) {
            $mergeResult = Invoke-DmarcScript -ScriptName 'Merge-GEOIntoMasterTable.ps1' -Parameters @{
                MasterTableJsonPath = $masterTablePath
                GeoIpJsonPath       = $geoPath
                OutputJsonPath      = $enrichedPath
            }
            $stages.Add([pscustomobject]@{ Name = 'MergeGeo'; Result = $mergeResult }) | Out-Null
            $reportInputPath = $enrichedPath
        }

        if ([bool]$config.CreateCsvReports -or [bool]$config.CreateHtmlReport) {
            $reportResult = Invoke-DmarcScript -ScriptName 'New-DMarcReport.ps1' -Parameters @{
                InputJsonPath   = $reportInputPath
                OutputDirectory = [string]$run.Paths.Reports
            }
            $stages.Add([pscustomobject]@{ Name = 'GenerateReports'; Result = $reportResult }) | Out-Null
        }

        if ([bool]$config.CreateDashboard) {
            $dashboardResult = Invoke-DmarcScript -ScriptName 'Generate-Dashboard.ps1' -Parameters @{
                InputJsonPath = $reportInputPath
                OutputPath    = $dashboardPath
            }
            $stages.Add([pscustomobject]@{ Name = 'GenerateDashboard'; Result = $dashboardResult }) | Out-Null
        }

        $manifest = Complete-DmarcDashboardRun -ManifestPath $run.Paths.Manifest -Status Completed -Stages $stages.ToArray() -Errors $errors.ToArray()

        [pscustomobject]@{
            RunId        = $run.RunId
            Status       = $manifest.Status
            RunRoot      = $run.Paths.Root
            ManifestPath = $run.Paths.Manifest
            MasterTable  = $masterTablePath
            GeoTable     = if (Test-Path -LiteralPath $geoPath) { $geoPath } else { $null }
            Enriched     = if (Test-Path -LiteralPath $enrichedPath) { $enrichedPath } else { $null }
            Reports      = $run.Paths.Reports
            Dashboard    = if (Test-Path -LiteralPath $dashboardPath) { $dashboardPath } else { $null }
            Stages       = $stages.ToArray()
        }
    }
    catch {
        $errors.Add([pscustomobject]@{
            Message = $_.Exception.Message
            Type    = $_.Exception.GetType().FullName
        }) | Out-Null

        Complete-DmarcDashboardRun -ManifestPath $run.Paths.Manifest -Status Failed -Stages $stages.ToArray() -Errors $errors.ToArray() | Out-Null
        throw
    }
}
