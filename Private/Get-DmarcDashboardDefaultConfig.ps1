function Get-DmarcDashboardDefaultConfig {
    [CmdletBinding()]
    param()

    [ordered]@{
        MailboxFolder         = $null
        OutputRoot            = '.\Runs'
        Days                  = 7
        MessageFilter         = 'UnreadOnly'
        MarkAsRead            = $false
        ExistingFileAction    = 'Skip'
        EnableGeoLookup       = $false
        UseGeoCache           = $true
        GeoCachePath          = '.\GEOIP.json'
        GeoProvider           = 'ipinfo'
        GeoApiBaseUri         = 'https://ipinfo.io'
        GeoApiToken           = $null
        GeoApiTokenEnvName    = 'DMARC_DASHBOARD_IPINFO_TOKEN'
        GeoApiTokenSecretName = 'DmarcDashboard-IpInfoToken'
        GeoDelayMilliseconds  = 250
        ForceGeoRefresh       = $false
        CreateCsvReports      = $true
        CreateHtmlReport      = $true
        CreateDashboard       = $true
        KeepRawAttachments    = $true
    }
}
