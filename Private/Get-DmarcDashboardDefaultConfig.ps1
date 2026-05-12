function Get-DmarcDashboardDefaultConfig {
    [CmdletBinding()]
    [OutputType([System.Collections.Specialized.OrderedDictionary])]
    param()

    $config = [System.Collections.Specialized.OrderedDictionary]::new()
    $config['MailboxFolder'] = $null
    $config['OutputRoot'] = '.\Runs'
    $config['Days'] = 7
    $config['MessageFilter'] = 'UnreadOnly'
    $config['MarkAsRead'] = $false
    $config['ExistingFileAction'] = 'Skip'
    $config['EnableGeoLookup'] = $false
    $config['UseGeoCache'] = $true
    $config['GeoCachePath'] = '.\GEOIP.json'
    $config['GeoProvider'] = 'ipinfo'
    $config['GeoApiBaseUri'] = 'https://ipinfo.io'
    $config['GeoApiToken'] = $null
    $config['GeoApiTokenEnvName'] = 'DMARC_DASHBOARD_IPINFO_TOKEN'
    $config['GeoApiTokenSecretName'] = 'DmarcDashboard-IpInfoToken'
    $config['GeoDelayMilliseconds'] = 250
    $config['ForceGeoRefresh'] = $false
    $config['CreateCsvReports'] = $true
    $config['CreateHtmlReport'] = $true
    $config['CreateDashboard'] = $true
    $config['KeepRawAttachments'] = $true
    return $config
}
