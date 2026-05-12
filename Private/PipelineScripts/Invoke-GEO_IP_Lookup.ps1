param(
    [Parameter(Mandatory = $false)]
    [string]$InputJsonPath = ".\mastertable.json",

    [Parameter(Mandatory = $false)]
    [string]$OutputJsonPath = ".\GEOIP.json",

    [Parameter(Mandatory = $false)]
    [string]$SourceIpProperty = "sourceip",

    [Parameter(Mandatory = $false)]
    [string]$ApiBaseUri = "https://ipinfo.io",

    [Parameter(Mandatory = $false)]
    [string]$ApiToken = "",

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 10000)]
    [int]$DelayMilliseconds = 250,

    [Parameter(Mandatory = $false)]
    [switch]$ForceRefresh
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

function ConvertTo-Array {
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return @()
    }

    if ($Value -is [System.Array]) {
        return @($Value)
    }

    return @($Value)
}

function ConvertTo-NullableDouble {
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        return $null
    }

    $number = 0.0

    if ([double]::TryParse([string]$Value, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
        return $number
    }

    return $null
}

function Get-IpVersion {
    param(
        [AllowNull()]
        [string]$IpAddress
    )

    if ([string]::IsNullOrWhiteSpace($IpAddress)) {
        return $null
    }

    $parsedIp = $null
    if (-not [System.Net.IPAddress]::TryParse($IpAddress, [ref]$parsedIp)) {
        return $null
    }

    switch ($parsedIp.AddressFamily) {
        ([System.Net.Sockets.AddressFamily]::InterNetwork) { return "IPv4" }
        ([System.Net.Sockets.AddressFamily]::InterNetworkV6) { return "IPv6" }
        default { return $null }
    }
}

function Get-OptionalPropertyValue {
    param(
        [AllowNull()]
        [object]$Object,

        [Parameter(Mandatory = $true)]
        [string]$PropertyName
    )

    if ($null -eq $Object) {
        return $null
    }

    $property = $Object.PSObject.Properties[$PropertyName]
    if ($null -eq $property) {
        return $null
    }

    return $property.Value
}

function ConvertTo-NormalizedGeoRecord {
    param(
        [Parameter(Mandatory = $true)]
        [string]$IpAddress,

        [AllowNull()]
        [object]$Response,

        [AllowNull()]
        [string]$LookupUri,

        [AllowNull()]
        [string]$ErrorMessage
    )

    $latitude = $null
    $longitude = $null

    $responseLoc = Get-OptionalPropertyValue -Object $Response -PropertyName "loc"

    if (-not [string]::IsNullOrWhiteSpace([string]$responseLoc)) {
        $locParts = [string]$responseLoc -split ','

        if ($locParts.Count -ge 2) {
            $latitude = ConvertTo-NullableDouble -Value $locParts[0]
            $longitude = ConvertTo-NullableDouble -Value $locParts[1]
        }
    }

    return [PSCustomObject]@{
        ip            = $IpAddress
        ip_version    = Get-IpVersion -IpAddress $IpAddress
        hostname      = Get-OptionalPropertyValue -Object $Response -PropertyName "hostname"
        city          = Get-OptionalPropertyValue -Object $Response -PropertyName "city"
        region        = Get-OptionalPropertyValue -Object $Response -PropertyName "region"
        country       = Get-OptionalPropertyValue -Object $Response -PropertyName "country"
        loc           = $responseLoc
        latitude      = $latitude
        longitude     = $longitude
        org           = if ($null -ne (Get-OptionalPropertyValue -Object $Response -PropertyName "org")) { Get-OptionalPropertyValue -Object $Response -PropertyName "org" } else { Get-OptionalPropertyValue -Object $Response -PropertyName "as_name" }
        postal        = Get-OptionalPropertyValue -Object $Response -PropertyName "postal"
        timezone      = Get-OptionalPropertyValue -Object $Response -PropertyName "timezone"
        anycast       = Get-OptionalPropertyValue -Object $Response -PropertyName "anycast"
        bogon         = Get-OptionalPropertyValue -Object $Response -PropertyName "bogon"
        readme        = Get-OptionalPropertyValue -Object $Response -PropertyName "readme"
        lookup_status = if ([string]::IsNullOrWhiteSpace($ErrorMessage)) { "ok" } else { "error" }
        lookup_error  = $ErrorMessage
        lookup_uri    = $LookupUri
        updated_utc   = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    }
}

if (-not (Test-Path -LiteralPath $InputJsonPath)) {
    throw "Input JSON file not found: $InputJsonPath"
}

$rawJson = Get-Content -LiteralPath $InputJsonPath -Raw
$records = $rawJson | ConvertFrom-Json

if ($null -eq $records) {
    throw "No records were loaded from: $InputJsonPath"
}

$existingGeoRecords = @()
if (Test-Path -LiteralPath $OutputJsonPath) {
    $existingGeoRaw = Get-Content -LiteralPath $OutputJsonPath -Raw

    if (-not [string]::IsNullOrWhiteSpace($existingGeoRaw)) {
        $existingGeoParsed = $existingGeoRaw | ConvertFrom-Json
        $existingGeoRecords = ConvertTo-Array -Value $existingGeoParsed
    }
}

$geoCache = @{}
foreach ($geoRecord in $existingGeoRecords) {
    if ($null -eq $geoRecord) {
        continue
    }

    $cacheIp = [string]$geoRecord.ip
    if ([string]::IsNullOrWhiteSpace($cacheIp)) {
        continue
    }

    $geoCache[$cacheIp] = $geoRecord
}

$sourceIps = $records |
    ForEach-Object { $_.$SourceIpProperty } |
    Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
    ForEach-Object { [string]$_ } |
    Sort-Object -Unique

if ($sourceIps.Count -eq 0) {
    throw "No non-empty '$SourceIpProperty' values were found in: $InputJsonPath"
}

$lookupCount = 0
$cacheHitCount = 0
$errorCount = 0

foreach ($ipAddress in $sourceIps) {
    if (-not $ForceRefresh.IsPresent -and $geoCache.ContainsKey($ipAddress)) {
        $cacheHitCount++
        continue
    }

    $lookupUri = '{0}/{1}/json' -f $ApiBaseUri.TrimEnd('/'), $ipAddress
    $headers = @{}

    if (-not [string]::IsNullOrWhiteSpace($ApiToken)) {
        $headers["Authorization"] = 'Bearer {0}' -f $ApiToken
    }

    try {
        if ($lookupCount -gt 0 -and $DelayMilliseconds -gt 0) {
            Start-Sleep -Milliseconds $DelayMilliseconds
        }

        $requestParameters = @{
            Uri             = $lookupUri
            Method          = 'Get'
            UseBasicParsing = $true
        }

        if ($headers.Count -gt 0) {
            $requestParameters.Headers = $headers
        }

        $webResponse = Invoke-WebRequest @requestParameters
        $response = $webResponse.Content | ConvertFrom-Json
        $geoCache[$ipAddress] = ConvertTo-NormalizedGeoRecord -IpAddress $ipAddress -Response $response -LookupUri $lookupUri
        $lookupCount++
    }
    catch {
        $geoCache[$ipAddress] = ConvertTo-NormalizedGeoRecord -IpAddress $ipAddress -LookupUri $lookupUri -ErrorMessage $_.Exception.Message
        $lookupCount++
        $errorCount++
    }
}

$outputRecords = $geoCache.Values | Sort-Object -Property ip
$outputRecords | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $OutputJsonPath -Encoding UTF8

[PSCustomObject]@{
    InputJsonPath    = (Resolve-Path -LiteralPath $InputJsonPath).Path
    OutputJsonPath   = (Resolve-Path -LiteralPath $OutputJsonPath).Path
    SourceIpProperty = $SourceIpProperty
    UniqueIpCount    = $sourceIps.Count
    CachedIpCount    = $cacheHitCount
    LookupCount      = $lookupCount
    ErrorCount       = $errorCount
}
