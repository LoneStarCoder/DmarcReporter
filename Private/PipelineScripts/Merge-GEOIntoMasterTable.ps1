param(
    [Parameter(Mandatory = $false)]
    [string]$MasterTableJsonPath = ".\mastertable.json",

    [Parameter(Mandatory = $false)]
    [string]$GeoIpJsonPath = ".\GEOIP.json",

    [Parameter(Mandatory = $false)]
    [string]$OutputJsonPath = ".\mastertable.enriched.json",

    [Parameter(Mandatory = $false)]
    [string]$SourceIpProperty = "sourceip"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

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

if (-not (Test-Path -LiteralPath $MasterTableJsonPath)) {
    throw "Master table JSON file not found: $MasterTableJsonPath"
}

if (-not (Test-Path -LiteralPath $GeoIpJsonPath)) {
    throw "GEO IP JSON file not found: $GeoIpJsonPath"
}

$masterTableRaw = Get-Content -LiteralPath $MasterTableJsonPath -Raw
[array]$masterTableRows = ConvertTo-Array -Value ($masterTableRaw | ConvertFrom-Json)

$geoIpRaw = Get-Content -LiteralPath $GeoIpJsonPath -Raw
[array]$geoIpRows = ConvertTo-Array -Value ($geoIpRaw | ConvertFrom-Json)

$geoByIp = @{}
foreach ($geoRow in $geoIpRows) {
    $ip = [string](Get-OptionalPropertyValue -Object $geoRow -PropertyName "ip")
    if ([string]::IsNullOrWhiteSpace($ip)) {
        continue
    }

    $geoByIp[$ip] = $geoRow
}

$matchedRowCount = 0
$missingGeoRowCount = 0

$enrichedRows = foreach ($row in $masterTableRows) {
    $enriched = [ordered]@{}

    foreach ($property in $row.PSObject.Properties) {
        $enriched[$property.Name] = $property.Value
    }

    $sourceIp = [string](Get-OptionalPropertyValue -Object $row -PropertyName $SourceIpProperty)
    $geoRecord = $null

    if (-not [string]::IsNullOrWhiteSpace($sourceIp) -and $geoByIp.ContainsKey($sourceIp)) {
        $geoRecord = $geoByIp[$sourceIp]
        $matchedRowCount++
    }
    else {
        $missingGeoRowCount++
    }

    $enriched["geo_lookup_status"] = Get-OptionalPropertyValue -Object $geoRecord -PropertyName "lookup_status"
    $enriched["geo_lookup_error"] = Get-OptionalPropertyValue -Object $geoRecord -PropertyName "lookup_error"
    $enriched["geo_lookup_uri"] = Get-OptionalPropertyValue -Object $geoRecord -PropertyName "lookup_uri"
    $enriched["geo_updated_utc"] = Get-OptionalPropertyValue -Object $geoRecord -PropertyName "updated_utc"
    $enriched["geo_hostname"] = Get-OptionalPropertyValue -Object $geoRecord -PropertyName "hostname"
    $enriched["geo_city"] = Get-OptionalPropertyValue -Object $geoRecord -PropertyName "city"
    $enriched["geo_region"] = Get-OptionalPropertyValue -Object $geoRecord -PropertyName "region"
    $enriched["geo_country"] = Get-OptionalPropertyValue -Object $geoRecord -PropertyName "country"
    $enriched["geo_loc"] = Get-OptionalPropertyValue -Object $geoRecord -PropertyName "loc"
    $enriched["geo_latitude"] = Get-OptionalPropertyValue -Object $geoRecord -PropertyName "latitude"
    $enriched["geo_longitude"] = Get-OptionalPropertyValue -Object $geoRecord -PropertyName "longitude"
    $enriched["geo_org"] = Get-OptionalPropertyValue -Object $geoRecord -PropertyName "org"
    $enriched["geo_postal"] = Get-OptionalPropertyValue -Object $geoRecord -PropertyName "postal"
    $enriched["geo_timezone"] = Get-OptionalPropertyValue -Object $geoRecord -PropertyName "timezone"
    $enriched["geo_anycast"] = Get-OptionalPropertyValue -Object $geoRecord -PropertyName "anycast"
    $enriched["geo_bogon"] = Get-OptionalPropertyValue -Object $geoRecord -PropertyName "bogon"
    $enriched["geo_readme"] = Get-OptionalPropertyValue -Object $geoRecord -PropertyName "readme"

    [PSCustomObject]$enriched
}

$enrichedRows | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $OutputJsonPath -Encoding UTF8

[PSCustomObject]@{
    MasterTableJsonPath = (Resolve-Path -LiteralPath $MasterTableJsonPath).Path
    GeoIpJsonPath       = (Resolve-Path -LiteralPath $GeoIpJsonPath).Path
    OutputJsonPath      = (Resolve-Path -LiteralPath $OutputJsonPath).Path
    SourceIpProperty    = $SourceIpProperty
    TotalRows           = $masterTableRows.Count
    GeoIpRecords        = $geoIpRows.Count
    MatchedRows         = $matchedRowCount
    MissingGeoRows      = $missingGeoRowCount
}
