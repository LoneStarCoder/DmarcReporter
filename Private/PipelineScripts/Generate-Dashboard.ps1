param(
    [Parameter(Mandatory = $false)]
    [string]$InputJsonPath,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".\DmarcDashboard\dashboard.html",

    [Parameter(Mandatory = $false)]
    [string]$Title = "DMARC Intelligence Dashboard"
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

function ConvertTo-IntSafe {
    param(
        [AllowNull()]
        [object]$Value
    )

    $number = 0

    if ([int]::TryParse([string]$Value, [ref]$number)) {
        return $number
    }

    return 0
}

function ConvertTo-NullableInt {
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        return $null
    }

    $number = 0

    if ([int]::TryParse([string]$Value, [ref]$number)) {
        return $number
    }

    return $null
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

function ConvertTo-JoinedString {
    param(
        [AllowNull()]
        [object]$Value,

        [string]$Separator = "; "
    )

    [array]$items = @(ConvertTo-Array -Value $Value)

    if ($items.Count -eq 0) {
        return $null
    }

    return (
        $items |
            Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
            ForEach-Object { [string]$_ } |
            Select-Object -Unique
    ) -join $Separator
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

function Get-ResolvedInputPath {
    param(
        [AllowNull()]
        [string]$RequestedPath
    )

    if (-not [string]::IsNullOrWhiteSpace($RequestedPath)) {
        return $RequestedPath
    }

    if (Test-Path -LiteralPath ".\mastertable.enriched.json") {
        return ".\mastertable.enriched.json"
    }

    return ".\mastertable.json"
}

function Get-RiskScore {
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Row
    )

    $score = 0

    if (-not $Row.DmarcAligned) { $score += 60 }
    if ($Row.DmarcSpf -ne "pass") { $score += 12 }
    if ($Row.DmarcDkim -ne "pass") { $score += 12 }
    if ($Row.SpfResult -and $Row.SpfResult -ne "pass") { $score += 8 }
    if ($Row.DkimResult -match 'fail|permerror|temperror|neutral|none') { $score += 8 }
    if ($Row.GeoLookupStatus -and $Row.GeoLookupStatus -ne "ok") { $score += 4 }
    if ($Row.GeoCountry -and $Row.GeoCountry -ne "US") { $score += 7 }
    if ($Row.PolicyOverrideType) { $score += 5 }

    if ($Row.SourceIpCount -ge 100) { $score += 12 }
    elseif ($Row.SourceIpCount -ge 25) { $score += 8 }
    elseif ($Row.SourceIpCount -ge 10) { $score += 4 }

    if ($score -gt 100) {
        $score = 100
    }

    return $score
}

function Get-RiskBand {
    param(
        [int]$RiskScore
    )

    if ($RiskScore -ge 75) { return "Critical" }
    if ($RiskScore -ge 55) { return "High" }
    if ($RiskScore -ge 35) { return "Elevated" }
    return "Low"
}

function Get-MessageSum {
    param(
        [AllowEmptyCollection()]
        [object[]]$Rows
    )

    $sum = 0

    foreach ($row in @($Rows)) {
        $sum += ConvertTo-IntSafe -Value $row.SourceIpCount
    }

    return $sum
}

function ConvertTo-HtmlEncodedText {
    param(
        [AllowNull()]
        [object]$Text
    )

    if ($null -eq $Text) {
        return ""
    }

    return [System.Net.WebUtility]::HtmlEncode([string]$Text)
}

function Format-Number {
    param(
        [AllowNull()]
        [object]$Value
    )

    $number = ConvertTo-NullableDouble -Value $Value
    if ($null -eq $number) {
        return "0"
    }

    return ('{0:N0}' -f $number)
}

function Format-Percent {
    param(
        [double]$Part,
        [double]$Whole
    )

    if ($Whole -le 0) {
        return "0.0%"
    }

    return ('{0:N1}%' -f (($Part / $Whole) * 100))
}

function Get-DateLabel {
    param(
        [AllowNull()]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return "Unknown"
    }

    try {
        return ([datetime]$Value).ToString("yyyy-MM-dd")
    }
    catch {
        return $Value
    }
}

function Get-GroupSummary {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [AllowEmptyCollection()]
        [object[]]$Rows,

        [Parameter(Mandatory = $true)]
        [string[]]$GroupProperties
    )

    if ($null -eq $Rows -or $Rows.Count -eq 0) {
        return @()
    }

    $Rows |
        Group-Object -Property $GroupProperties |
        ForEach-Object {
            $first = $_.Group[0]
            $messageCount = 0
            $quarantineCount = 0

            foreach ($groupRow in @($_.Group)) {
                $messageCount += ConvertTo-IntSafe -Value $groupRow.SourceIpCount

                if ($groupRow.WouldQuarantine) {
                    $quarantineCount += ConvertTo-IntSafe -Value $groupRow.SourceIpCount
                }
            }

            $output = [ordered]@{
                RecordCount     = $_.Count
                MessageCount    = [int]$messageCount
                QuarantineCount = [int]$quarantineCount
            }

            foreach ($property in $GroupProperties) {
                $output[$property] = $first.$property
            }

            [PSCustomObject]$output
        } |
        Sort-Object -Property MessageCount -Descending
}

function ConvertTo-MetricCardHtml {
    param(
        [string]$Label,
        [string]$Value,
        [string]$Hint = "",
        [string]$Tone = "default"
    )

    return @"
<article class="metric-card tone-$Tone">
  <div class="metric-label">$(ConvertTo-HtmlEncodedText $Label)</div>
  <div class="metric-value">$(ConvertTo-HtmlEncodedText $Value)</div>
  <div class="metric-hint">$(ConvertTo-HtmlEncodedText $Hint)</div>
</article>
"@
}

function ConvertTo-LeaderboardHtml {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [AllowEmptyCollection()]
        [object[]]$Rows,

        [Parameter(Mandatory = $true)]
        [string]$LabelProperty,

        [Parameter(Mandatory = $true)]
        [string]$ValueProperty,

        [string[]]$MetaProperties = @(),

        [int]$MaxRows = 10
    )

    $items = @($Rows | Select-Object -First $MaxRows)
    if ($items.Count -eq 0) {
        return @"
<section class="panel">
  <div class="panel-head">
    <h2>$(ConvertTo-HtmlEncodedText $Title)</h2>
  </div>
  <div class="empty-state">No data available.</div>
</section>
"@
    }

    $maxValue = [double](($items | Measure-Object -Property $ValueProperty -Maximum).Maximum)
    if ($maxValue -le 0) {
        $maxValue = 1
    }

    $rowsHtml = foreach ($row in $items) {
        $label = $row.$LabelProperty
        if ([string]::IsNullOrWhiteSpace([string]$label)) {
            $label = "Unknown"
        }

        $value = [double]$row.$ValueProperty
        $width = [math]::Round(($value / $maxValue) * 100, 1)

        $metaParts = foreach ($metaProperty in $MetaProperties) {
            $metaValue = $row.$metaProperty
            if (-not [string]::IsNullOrWhiteSpace([string]$metaValue)) {
                ConvertTo-HtmlEncodedText $metaValue
            }
        }

        @"
<div class="leader-row">
  <div class="leader-row-head">
    <div class="leader-label">$(ConvertTo-HtmlEncodedText $label)</div>
    <div class="leader-value">$(ConvertTo-HtmlEncodedText (Format-Number $value))</div>
  </div>
  <div class="leader-meta">$([string]::Join(' | ', @($metaParts)))</div>
  <div class="leader-bar-track"><div class="leader-bar" style="width: $width%"></div></div>
</div>
"@
    }

    return @"
<section class="panel">
  <div class="panel-head">
    <h2>$(ConvertTo-HtmlEncodedText $Title)</h2>
  </div>
  <div class="leaderboard">
    $([string]::Join("`n", @($rowsHtml)))
  </div>
</section>
"@
}

function ConvertTo-TableHtml {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $false)]
        [string]$Subtitle = "",

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [AllowEmptyCollection()]
        [object[]]$Rows,

        [Parameter(Mandatory = $true)]
        [hashtable[]]$Columns,

        [int]$MaxRows = 25
    )

    $items = @($Rows | Select-Object -First $MaxRows)
    if ($items.Count -eq 0) {
        return @"
<section class="panel">
  <div class="panel-head">
    <div>
      <h2>$(ConvertTo-HtmlEncodedText $Title)</h2>
      <p class="panel-subtitle">$(ConvertTo-HtmlEncodedText $Subtitle)</p>
    </div>
  </div>
  <div class="empty-state">No data available.</div>
</section>
"@
    }

    $headerHtml = foreach ($column in $Columns) {
        "<th>$(ConvertTo-HtmlEncodedText $column.Label)</th>"
    }

    $bodyHtml = foreach ($row in $items) {
        $cells = foreach ($column in $Columns) {
            $rawValue = $row.($column.Property)

            if ($column.ContainsKey("Formatter") -and $null -ne $column.Formatter) {
                $displayValue = & $column.Formatter $rawValue $row
            }
            else {
                $displayValue = $rawValue
            }

            "<td>$(ConvertTo-HtmlEncodedText $displayValue)</td>"
        }

        "<tr>$([string]::Join('', @($cells)))</tr>"
    }

    return @"
<section class="panel">
  <div class="panel-head">
    <div>
      <h2>$(ConvertTo-HtmlEncodedText $Title)</h2>
      <p class="panel-subtitle">$(ConvertTo-HtmlEncodedText $Subtitle)</p>
    </div>
  </div>
  <div class="table-wrap">
    <table class="data-table">
      <thead><tr>$([string]::Join('', @($headerHtml)))</tr></thead>
      <tbody>
        $([string]::Join("`n", @($bodyHtml)))
      </tbody>
    </table>
  </div>
</section>
"@
}

function ConvertTo-DistributionHtml {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $true)]
        [hashtable[]]$Segments
    )

    $total = 0
    foreach ($segment in $Segments) {
        $total += [double]$segment["Value"]
    }
    if ($total -le 0) {
        $total = 1
    }

    $barsHtml = foreach ($segment in $Segments) {
        $width = [math]::Round(($segment.Value / $total) * 100, 2)
        "<div class=""stack-segment tone-$($segment.Tone)"" style=""width: $width%"" title=""$(ConvertTo-HtmlEncodedText $segment.Label): $(ConvertTo-HtmlEncodedText (Format-Number $segment.Value))""></div>"
    }

    $legendHtml = foreach ($segment in $Segments) {
        @"
<div class="legend-item">
  <span class="legend-swatch tone-$($segment.Tone)"></span>
  <span class="legend-label">$(ConvertTo-HtmlEncodedText $segment.Label)</span>
  <span class="legend-value">$(ConvertTo-HtmlEncodedText (Format-Number $segment.Value))</span>
  <span class="legend-percent">$(ConvertTo-HtmlEncodedText (Format-Percent $segment.Value $total))</span>
</div>
"@
    }

    return @"
<section class="panel">
  <div class="panel-head">
    <h2>$(ConvertTo-HtmlEncodedText $Title)</h2>
  </div>
  <div class="stack-chart">
    $([string]::Join('', @($barsHtml)))
  </div>
  <div class="legend-grid">
    $([string]::Join("`n", @($legendHtml)))
  </div>
</section>
"@
}

$resolvedInputJsonPath = Get-ResolvedInputPath -RequestedPath $InputJsonPath
if (-not (Test-Path -LiteralPath $resolvedInputJsonPath)) {
    throw "Input JSON file not found: $resolvedInputJsonPath"
}

$outputDirectory = Split-Path -Path $OutputPath -Parent
if ([string]::IsNullOrWhiteSpace($outputDirectory)) {
    $outputDirectory = "."
}

if (-not (Test-Path -LiteralPath $outputDirectory)) {
    New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
}

$records = Get-Content -LiteralPath $resolvedInputJsonPath -Raw | ConvertFrom-Json
$records = ConvertTo-Array -Value $records

if ($records.Count -eq 0) {
    throw "No records were loaded from: $resolvedInputJsonPath"
}

$normalized = foreach ($record in $records) {
    $sourceIpCount = ConvertTo-IntSafe -Value $record.sourceipcount
    [array]$dkimDomains = @(ConvertTo-Array -Value (Get-OptionalPropertyValue -Object $record -PropertyName "dkimdomain"))
    [array]$dkimSelectors = @(ConvertTo-Array -Value (Get-OptionalPropertyValue -Object $record -PropertyName "dkimselector"))
    [array]$dkimResults = @(ConvertTo-Array -Value (Get-OptionalPropertyValue -Object $record -PropertyName "dkimresult"))
    [array]$policyOverrideTypes = @(ConvertTo-Array -Value (Get-OptionalPropertyValue -Object $record -PropertyName "policyoverridetype"))
    [array]$policyOverrideComments = @(ConvertTo-Array -Value (Get-OptionalPropertyValue -Object $record -PropertyName "policyoverridecomment"))

    $dkimPairs = for ($index = 0; $index -lt ([Math]::Max($dkimDomains.Count, $dkimResults.Count)); $index++) {
        $domain = if ($index -lt $dkimDomains.Count) { [string]$dkimDomains[$index] } else { "" }
        $result = if ($index -lt $dkimResults.Count) { [string]$dkimResults[$index] } else { "" }

        if (-not [string]::IsNullOrWhiteSpace($domain) -or -not [string]::IsNullOrWhiteSpace($result)) {
            if ([string]::IsNullOrWhiteSpace($domain)) {
                $result
            }
            elseif ([string]::IsNullOrWhiteSpace($result)) {
                $domain
            }
            else {
                "$domain=$result"
            }
        }
    }

    $spfPass = [string]$record.dmarcspf -eq "pass"
    $dkimPass = [string]$record.dmarcdkim -eq "pass"
    $dmarcAligned = $spfPass -or $dkimPass
    $geoCountry = Get-OptionalPropertyValue -Object $record -PropertyName "geo_country"
    if ([string]::IsNullOrWhiteSpace([string]$geoCountry)) {
        $geoCountry = Get-OptionalPropertyValue -Object $record -PropertyName "country"
    }

    $row = [PSCustomObject]@{
        ProcessDate         = Get-OptionalPropertyValue -Object $record -PropertyName "processdate"
        ReportId            = Get-OptionalPropertyValue -Object $record -PropertyName "reportid"
        ReportEmail         = Get-OptionalPropertyValue -Object $record -PropertyName "reportemail"
        ReportDateBegin     = ConvertTo-NullableInt -Value (Get-OptionalPropertyValue -Object $record -PropertyName "reportdatebegin")
        ReportDateEnd       = ConvertTo-NullableInt -Value (Get-OptionalPropertyValue -Object $record -PropertyName "reportdateend")
        ReportDateBeginUtc  = Get-OptionalPropertyValue -Object $record -PropertyName "reportdatebeginutc"
        ReportDateEndUtc    = Get-OptionalPropertyValue -Object $record -PropertyName "reportdateendutc"
        ReportDay           = Get-DateLabel -Value (Get-OptionalPropertyValue -Object $record -PropertyName "reportdatebeginutc")
        OrgName             = Get-OptionalPropertyValue -Object $record -PropertyName "orgname"
        DmarcDomain         = Get-OptionalPropertyValue -Object $record -PropertyName "dmarcdomain"
        PolicyP             = Get-OptionalPropertyValue -Object $record -PropertyName "policyp"
        PolicySp            = Get-OptionalPropertyValue -Object $record -PropertyName "policysp"
        PolicyPct           = ConvertTo-NullableInt -Value (Get-OptionalPropertyValue -Object $record -PropertyName "policypct")
        PolicyAdkim         = Get-OptionalPropertyValue -Object $record -PropertyName "policyadkim"
        PolicyAspf          = Get-OptionalPropertyValue -Object $record -PropertyName "policyaspf"
        PolicyFo            = Get-OptionalPropertyValue -Object $record -PropertyName "policyfo"
        SourceIp            = Get-OptionalPropertyValue -Object $record -PropertyName "sourceip"
        SourceIpVersion     = Get-OptionalPropertyValue -Object $record -PropertyName "sourceipversion"
        SourceIpCount       = $sourceIpCount
        DmarcDisposition    = Get-OptionalPropertyValue -Object $record -PropertyName "dmarcdisposition"
        DmarcSpf            = Get-OptionalPropertyValue -Object $record -PropertyName "dmarcspf"
        DmarcDkim           = Get-OptionalPropertyValue -Object $record -PropertyName "dmarcdkim"
        PolicyOverrideType  = ConvertTo-JoinedString -Value $policyOverrideTypes
        PolicyOverrideComment = ConvertTo-JoinedString -Value $policyOverrideComments
        EnvelopeFrom        = Get-OptionalPropertyValue -Object $record -PropertyName "envelopefrom"
        HeaderFrom          = Get-OptionalPropertyValue -Object $record -PropertyName "headerfrom"
        DkimDomain          = ConvertTo-JoinedString -Value $dkimDomains
        DkimSelector        = ConvertTo-JoinedString -Value $dkimSelectors
        DkimResult          = ConvertTo-JoinedString -Value $dkimResults
        DkimPairs           = ($dkimPairs | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join "; "
        SpfDomain           = Get-OptionalPropertyValue -Object $record -PropertyName "spfdomain"
        SpfScope            = Get-OptionalPropertyValue -Object $record -PropertyName "spfscope"
        SpfResult           = Get-OptionalPropertyValue -Object $record -PropertyName "spfresult"
        GeoLookupStatus     = Get-OptionalPropertyValue -Object $record -PropertyName "geo_lookup_status"
        GeoLookupError      = Get-OptionalPropertyValue -Object $record -PropertyName "geo_lookup_error"
        GeoHostname         = Get-OptionalPropertyValue -Object $record -PropertyName "geo_hostname"
        GeoCity             = Get-OptionalPropertyValue -Object $record -PropertyName "geo_city"
        GeoRegion           = Get-OptionalPropertyValue -Object $record -PropertyName "geo_region"
        GeoCountry          = $geoCountry
        GeoLatitude         = ConvertTo-NullableDouble -Value (Get-OptionalPropertyValue -Object $record -PropertyName "geo_latitude")
        GeoLongitude        = ConvertTo-NullableDouble -Value (Get-OptionalPropertyValue -Object $record -PropertyName "geo_longitude")
        GeoOrg              = Get-OptionalPropertyValue -Object $record -PropertyName "geo_org"
        GeoPostal           = Get-OptionalPropertyValue -Object $record -PropertyName "geo_postal"
        GeoTimezone         = Get-OptionalPropertyValue -Object $record -PropertyName "geo_timezone"
        DmarcAligned        = $dmarcAligned
        WouldQuarantine     = (-not $dmarcAligned)
        BothSpfAndDkimPass  = ($spfPass -and $dkimPass)
        BothSpfAndDkimFail  = (-not $spfPass -and -not $dkimPass)
    }

    $riskScore = Get-RiskScore -Row $row
    $outputRow = [ordered]@{}

    foreach ($property in $row.PSObject.Properties) {
        $outputRow[$property.Name] = $property.Value
    }

    $outputRow["RiskScore"] = $riskScore
    $outputRow["RiskBand"] = Get-RiskBand -RiskScore $riskScore

    [PSCustomObject]$outputRow
}

$totalRecords = $normalized.Count
$totalMessages = Get-MessageSum -Rows @($normalized)
$alignedMessages = Get-MessageSum -Rows @($normalized | Where-Object { $_.DmarcAligned })
$notAlignedMessages = Get-MessageSum -Rows @($normalized | Where-Object { -not $_.DmarcAligned })
$spfPassMessages = Get-MessageSum -Rows @($normalized | Where-Object { $_.DmarcSpf -eq "pass" })
$spfFailMessages = Get-MessageSum -Rows @($normalized | Where-Object { $_.DmarcSpf -ne "pass" })
$dkimPassMessages = Get-MessageSum -Rows @($normalized | Where-Object { $_.DmarcDkim -eq "pass" })
$dkimFailMessages = Get-MessageSum -Rows @($normalized | Where-Object { $_.DmarcDkim -ne "pass" })
$quarantineMessages = Get-MessageSum -Rows @($normalized | Where-Object { $_.WouldQuarantine })
$criticalMessages = Get-MessageSum -Rows @($normalized | Where-Object { $_.RiskBand -eq "Critical" })
$geoResolvedMessages = Get-MessageSum -Rows @($normalized | Where-Object { $_.GeoLookupStatus -eq "ok" })
$geoCountries = @($normalized | Where-Object { $_.GeoCountry } | Select-Object -ExpandProperty GeoCountry -Unique)
$geoCoverage = Format-Percent $geoResolvedMessages $totalMessages
$dateRangeStart = ($normalized | Where-Object { $_.ReportDateBeginUtc } | Sort-Object -Property ReportDateBeginUtc | Select-Object -First 1 -ExpandProperty ReportDateBeginUtc)
$dateRangeEnd = ($normalized | Where-Object { $_.ReportDateEndUtc } | Sort-Object -Property ReportDateEndUtc -Descending | Select-Object -First 1 -ExpandProperty ReportDateEndUtc)
$uniqueSourceIps = @($normalized | Select-Object -ExpandProperty SourceIp -Unique).Count
$uniqueReporters = @($normalized | Select-Object -ExpandProperty OrgName -Unique).Count
$uniquePolicyDomains = @($normalized | Select-Object -ExpandProperty DmarcDomain -Unique).Count
$uniqueHeaderFrom = @($normalized | Select-Object -ExpandProperty HeaderFrom -Unique).Count

$policySummary = Get-GroupSummary -Rows @($normalized) -GroupProperties @("PolicyP", "PolicyAdkim", "PolicyAspf", "PolicyPct")
$orgSummary = Get-GroupSummary -Rows @($normalized) -GroupProperties @("OrgName")
$ipSummary = Get-GroupSummary -Rows @($normalized) -GroupProperties @("SourceIp", "GeoCountry", "GeoCity", "GeoOrg")
$countrySummary = Get-GroupSummary -Rows @($normalized | Where-Object { $_.GeoCountry }) -GroupProperties @("GeoCountry")
$citySummary = Get-GroupSummary -Rows @($normalized | Where-Object { $_.GeoCity }) -GroupProperties @("GeoCity", "GeoRegion", "GeoCountry")
$spfDomainSummary = Get-GroupSummary -Rows @($normalized | Where-Object { $_.SpfDomain }) -GroupProperties @("SpfDomain", "SpfResult")
$dkimDomainSummary = Get-GroupSummary -Rows @($normalized | Where-Object { $_.DkimDomain }) -GroupProperties @("DkimDomain", "DkimResult")
$headerFromSummary = Get-GroupSummary -Rows @($normalized | Where-Object { $_.HeaderFrom }) -GroupProperties @("HeaderFrom")
$reportDaySummary = Get-GroupSummary -Rows @($normalized) -GroupProperties @("ReportDay")
$geoOrgSummary = Get-GroupSummary -Rows @($normalized | Where-Object { $_.GeoOrg }) -GroupProperties @("GeoOrg", "GeoCountry")
$overrideSummary = Get-GroupSummary -Rows @($normalized | Where-Object { $_.PolicyOverrideType }) -GroupProperties @("PolicyOverrideType")

$highRiskRows = $normalized | Sort-Object -Property @{ Expression = "RiskScore"; Descending = $true }, @{ Expression = "SourceIpCount"; Descending = $true }
$quarantineRows = $normalized | Where-Object { $_.WouldQuarantine } | Sort-Object -Property @{ Expression = "SourceIpCount"; Descending = $true }, @{ Expression = "RiskScore"; Descending = $true }
$internationalRows = $normalized | Where-Object { $_.GeoCountry -and $_.GeoCountry -ne "US" } | Sort-Object -Property SourceIpCount -Descending
$authMatrix = Get-GroupSummary -Rows @($normalized) -GroupProperties @("DmarcSpf", "DmarcDkim", "DmarcDisposition")

$metricsHtml = @(
    (ConvertTo-MetricCardHtml -Label "Messages Observed" -Value (Format-Number $totalMessages) -Hint "$totalRecords aggregate rows" -Tone "default"),
    (ConvertTo-MetricCardHtml -Label "DMARC Aligned" -Value (Format-Percent $alignedMessages $totalMessages) -Hint "$(Format-Number $alignedMessages) messages" -Tone "good"),
    (ConvertTo-MetricCardHtml -Label "Would Quarantine" -Value (Format-Percent $quarantineMessages $totalMessages) -Hint "$(Format-Number $quarantineMessages) messages" -Tone "danger"),
    (ConvertTo-MetricCardHtml -Label "Critical Risk Volume" -Value (Format-Number $criticalMessages) -Hint "weighted high-risk messages" -Tone "danger"),
    (ConvertTo-MetricCardHtml -Label "Unique Source IPs" -Value (Format-Number $uniqueSourceIps) -Hint "$uniqueReporters reporting orgs" -Tone "default"),
    (ConvertTo-MetricCardHtml -Label "Geo Coverage" -Value $geoCoverage -Hint "$(Format-Number $geoCountries.Count) countries resolved" -Tone "info"),
    (ConvertTo-MetricCardHtml -Label "Policy Domains" -Value (Format-Number $uniquePolicyDomains) -Hint "$(Format-Number $uniqueHeaderFrom) header-from identities" -Tone "default"),
    (ConvertTo-MetricCardHtml -Label "Report Window" -Value ("{0} to {1}" -f (Get-DateLabel $dateRangeStart), (Get-DateLabel $dateRangeEnd)) -Hint (Split-Path -Path $resolvedInputJsonPath -Leaf) -Tone "default")
)

$overviewDistributionHtml = ConvertTo-DistributionHtml -Title "Authentication Posture" -Segments @(
    @{ Label = "Aligned"; Value = [int]$alignedMessages; Tone = "good" },
    @{ Label = "Not Aligned"; Value = [int]$notAlignedMessages; Tone = "danger" }
)

$authDistributionHtml = ConvertTo-DistributionHtml -Title "SPF vs DKIM Outcomes" -Segments @(
    @{ Label = "SPF Pass"; Value = [int]$spfPassMessages; Tone = "good" },
    @{ Label = "SPF Fail"; Value = [int]$spfFailMessages; Tone = "warning" },
    @{ Label = "DKIM Pass"; Value = [int]$dkimPassMessages; Tone = "info" },
    @{ Label = "DKIM Fail"; Value = [int]$dkimFailMessages; Tone = "danger" }
)

$riskDistributionHtml = ConvertTo-DistributionHtml -Title "Risk Band Distribution" -Segments @(
    @{ Label = "Critical"; Value = [int](Get-MessageSum -Rows @($normalized | Where-Object { $_.RiskBand -eq "Critical" })); Tone = "danger" },
    @{ Label = "High"; Value = [int](Get-MessageSum -Rows @($normalized | Where-Object { $_.RiskBand -eq "High" })); Tone = "warning" },
    @{ Label = "Elevated"; Value = [int](Get-MessageSum -Rows @($normalized | Where-Object { $_.RiskBand -eq "Elevated" })); Tone = "info" },
    @{ Label = "Low"; Value = [int](Get-MessageSum -Rows @($normalized | Where-Object { $_.RiskBand -eq "Low" })); Tone = "good" }
)

$policySummaryHtml = ConvertTo-TableHtml -Title "Published Policy Landscape" -Rows $policySummary -MaxRows 20 -Columns @(
    @{ Property = "PolicyP"; Label = "Policy P" },
    @{ Property = "PolicyAdkim"; Label = "ADKIM" },
    @{ Property = "PolicyAspf"; Label = "ASPF" },
    @{ Property = "PolicyPct"; Label = "PCT" },
    @{ Property = "RecordCount"; Label = "Rows"; Formatter = { param($value) Format-Number $value } },
    @{ Property = "MessageCount"; Label = "Messages"; Formatter = { param($value) Format-Number $value } },
    @{ Property = "QuarantineCount"; Label = "Would Quarantine"; Formatter = { param($value) Format-Number $value } }
)

$authMatrixHtml = ConvertTo-TableHtml -Title "Disposition / Alignment Matrix" -Rows $authMatrix -MaxRows 20 -Columns @(
    @{ Property = "DmarcSpf"; Label = "DMARC SPF" },
    @{ Property = "DmarcDkim"; Label = "DMARC DKIM" },
    @{ Property = "DmarcDisposition"; Label = "Disposition" },
    @{ Property = "RecordCount"; Label = "Rows"; Formatter = { param($value) Format-Number $value } },
    @{ Property = "MessageCount"; Label = "Messages"; Formatter = { param($value) Format-Number $value } },
    @{ Property = "QuarantineCount"; Label = "Would Quarantine"; Formatter = { param($value) Format-Number $value } }
)

$riskTableHtml = ConvertTo-TableHtml -Title "Highest Risk Records" -Rows $highRiskRows -MaxRows 40 -Columns @(
    @{ Property = "RiskBand"; Label = "Risk" },
    @{ Property = "RiskScore"; Label = "Score"; Formatter = { param($value) Format-Number $value } },
    @{ Property = "OrgName"; Label = "Reporting Org" },
    @{ Property = "SourceIp"; Label = "Source IP" },
    @{ Property = "SourceIpCount"; Label = "Messages"; Formatter = { param($value) Format-Number $value } },
    @{ Property = "GeoCountry"; Label = "Country" },
    @{ Property = "GeoCity"; Label = "City" },
    @{ Property = "HeaderFrom"; Label = "Header From" },
    @{ Property = "SpfDomain"; Label = "SPF Domain" },
    @{ Property = "DkimDomain"; Label = "DKIM Domain" },
    @{ Property = "DmarcSpf"; Label = "SPF" },
    @{ Property = "DmarcDkim"; Label = "DKIM" }
)

$quarantineTableHtml = ConvertTo-TableHtml -Title "Would Quarantine Under p=quarantine" -Rows $quarantineRows -MaxRows 40 -Columns @(
    @{ Property = "OrgName"; Label = "Reporting Org" },
    @{ Property = "SourceIp"; Label = "Source IP" },
    @{ Property = "SourceIpCount"; Label = "Messages"; Formatter = { param($value) Format-Number $value } },
    @{ Property = "GeoCountry"; Label = "Country" },
    @{ Property = "GeoCity"; Label = "City" },
    @{ Property = "HeaderFrom"; Label = "Header From" },
    @{ Property = "EnvelopeFrom"; Label = "Envelope From" },
    @{ Property = "DkimDomain"; Label = "Dkim Domain" },
    @{ Property = "DkimResult"; Label = "DKIM Result" },
    @{ Property = "SpfResult"; Label = "SPF Result" },
    @{ Property = "RiskBand"; Label = "Risk" }
)

$internationalTableHtml = ConvertTo-TableHtml -Title "International Traffic Spotlight" -Rows $internationalRows -MaxRows 30 -Columns @(
    @{ Property = "GeoCountry"; Label = "Country" },
    @{ Property = "GeoCity"; Label = "City" },
    @{ Property = "GeoOrg"; Label = "Geo Org" },
    @{ Property = "SourceIp"; Label = "Source IP" },
    @{ Property = "SourceIpCount"; Label = "Messages"; Formatter = { param($value) Format-Number $value } },
    @{ Property = "HeaderFrom"; Label = "Header From" },
    @{ Property = "DmarcSpf"; Label = "SPF" },
    @{ Property = "DmarcDkim"; Label = "DKIM" },
    @{ Property = "RiskBand"; Label = "Risk" }
)

$topOrgHtml = ConvertTo-LeaderboardHtml -Title "Top Reporting Organizations" -Rows $orgSummary -LabelProperty "OrgName" -ValueProperty "MessageCount" -MetaProperties @("RecordCount", "QuarantineCount") -MaxRows 12
$topIpHtml = ConvertTo-LeaderboardHtml -Title "Top Source IPs" -Rows $ipSummary -LabelProperty "SourceIp" -ValueProperty "MessageCount" -MetaProperties @("GeoCountry", "GeoCity", "GeoOrg") -MaxRows 12
$topCountryHtml = ConvertTo-LeaderboardHtml -Title "Top GEO Countries" -Rows $countrySummary -LabelProperty "GeoCountry" -ValueProperty "MessageCount" -MetaProperties @("QuarantineCount", "RecordCount") -MaxRows 12
$topCityHtml = ConvertTo-LeaderboardHtml -Title "Top GEO Cities" -Rows $citySummary -LabelProperty "GeoCity" -ValueProperty "MessageCount" -MetaProperties @("GeoRegion", "GeoCountry") -MaxRows 12
$topSpfDomainHtml = ConvertTo-LeaderboardHtml -Title "Top SPF Domains" -Rows $spfDomainSummary -LabelProperty "SpfDomain" -ValueProperty "MessageCount" -MetaProperties @("SpfResult", "QuarantineCount") -MaxRows 12
$topDkimDomainHtml = ConvertTo-LeaderboardHtml -Title "Top DKIM Domains" -Rows $dkimDomainSummary -LabelProperty "DkimDomain" -ValueProperty "MessageCount" -MetaProperties @("DkimResult", "QuarantineCount") -MaxRows 12
$topHeaderFromHtml = ConvertTo-LeaderboardHtml -Title "Top Header-From Identities" -Rows $headerFromSummary -LabelProperty "HeaderFrom" -ValueProperty "MessageCount" -MetaProperties @("QuarantineCount", "RecordCount") -MaxRows 12
$topGeoOrgHtml = ConvertTo-LeaderboardHtml -Title "Top Network Owners / GEO Orgs" -Rows $geoOrgSummary -LabelProperty "GeoOrg" -ValueProperty "MessageCount" -MetaProperties @("GeoCountry", "QuarantineCount") -MaxRows 12
$reportDayHtml = ConvertTo-LeaderboardHtml -Title "Reporting Days by Message Volume" -Rows $reportDaySummary -LabelProperty "ReportDay" -ValueProperty "MessageCount" -MetaProperties @("RecordCount") -MaxRows 14

$overrideHtml = ConvertTo-TableHtml -Title "Policy Override Reasons" -Subtitle "The receiver reported that it applied its own internal policy or exception handling instead of relying only on the published DMARC policy result." -Rows $overrideSummary -MaxRows 20 -Columns @(
    @{ Property = "PolicyOverrideType"; Label = "Override Type" },
    @{ Property = "RecordCount"; Label = "Rows"; Formatter = { param($value) Format-Number $value } },
    @{ Property = "MessageCount"; Label = "Messages"; Formatter = { param($value) Format-Number $value } },
    @{ Property = "QuarantineCount"; Label = "Would Quarantine"; Formatter = { param($value) Format-Number $value } }
)

$inputFileLabel = Split-Path -Path $resolvedInputJsonPath -Leaf
$generatedAt = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

$explorerRows = $normalized | ForEach-Object {
    [ordered]@{
        ReportDay         = $_.ReportDay
        OrgName           = $_.OrgName
        SourceIp          = $_.SourceIp
        SourceIpCount     = $_.SourceIpCount
        GeoCountry        = $_.GeoCountry
        GeoCity           = $_.GeoCity
        GeoOrg            = $_.GeoOrg
        HeaderFrom        = $_.HeaderFrom
        SpfDomain         = $_.SpfDomain
        DkimDomain        = $_.DkimDomain
        DkimSelector      = $_.DkimSelector
        DmarcSpf          = $_.DmarcSpf
        DmarcDkim         = $_.DmarcDkim
        DmarcDisposition  = $_.DmarcDisposition
        DmarcAligned      = $_.DmarcAligned
        WouldQuarantine   = $_.WouldQuarantine
        RiskScore         = $_.RiskScore
        RiskBand          = $_.RiskBand
        PolicyP           = $_.PolicyP
        GeoLookupStatus   = $_.GeoLookupStatus
        GeoTimezone       = $_.GeoTimezone
        ReportId          = $_.ReportId
    }
}

$dashboardDataJson = ($explorerRows | ConvertTo-Json -Depth 10 -Compress)

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>$(ConvertTo-HtmlEncodedText $Title)</title>
  <style>
    :root {
      --bg: #0d1117;
      --bg-2: #131a23;
      --panel: rgba(18, 26, 35, 0.86);
      --panel-strong: rgba(10, 15, 22, 0.94);
      --panel-border: rgba(255,255,255,0.08);
      --text: #edf2f7;
      --muted: #9fb0c4;
      --accent: #ffb454;
      --accent-2: #59c3c3;
      --good: #35c98a;
      --warning: #f7b955;
      --danger: #ff6b6b;
      --info: #7aa2ff;
      --shadow: 0 18px 55px rgba(0,0,0,0.28);
      --radius: 18px;
      --radius-sm: 12px;
    }

    * {
      box-sizing: border-box;
    }

    html {
      scroll-behavior: smooth;
    }

    body {
      margin: 0;
      color: var(--text);
      background:
        radial-gradient(circle at top left, rgba(255,180,84,0.18), transparent 30%),
        radial-gradient(circle at top right, rgba(89,195,195,0.16), transparent 26%),
        linear-gradient(180deg, #091018 0%, #0d1117 35%, #101723 100%);
      font-family: Bahnschrift, "Segoe UI Variable", "Trebuchet MS", sans-serif;
    }

    a {
      color: inherit;
      text-decoration: none;
    }

    .shell {
      display: grid;
      grid-template-columns: 280px 1fr;
      min-height: 100vh;
    }

    .sidebar {
      position: sticky;
      top: 0;
      height: 100vh;
      padding: 28px 20px;
      background: linear-gradient(180deg, rgba(8, 14, 22, 0.96), rgba(10, 18, 27, 0.86));
      border-right: 1px solid var(--panel-border);
      backdrop-filter: blur(18px);
    }

    .brand {
      padding-bottom: 22px;
      border-bottom: 1px solid rgba(255,255,255,0.08);
      margin-bottom: 22px;
    }

    .brand-kicker {
      color: var(--accent);
      letter-spacing: 0.16em;
      text-transform: uppercase;
      font-size: 11px;
      margin-bottom: 10px;
    }

    .brand h1 {
      font-size: 28px;
      line-height: 1.05;
      margin: 0 0 8px 0;
    }

    .brand p {
      color: var(--muted);
      margin: 0;
      font-size: 13px;
      line-height: 1.5;
    }

    .nav-group {
      margin-top: 18px;
    }

    .nav-label {
      color: var(--muted);
      font-size: 11px;
      letter-spacing: 0.14em;
      text-transform: uppercase;
      margin-bottom: 10px;
    }

    .nav-link {
      display: block;
      padding: 11px 12px;
      border-radius: 12px;
      color: #d6e2ef;
      transition: 160ms ease;
      margin-bottom: 6px;
    }

    .nav-link:hover {
      background: rgba(255,255,255,0.06);
      transform: translateX(2px);
    }

    .main {
      padding: 28px;
    }

    .hero {
      position: relative;
      overflow: hidden;
      background:
        linear-gradient(135deg, rgba(255,180,84,0.18), rgba(89,195,195,0.12)),
        rgba(12, 18, 28, 0.92);
      border: 1px solid var(--panel-border);
      border-radius: 28px;
      box-shadow: var(--shadow);
      padding: 30px;
      margin-bottom: 24px;
    }

    .hero::after {
      content: "";
      position: absolute;
      inset: auto -10% -35% auto;
      width: 340px;
      height: 340px;
      background: radial-gradient(circle, rgba(255,180,84,0.24), transparent 60%);
      pointer-events: none;
    }

    .hero-grid {
      display: grid;
      grid-template-columns: 1.45fr 0.95fr;
      gap: 22px;
      position: relative;
      z-index: 1;
    }

    .hero h2 {
      font-size: 42px;
      line-height: 1;
      margin: 0 0 12px 0;
      max-width: 9ch;
    }

    .hero p {
      max-width: 70ch;
      color: #d9e5f2;
      line-height: 1.6;
      margin: 0;
    }

    .hero-meta {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 14px;
    }

    .hero-meta-card {
      padding: 16px;
      border-radius: 16px;
      background: rgba(255,255,255,0.05);
      border: 1px solid rgba(255,255,255,0.08);
    }

    .hero-meta-card .eyebrow {
      color: var(--muted);
      font-size: 11px;
      text-transform: uppercase;
      letter-spacing: 0.12em;
      margin-bottom: 8px;
    }

    .hero-meta-card .value {
      font-size: 17px;
      line-height: 1.35;
    }

    .section {
      margin-top: 26px;
    }

    .section-head {
      display: flex;
      align-items: end;
      justify-content: space-between;
      gap: 12px;
      margin-bottom: 14px;
    }

    .section-head h2 {
      margin: 0;
      font-size: 24px;
    }

    .section-head p {
      margin: 0;
      color: var(--muted);
      max-width: 68ch;
      line-height: 1.5;
    }

    .metrics-grid {
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 14px;
    }

    .metric-card {
      padding: 18px;
      border-radius: var(--radius);
      background: var(--panel);
      border: 1px solid var(--panel-border);
      backdrop-filter: blur(14px);
      box-shadow: var(--shadow);
      min-height: 128px;
    }

    .metric-card.tone-good { border-color: rgba(53,201,138,0.25); }
    .metric-card.tone-danger { border-color: rgba(255,107,107,0.24); }
    .metric-card.tone-info { border-color: rgba(122,162,255,0.24); }

    .metric-label {
      color: var(--muted);
      text-transform: uppercase;
      letter-spacing: 0.12em;
      font-size: 11px;
      margin-bottom: 10px;
    }

    .metric-value {
      font-size: 30px;
      line-height: 1;
      margin-bottom: 10px;
      font-weight: 700;
    }

    .metric-hint {
      color: #c8d4df;
      font-size: 13px;
      line-height: 1.45;
    }

    .three-up,
    .two-up {
      display: grid;
      gap: 14px;
    }

    .three-up { grid-template-columns: repeat(3, minmax(0, 1fr)); }
    .two-up { grid-template-columns: repeat(2, minmax(0, 1fr)); }

    .panel {
      padding: 18px;
      border-radius: var(--radius);
      background: var(--panel);
      border: 1px solid var(--panel-border);
      box-shadow: var(--shadow);
      backdrop-filter: blur(14px);
    }

    .panel-head {
      display: flex;
      justify-content: space-between;
      align-items: end;
      gap: 10px;
      margin-bottom: 16px;
    }

    .panel-head h2 {
      margin: 0;
      font-size: 18px;
    }

    .panel-subtitle {
      margin: 6px 0 0 0;
      color: var(--muted);
      font-size: 13px;
      line-height: 1.45;
      max-width: 64ch;
    }

    .empty-state {
      color: var(--muted);
      padding: 8px 0;
    }

    .stack-chart {
      display: flex;
      overflow: hidden;
      border-radius: 999px;
      height: 18px;
      background: rgba(255,255,255,0.05);
      border: 1px solid rgba(255,255,255,0.06);
      margin-bottom: 14px;
    }

    .stack-segment.tone-good,
    .legend-swatch.tone-good { background: linear-gradient(90deg, #1bb36d, #35c98a); }
    .stack-segment.tone-danger,
    .legend-swatch.tone-danger { background: linear-gradient(90deg, #ee5a6f, #ff6b6b); }
    .stack-segment.tone-warning,
    .legend-swatch.tone-warning { background: linear-gradient(90deg, #f2a93b, #f7b955); }
    .stack-segment.tone-info,
    .legend-swatch.tone-info { background: linear-gradient(90deg, #5e8cff, #7aa2ff); }

    .legend-grid {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 10px 14px;
    }

    .legend-item {
      display: grid;
      grid-template-columns: 14px 1fr auto auto;
      gap: 8px;
      align-items: center;
      color: #d6e1ec;
      font-size: 13px;
    }

    .legend-swatch {
      width: 14px;
      height: 14px;
      border-radius: 999px;
      display: inline-block;
    }

    .legend-label { color: #dfe8f2; }
    .legend-value,
    .legend-percent { color: var(--muted); }

    .leaderboard {
      display: flex;
      flex-direction: column;
      gap: 12px;
    }

    .leader-row-head {
      display: flex;
      justify-content: space-between;
      gap: 12px;
      margin-bottom: 5px;
    }

    .leader-label {
      font-weight: 600;
    }

    .leader-value {
      color: #f3f7fb;
      font-variant-numeric: tabular-nums;
    }

    .leader-meta {
      color: var(--muted);
      font-size: 12px;
      min-height: 18px;
      margin-bottom: 7px;
    }

    .leader-bar-track {
      height: 10px;
      border-radius: 999px;
      overflow: hidden;
      background: rgba(255,255,255,0.05);
      border: 1px solid rgba(255,255,255,0.05);
    }

    .leader-bar {
      height: 100%;
      border-radius: 999px;
      background: linear-gradient(90deg, var(--accent), var(--accent-2));
    }

    .table-wrap {
      overflow: auto;
    }

    .data-table {
      width: 100%;
      border-collapse: collapse;
      font-size: 13px;
    }

    .data-table th,
    .data-table td {
      text-align: left;
      padding: 10px 12px;
      border-bottom: 1px solid rgba(255,255,255,0.08);
      vertical-align: top;
    }

    .data-table th {
      position: sticky;
      top: 0;
      background: rgba(16, 22, 32, 0.96);
      color: #f5f8fb;
      z-index: 1;
    }

    .data-table tbody tr:hover {
      background: rgba(255,255,255,0.04);
    }

    .explorer-controls {
      display: grid;
      grid-template-columns: 1.3fr repeat(5, minmax(0, 1fr));
      gap: 10px;
      margin-bottom: 14px;
    }

    .explorer-controls input,
    .explorer-controls select {
      width: 100%;
      padding: 11px 12px;
      border-radius: 12px;
      border: 1px solid rgba(255,255,255,0.10);
      background: rgba(6, 10, 16, 0.58);
      color: var(--text);
      outline: none;
    }

    .explorer-controls input::placeholder {
      color: #93a6ba;
    }

    .explorer-summary {
      display: flex;
      flex-wrap: wrap;
      gap: 16px;
      color: var(--muted);
      margin-bottom: 12px;
      font-size: 13px;
    }

    .record-grid {
      display: grid;
      grid-template-columns: 1.35fr 0.9fr;
      gap: 14px;
    }

    .detail-panel {
      background: rgba(6, 10, 16, 0.56);
      border: 1px solid rgba(255,255,255,0.08);
      border-radius: 14px;
      padding: 14px;
      min-height: 160px;
    }

    .detail-empty {
      color: var(--muted);
      line-height: 1.5;
    }

    .detail-grid {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 10px 14px;
      font-size: 13px;
    }

    .detail-item {
      padding-bottom: 8px;
      border-bottom: 1px solid rgba(255,255,255,0.07);
    }

    .detail-label {
      color: var(--muted);
      text-transform: uppercase;
      letter-spacing: 0.08em;
      font-size: 11px;
      margin-bottom: 4px;
    }

    .badge {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      padding: 6px 10px;
      border-radius: 999px;
      font-size: 12px;
      background: rgba(255,255,255,0.06);
      border: 1px solid rgba(255,255,255,0.08);
      color: #edf2f7;
    }

    .badge.good { color: #c9ffe2; border-color: rgba(53,201,138,0.25); }
    .badge.danger { color: #ffd5d5; border-color: rgba(255,107,107,0.25); }
    .badge.warning { color: #ffe6b8; border-color: rgba(247,185,85,0.25); }
    .badge.info { color: #d8e3ff; border-color: rgba(122,162,255,0.25); }

    .footer-note {
      margin-top: 24px;
      color: var(--muted);
      font-size: 12px;
      line-height: 1.6;
    }

    @media (max-width: 1240px) {
      .shell {
        grid-template-columns: 1fr;
      }

      .sidebar {
        position: relative;
        height: auto;
      }

      .hero-grid,
      .metrics-grid,
      .three-up,
      .two-up,
      .record-grid,
      .explorer-controls {
        grid-template-columns: 1fr;
      }
    }
  </style>
</head>
<body>
  <div class="shell">
    <aside class="sidebar">
      <div class="brand">
        <div class="brand-kicker">Dmarc Reporter Dashboard</div>
        <h1>DMARC Intelligence</h1>
        <p>Built from $(ConvertTo-HtmlEncodedText $inputFileLabel). Generated $(ConvertTo-HtmlEncodedText $generatedAt).</p>
      </div>

      <div class="nav-group">
        <div class="nav-label">Navigate</div>
        <a class="nav-link" href="#overview">Overview</a>
        <a class="nav-link" href="#posture">Posture</a>
        <a class="nav-link" href="#leaders">Leaders</a>
        <a class="nav-link" href="#domains">Domains</a>
        <a class="nav-link" href="#geography">Geography</a>
        <a class="nav-link" href="#risk">Risk Watch</a>
        <a class="nav-link" href="#explorer">Record Explorer</a>
      </div>
    </aside>

    <main class="main">
      <section class="hero">
        <div class="hero-grid">
          <div>
            <div class="brand-kicker">Dmarc Reporter Dashboard</div>
            <h2>$(ConvertTo-HtmlEncodedText $Title)</h2>
            <p>
              This dashboard blends DMARC policy, authentication outcomes, source infrastructure, geography, and weighted risk scoring into one view.
              It is intended to answer the operational questions quickly: what aligned, what would quarantine, who sent it, from where, and what deserves attention first.
            </p>
          </div>
          <div class="hero-meta">
            <div class="hero-meta-card">
              <div class="eyebrow">Input Dataset</div>
              <div class="value">$(ConvertTo-HtmlEncodedText $inputFileLabel)</div>
            </div>
            <div class="hero-meta-card">
              <div class="eyebrow">Report Window</div>
              <div class="value">$(ConvertTo-HtmlEncodedText (Get-DateLabel $dateRangeStart)) to $(ConvertTo-HtmlEncodedText (Get-DateLabel $dateRangeEnd))</div>
            </div>
            <div class="hero-meta-card">
              <div class="eyebrow">Resolved Countries</div>
              <div class="value">$(ConvertTo-HtmlEncodedText (Format-Number $geoCountries.Count))</div>
            </div>
            <div class="hero-meta-card">
              <div class="eyebrow">Quarantine Candidates</div>
              <div class="value">$(ConvertTo-HtmlEncodedText (Format-Number $quarantineMessages))</div>
            </div>
          </div>
        </div>
      </section>

      <section class="section" id="overview">
        <div class="section-head">
          <div>
            <h2>Executive Overview</h2>
            <p>Core volume, alignment, quarantine exposure, identity breadth, and GEO coverage.</p>
          </div>
        </div>
        <div class="metrics-grid">
          $([string]::Join("`n", @($metricsHtml)))
        </div>
      </section>

      <section class="section" id="posture">
        <div class="section-head">
          <div>
            <h2>Authentication Posture</h2>
            <p>Weighted message distributions for alignment, protocol outcomes, risk, and the published policy landscape seen in your aggregate reports.</p>
          </div>
        </div>
        <div class="three-up">
          $overviewDistributionHtml
          $authDistributionHtml
          $riskDistributionHtml
        </div>
        <div class="two-up" style="margin-top: 14px;">
          $policySummaryHtml
          $authMatrixHtml
        </div>
      </section>

      <section class="section" id="leaders">
        <div class="section-head">
          <div>
            <h2>Volume Leaders</h2>
            <p>Who is reporting the most, which source IPs dominate the traffic, and which date buckets and infrastructure owners account for the largest message populations.</p>
          </div>
        </div>
        <div class="two-up">
          $topOrgHtml
          $topIpHtml
          $topGeoOrgHtml
          $reportDayHtml
        </div>
      </section>

      <section class="section" id="domains">
        <div class="section-head">
          <div>
            <h2>Domain & Identity Analytics</h2>
            <p>Header-from identities, SPF domains, and DKIM domains tend to expose forwarding, third-party sending, or misaligned signing paths quickly.</p>
          </div>
        </div>
        <div class="three-up">
          $topHeaderFromHtml
          $topSpfDomainHtml
          $topDkimDomainHtml
        </div>
      </section>

      <section class="section" id="geography">
        <div class="section-head">
          <div>
            <h2>Geography & Network Footprint</h2>
            <p>Resolved GEO and network data can separate expected sender infrastructure from unusual countries, cities, or upstream networks.</p>
          </div>
        </div>
        <div class="two-up">
          $topCountryHtml
          $topCityHtml
        </div>
        <div class="section" style="margin-top: 14px;">
          $internationalTableHtml
        </div>
      </section>

      <section class="section" id="risk">
        <div class="section-head">
          <div>
            <h2>Risk Watch</h2>
            <p>Weighted record scoring prioritizes non-aligned traffic, protocol failures, international paths, and larger volumes so you can triage the highest-signal problems first.</p>
          </div>
        </div>
        <div class="two-up">
          $riskTableHtml
          $quarantineTableHtml
        </div>
        <div class="section" style="margin-top: 14px;">
          $overrideHtml
        </div>
      </section>

      <section class="section" id="explorer">
        <div class="section-head">
          <div>
            <h2>Interactive Record Explorer</h2>
            <p>Search and filter the embedded record set in-browser. Click a row to inspect the key fields without leaving the dashboard.</p>
          </div>
        </div>
        <section class="panel">
          <div class="explorer-controls">
            <input type="search" id="searchBox" placeholder="Search org, IP, domain, hostname, city, network owner">
            <select id="orgFilter"><option value="">All Orgs</option></select>
            <select id="countryFilter"><option value="">All Countries</option></select>
            <select id="alignmentFilter">
              <option value="">All Alignment</option>
              <option value="aligned">Aligned</option>
              <option value="not-aligned">Not aligned</option>
            </select>
            <select id="quarantineFilter">
              <option value="">All Quarantine States</option>
              <option value="yes">Would quarantine</option>
              <option value="no">Would not quarantine</option>
            </select>
            <select id="riskFilter">
              <option value="">All Risk Bands</option>
              <option value="Critical">Critical</option>
              <option value="High">High</option>
              <option value="Elevated">Elevated</option>
              <option value="Low">Low</option>
            </select>
          </div>
          <div class="explorer-summary">
            <span id="summaryRows"></span>
            <span id="summaryMessages"></span>
            <span id="summaryQuarantine"></span>
          </div>
          <div class="record-grid">
            <div class="table-wrap">
              <table class="data-table" id="explorerTable">
                <thead>
                  <tr>
                    <th>Day</th>
                    <th>Org</th>
                    <th>Source IP</th>
                    <th>Messages</th>
                    <th>Country</th>
                    <th>City</th>
                    <th>Header From</th>
                    <th>SPF</th>
                    <th>DKIM</th>
                    <th>Risk</th>
                  </tr>
                </thead>
                <tbody></tbody>
              </table>
            </div>
            <div class="detail-panel" id="detailPanel">
              <div class="detail-empty">Select a row to inspect its detailed fields.</div>
            </div>
          </div>
        </section>
      </section>

      <div class="footer-note">
        Dashboard generated from <strong>$(ConvertTo-HtmlEncodedText $inputFileLabel)</strong>. Risk scoring is heuristic and intended for prioritization, not as a replacement for policy evaluation logic.
      </div>
    </main>
  </div>

  <script>
    const dashboardRows = $dashboardDataJson;

    const state = {
      selectedIndex: null
    };

    const filters = {
      searchBox: document.getElementById('searchBox'),
      orgFilter: document.getElementById('orgFilter'),
      countryFilter: document.getElementById('countryFilter'),
      alignmentFilter: document.getElementById('alignmentFilter'),
      quarantineFilter: document.getElementById('quarantineFilter'),
      riskFilter: document.getElementById('riskFilter')
    };

    function populateSelect(select, values) {
      values
        .filter(Boolean)
        .sort((a, b) => String(a).localeCompare(String(b)))
        .forEach(value => {
          const option = document.createElement('option');
          option.value = value;
          option.textContent = value;
          select.appendChild(option);
        });
    }

    populateSelect(filters.orgFilter, [...new Set(dashboardRows.map(row => row.OrgName))]);
    populateSelect(filters.countryFilter, [...new Set(dashboardRows.map(row => row.GeoCountry))]);

    function rowMatches(row) {
      const search = filters.searchBox.value.trim().toLowerCase();
      const haystack = [
        row.ReportDay, row.OrgName, row.SourceIp, row.GeoCountry, row.GeoCity, row.GeoOrg,
        row.HeaderFrom, row.SpfDomain, row.DkimDomain, row.DkimSelector, row.ReportId
      ].filter(Boolean).join(' | ').toLowerCase();

      if (search && !haystack.includes(search)) {
        return false;
      }

      if (filters.orgFilter.value && row.OrgName !== filters.orgFilter.value) {
        return false;
      }

      if (filters.countryFilter.value && row.GeoCountry !== filters.countryFilter.value) {
        return false;
      }

      if (filters.alignmentFilter.value === 'aligned' && !row.DmarcAligned) {
        return false;
      }

      if (filters.alignmentFilter.value === 'not-aligned' && row.DmarcAligned) {
        return false;
      }

      if (filters.quarantineFilter.value === 'yes' && !row.WouldQuarantine) {
        return false;
      }

      if (filters.quarantineFilter.value === 'no' && row.WouldQuarantine) {
        return false;
      }

      if (filters.riskFilter.value && row.RiskBand !== filters.riskFilter.value) {
        return false;
      }

      return true;
    }

    function badgeClassForRisk(riskBand) {
      if (riskBand === 'Critical') return 'danger';
      if (riskBand === 'High') return 'warning';
      if (riskBand === 'Elevated') return 'info';
      return 'good';
    }

    function renderDetail(row) {
      const panel = document.getElementById('detailPanel');

      if (!row) {
        panel.innerHTML = '<div class="detail-empty">Select a row to inspect its detailed fields.</div>';
        return;
      }

      const items = [
        ['Report Day', row.ReportDay],
        ['Reporting Org', row.OrgName],
        ['Report ID', row.ReportId],
        ['Source IP', row.SourceIp],
        ['Messages', row.SourceIpCount],
        ['Country', row.GeoCountry],
        ['City', row.GeoCity],
        ['Network Owner', row.GeoOrg],
        ['Header From', row.HeaderFrom],
        ['SPF Domain', row.SpfDomain],
        ['DKIM Domain', row.DkimDomain],
        ['DKIM Selector', row.DkimSelector],
        ['DMARC SPF', row.DmarcSpf],
        ['DMARC DKIM', row.DmarcDkim],
        ['Disposition', row.DmarcDisposition],
        ['Aligned', row.DmarcAligned],
        ['Would Quarantine', row.WouldQuarantine],
        ['Policy P', row.PolicyP],
        ['Risk Band', row.RiskBand],
        ['Risk Score', row.RiskScore],
        ['Timezone', row.GeoTimezone],
        ['GEO Lookup', row.GeoLookupStatus]
      ];

      panel.innerHTML =
        '<div style="margin-bottom: 12px;">' +
          '<span class="badge ' + badgeClassForRisk(row.RiskBand) + '">' + row.RiskBand + ' risk</span>' +
          '<span class="badge ' + (row.WouldQuarantine ? 'danger' : 'good') + '" style="margin-left: 8px;">' + (row.WouldQuarantine ? 'Would quarantine' : 'Would not quarantine') + '</span>' +
        '</div>' +
        '<div class="detail-grid">' +
          items.map(([label, value]) =>
            '<div class="detail-item">' +
              '<div class="detail-label">' + label + '</div>' +
              '<div>' + (value ?? '') + '</div>' +
            '</div>'
          ).join('') +
        '</div>';
    }

    function renderExplorer() {
      const rows = dashboardRows.filter(rowMatches).sort((a, b) => {
        if (b.RiskScore !== a.RiskScore) return b.RiskScore - a.RiskScore;
        return b.SourceIpCount - a.SourceIpCount;
      });

      const tbody = document.querySelector('#explorerTable tbody');
      tbody.innerHTML = '';

      rows.forEach((row, index) => {
        const tr = document.createElement('tr');
        tr.innerHTML =
          '<td>' + (row.ReportDay ?? '') + '</td>' +
          '<td>' + (row.OrgName ?? '') + '</td>' +
          '<td>' + (row.SourceIp ?? '') + '</td>' +
          '<td>' + (row.SourceIpCount ?? 0) + '</td>' +
          '<td>' + (row.GeoCountry ?? '') + '</td>' +
          '<td>' + (row.GeoCity ?? '') + '</td>' +
          '<td>' + (row.HeaderFrom ?? '') + '</td>' +
          '<td>' + (row.DmarcSpf ?? '') + '</td>' +
          '<td>' + (row.DmarcDkim ?? '') + '</td>' +
          '<td><span class="badge ' + badgeClassForRisk(row.RiskBand) + '">' + row.RiskBand + ' (' + row.RiskScore + ')</span></td>';

        tr.addEventListener('click', () => {
          state.selectedIndex = index;
          renderDetail(row);
        });

        tbody.appendChild(tr);
      });

      const totalMessages = rows.reduce((sum, row) => sum + (Number(row.SourceIpCount) || 0), 0);
      const quarantineMessages = rows.filter(row => row.WouldQuarantine).reduce((sum, row) => sum + (Number(row.SourceIpCount) || 0), 0);

      document.getElementById('summaryRows').textContent = rows.length + ' rows';
      document.getElementById('summaryMessages').textContent = totalMessages.toLocaleString() + ' messages';
      document.getElementById('summaryQuarantine').textContent = quarantineMessages.toLocaleString() + ' would quarantine';

      if (rows.length === 0) {
        renderDetail(null);
      } else if (state.selectedIndex === null || state.selectedIndex >= rows.length) {
        renderDetail(rows[0]);
      }
    }

    Object.values(filters).forEach(control => {
      control.addEventListener('input', renderExplorer);
      control.addEventListener('change', renderExplorer);
    });

    renderExplorer();
  </script>
</body>
</html>
"@

$html | Set-Content -LiteralPath $OutputPath -Encoding UTF8

[PSCustomObject]@{
    InputJsonPath         = (Resolve-Path -LiteralPath $resolvedInputJsonPath).Path
    OutputPath            = (Resolve-Path -LiteralPath $OutputPath).Path
    OutputDirectory       = (Resolve-Path -LiteralPath $outputDirectory).Path
    TotalRows             = $totalRecords
    TotalMessages         = $totalMessages
    WouldQuarantine       = $quarantineMessages
    GeoCoverageMessages   = $geoResolvedMessages
    UniqueSourceIps       = $uniqueSourceIps
    UniqueCountries       = $geoCountries.Count
}

