param(
    [Parameter(Mandatory = $false)]
    [string]$InputJsonPath = ".\mastertable.json",

    [Parameter(Mandatory = $false)]
    [string]$OutputDirectory = ".\DmarcReport"
)

#Set-StrictMode -Version Latest
Set-StrictMode -Off
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

function ConvertTo-JoinedString {
    param(
        [AllowNull()]
        [object]$Value,

        [string]$Separator = "; "
    )

    [array]$items = ConvertTo-Array -Value $Value

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

function ConvertTo-FlatCsvValue {
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [string] -or $Value -is [ValueType]) {
        return $Value
    }

    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        $items = @($Value)

        if ($items.Count -eq 0) {
            return $null
        }

        $flattenedItems = foreach ($item in $items) {
            if ($null -eq $item) {
                continue
            }

            if ($item -is [string] -or $item -is [ValueType]) {
                [string]$item
                continue
            }

            if ($item.PSObject.Properties.Count -gt 0) {
                (
                    $item.PSObject.Properties |
                        ForEach-Object {
                            if ($null -eq $_.Value -or [string]::IsNullOrWhiteSpace([string]$_.Value)) {
                                return
                            }

                            '{0}={1}' -f $_.Name, $_.Value
                        } |
                        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
                ) -join ', '
                continue
            }

            [string]$item
        }

        return ($flattenedItems | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join '; '
    }

    if ($Value.PSObject.Properties.Count -gt 0) {
        return (
            $Value.PSObject.Properties |
                ForEach-Object {
                    if ($null -eq $_.Value -or [string]::IsNullOrWhiteSpace([string]$_.Value)) {
                        return
                    }

                    '{0}={1}' -f $_.Name, $_.Value
                } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        ) -join '; '
    }

    return [string]$Value
}

function New-Summary {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Rows,

        [Parameter(Mandatory = $true)]
        [string[]]$GroupProperties,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $Rows |
        Group-Object -Property $GroupProperties |
        ForEach-Object {
            $first = $_.Group[0]
            $messageCount = ($_.Group | Measure-Object -Property SourceIpCount -Sum).Sum

            $output = [ordered]@{
                ReportName   = $Name
                RecordCount  = $_.Count
                MessageCount = [int]$messageCount
            }

            foreach ($property in $GroupProperties) {
                $output[$property] = $first.$property
            }

            [PSCustomObject]$output
        } |
        Sort-Object -Property MessageCount -Descending
}

if (-not (Test-Path -LiteralPath $InputJsonPath)) {
    throw "Input JSON file not found: $InputJsonPath"
}

if (-not (Test-Path -LiteralPath $OutputDirectory)) {
    New-Item -Path $OutputDirectory -ItemType Directory | Out-Null
}

$rawJson = Get-Content -LiteralPath $InputJsonPath -Raw
$records = $rawJson | ConvertFrom-Json

if ($null -eq $records) {
    throw "No records were loaded from: $InputJsonPath"
}

$normalized = foreach ($record in $records) {
    $sourceIpCount = ConvertTo-IntSafe -Value $record.sourceipcount
    $dkimDomains = @()
    $dkimSelectors = @()
    $dkimResults = @()
    $policyOverrideTypes = @()
    $policyOverrideComments = @()
    [array]$dkimDomains = ConvertTo-Array -Value $record.dkimdomain
    [array]$dkimSelectors = ConvertTo-Array -Value $record.dkimselector
    [array]$dkimResults = ConvertTo-Array -Value $record.dkimresult
    [array]$policyOverrideTypes = ConvertTo-Array -Value $record.policyoverridetype
    [array]$policyOverrideComments = ConvertTo-Array -Value $record.policyoverridecomment

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

    $spfPass = $record.dmarcspf -eq "pass"
    $dkimPass = $record.dmarcdkim -eq "pass"
    $dmarcAligned = $spfPass -or $dkimPass
    $sourceIpVersion = if ([string]::IsNullOrWhiteSpace([string]$record.sourceipversion)) {
        if ([string]::IsNullOrWhiteSpace([string]$record.sourceip)) {
            $null
        }
        elseif ([string]$record.sourceip -like '*:*') {
            'IPv6'
        }
        else {
            'IPv4'
        }
    }
    else {
        [string]$record.sourceipversion
    }

    [PSCustomObject]@{
        ProcessDate       = $record.processdate
        ReportId          = $record.reportid
        ReportEmail       = $record.reportemail
        ReportDateBegin   = ConvertTo-NullableInt -Value $record.reportdatebegin
        ReportDateEnd     = ConvertTo-NullableInt -Value $record.reportdateend
        ReportDateBeginUtc = $record.reportdatebeginutc
        ReportDateEndUtc  = $record.reportdateendutc
        OrgName           = $record.orgname
        DmarcDomain       = $record.dmarcdomain
        PolicyP           = $record.policyp
        PolicySp          = $record.policysp
        PolicyPct         = ConvertTo-NullableInt -Value $record.policypct
        PolicyAdkim       = $record.policyadkim
        PolicyAspf        = $record.policyaspf
        PolicyFo          = $record.policyfo
        SourceIp          = $record.sourceip
        SourceIpVersion   = $sourceIpVersion
        SourceIpCount     = $sourceIpCount
        DmarcDisposition  = $record.dmarcdisposition
        DmarcSpf          = $record.dmarcspf
        DmarcDkim         = $record.dmarcdkim
        PolicyOverrideType = ConvertTo-JoinedString -Value $policyOverrideTypes
        PolicyOverrideComment = ConvertTo-JoinedString -Value $policyOverrideComments
        EnvelopeFrom      = $record.envelopefrom
        HeaderFrom        = $record.headerfrom
        DkimDomain        = ConvertTo-JoinedString -Value $record.dkimdomain
        DkimSelector      = ConvertTo-JoinedString -Value $dkimSelectors
        DkimResult        = ConvertTo-JoinedString -Value $record.dkimresult
        DkimPairs         = ($dkimPairs | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join "; "
        SpfDomain         = $record.spfdomain
        SpfScope          = $record.spfscope
        SpfResult         = $record.spfresult
        DmarcAligned      = $dmarcAligned
        WouldQuarantineUnderPQuarantine = (-not $dmarcAligned)
        QuarantineReason = if ($dmarcAligned) { "DMARC aligned" } else { "DMARC not aligned" }
        BothSpfAndDkimPass = ($spfPass -and $dkimPass)
        BothSpfAndDkimFail = (-not $spfPass -and -not $dkimPass)
    }
}

$detailCsvPath              = Join-Path $OutputDirectory "dmarc-detail-normalized.csv"
$masterTableCsvPath         = Join-Path $OutputDirectory "dmarc-mastertable-raw.csv"
$masterTableFlattenedCsvPath = Join-Path $OutputDirectory "dmarc-mastertable-flattened.csv"
$quarantineSimulationPath   = Join-Path $OutputDirectory "dmarc-quarantine-simulation.csv"
$summaryByOrgPath           = Join-Path $OutputDirectory "dmarc-summary-by-org.csv"
$summaryBySourceIpPath      = Join-Path $OutputDirectory "dmarc-summary-by-sourceip.csv"
$summaryByAuthPath          = Join-Path $OutputDirectory "dmarc-summary-by-auth-result.csv"
$summaryBySpfDomainPath     = Join-Path $OutputDirectory "dmarc-summary-by-spf-domain.csv"
$summaryByDkimDomainPath    = Join-Path $OutputDirectory "dmarc-summary-by-dkim-domain.csv"
$failuresPath               = Join-Path $OutputDirectory "dmarc-auth-failures.csv"
$htmlReportPath             = Join-Path $OutputDirectory "dmarc-report.html"

$normalized |
    Export-Csv -Path $detailCsvPath -NoTypeInformation -Encoding UTF8

$records |
    Export-Csv -Path $masterTableCsvPath -NoTypeInformation -Encoding UTF8

$flattenedMasterTable = foreach ($record in $records) {
    $flattenedRecord = [ordered]@{}

    foreach ($property in $record.PSObject.Properties) {
        $flattenedRecord[$property.Name] = ConvertTo-FlatCsvValue -Value $property.Value
    }

    [PSCustomObject]$flattenedRecord
}

$flattenedMasterTable |
    Export-Csv -Path $masterTableFlattenedCsvPath -NoTypeInformation -Encoding UTF8

$quarantineSimulation = $normalized |
    Select-Object ProcessDate, ReportId, ReportDateBeginUtc, ReportDateEndUtc, OrgName, DmarcDomain, PolicyP, PolicySp, PolicyPct, HeaderFrom, EnvelopeFrom, SourceIp, SourceIpVersion, SourceIpCount, DmarcDisposition, DmarcSpf, DmarcDkim, SpfDomain, SpfResult, DkimDomain, DkimSelector, DkimResult, PolicyOverrideType, PolicyOverrideComment, DmarcAligned, WouldQuarantineUnderPQuarantine, QuarantineReason

$summaryByOrg = New-Summary -Rows $normalized -GroupProperties @("OrgName") -Name "ByOrg"
$summaryBySourceIp = New-Summary -Rows $normalized -GroupProperties @("SourceIp", "OrgName") -Name "BySourceIp"
$summaryByAuth = New-Summary -Rows $normalized -GroupProperties @("DmarcSpf", "DmarcDkim", "DmarcDisposition") -Name "ByAuthResult"
$summaryBySpfDomain = New-Summary -Rows $normalized -GroupProperties @("SpfDomain", "SpfResult", "DmarcSpf") -Name "BySpfDomain"
$summaryByDkimDomain = New-Summary -Rows $normalized -GroupProperties @("DkimDomain", "DkimResult", "DmarcDkim") -Name "ByDkimDomain"

$failures = $normalized |
    Where-Object {
        $_.DmarcSpf -ne "pass" -or
        $_.DmarcDkim -ne "pass" -or
        $_.SpfResult -ne "pass" -or
        $_.DkimResult -match "fail"
    } |
    Sort-Object -Property SourceIpCount -Descending

$summaryByOrg | Export-Csv -Path $summaryByOrgPath -NoTypeInformation -Encoding UTF8
$summaryBySourceIp | Export-Csv -Path $summaryBySourceIpPath -NoTypeInformation -Encoding UTF8
$summaryByAuth | Export-Csv -Path $summaryByAuthPath -NoTypeInformation -Encoding UTF8
$summaryBySpfDomain | Export-Csv -Path $summaryBySpfDomainPath -NoTypeInformation -Encoding UTF8
$summaryByDkimDomain | Export-Csv -Path $summaryByDkimDomainPath -NoTypeInformation -Encoding UTF8
$quarantineSimulation | Export-Csv -Path $quarantineSimulationPath -NoTypeInformation -Encoding UTF8
$failures | Export-Csv -Path $failuresPath -NoTypeInformation -Encoding UTF8

$totalRecords = $normalized.Count
$totalMessages = ($normalized | Measure-Object -Property SourceIpCount -Sum).Sum
$spfPassMessages = ($normalized | Where-Object { $_.DmarcSpf -eq "pass" } | Measure-Object -Property SourceIpCount -Sum).Sum
$spfFailMessages = ($normalized | Where-Object { $_.DmarcSpf -ne "pass" } | Measure-Object -Property SourceIpCount -Sum).Sum
$dkimPassMessages = ($normalized | Where-Object { $_.DmarcDkim -eq "pass" } | Measure-Object -Property SourceIpCount -Sum).Sum
$dkimFailMessages = ($normalized | Where-Object { $_.DmarcDkim -ne "pass" } | Measure-Object -Property SourceIpCount -Sum).Sum
$alignedMessages = ($normalized | Where-Object { $_.DmarcAligned } | Measure-Object -Property SourceIpCount -Sum).Sum
$notAlignedMessages = ($normalized | Where-Object { -not $_.DmarcAligned } | Measure-Object -Property SourceIpCount -Sum).Sum
$wouldQuarantineMessages = ($normalized | Where-Object { $_.WouldQuarantineUnderPQuarantine } | Measure-Object -Property SourceIpCount -Sum).Sum

$topOrgHtml = $summaryByOrg |
    Select-Object -First 20 OrgName, RecordCount, MessageCount |
    ConvertTo-Html -Fragment

$topIpHtml = $summaryBySourceIp |
    Select-Object -First 25 SourceIp, OrgName, RecordCount, MessageCount |
    ConvertTo-Html -Fragment

$authHtml = $summaryByAuth |
    Select-Object DmarcSpf, DmarcDkim, DmarcDisposition, RecordCount, MessageCount |
    ConvertTo-Html -Fragment

$failureHtml = $failures |
    Select-Object -First 50 OrgName, SourceIp, SourceIpCount, DmarcSpf, DmarcDkim, SpfDomain, SpfResult, DkimDomain, DkimResult |
    ConvertTo-Html -Fragment

$generatedAt = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

$html = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>DMARC Report</title>
<style>
body {
    font-family: Segoe UI, Arial, sans-serif;
    margin: 24px;
    color: #222;
}
h1, h2 {
    color: #1f2937;
}
.summary-grid {
    display: grid;
    grid-template-columns: repeat(4, minmax(180px, 1fr));
    gap: 12px;
    margin-bottom: 24px;
}
.card {
    border: 1px solid #d1d5db;
    border-radius: 8px;
    padding: 12px;
    background: #f9fafb;
}
.card .label {
    font-size: 12px;
    color: #4b5563;
    text-transform: uppercase;
}
.card .value {
    font-size: 24px;
    font-weight: 700;
    margin-top: 4px;
}
table {
    border-collapse: collapse;
    width: 100%;
    margin-bottom: 24px;
}
th, td {
    border: 1px solid #d1d5db;
    padding: 6px 8px;
    text-align: left;
    font-size: 13px;
}
th {
    background: #e5e7eb;
}
tr:nth-child(even) {
    background: #f9fafb;
}
.note {
    color: #4b5563;
    font-size: 13px;
}
</style>
</head>
<body>
<h1>DMARC Report</h1>
<p class="note">Generated: $generatedAt</p>

<div class="summary-grid">
    <div class="card"><div class="label">Records</div><div class="value">$totalRecords</div></div>
    <div class="card"><div class="label">Messages</div><div class="value">$totalMessages</div></div>
    <div class="card"><div class="label">SPF Pass Messages</div><div class="value">$spfPassMessages</div></div>
    <div class="card"><div class="label">SPF Fail Messages</div><div class="value">$spfFailMessages</div></div>
    <div class="card"><div class="label">DKIM Pass Messages</div><div class="value">$dkimPassMessages</div></div>
    <div class="card"><div class="label">DKIM Fail Messages</div><div class="value">$dkimFailMessages</div></div>
    <div class="card"><div class="label">DMARC Aligned Messages</div><div class="value">$alignedMessages</div></div>
    <div class="card"><div class="label">Not Aligned Messages</div><div class="value">$notAlignedMessages</div></div>
</div>

<h2>Authentication Summary</h2>
$authHtml

<h2>Top Organizations</h2>
$topOrgHtml

<h2>Top Source IPs</h2>
$topIpHtml

<h2>Top Authentication Failures / Warnings</h2>
$failureHtml

<p class="note">
Output files are in: $OutputDirectory
</p>
</body>
</html>
"@

$html | Set-Content -Path $htmlReportPath -Encoding UTF8

[PSCustomObject]@{
    InputJson                  = (Resolve-Path -LiteralPath $InputJsonPath).Path
    OutputDirectory            = (Resolve-Path -LiteralPath $OutputDirectory).Path
    DetailCsv                  = $detailCsvPath
    MasterTableRawCsv          = $masterTableCsvPath
    MasterTableFlattenedCsv    = $masterTableFlattenedCsvPath
    QuarantineSimulationCsv    = $quarantineSimulationPath
    SummaryByOrgCsv            = $summaryByOrgPath
    SummaryBySourceIpCsv       = $summaryBySourceIpPath
    SummaryByAuthResultCsv     = $summaryByAuthPath
    SummaryBySpfDomainCsv      = $summaryBySpfDomainPath
    SummaryByDkimDomainCsv     = $summaryByDkimDomainPath
    AuthFailuresCsv            = $failuresPath
    HtmlReport                 = $htmlReportPath
    TotalRecords               = $totalRecords
    TotalMessages              = $totalMessages
    SpfPassMessages            = $spfPassMessages
    SpfFailMessages            = $spfFailMessages
    DkimPassMessages           = $dkimPassMessages
    DkimFailMessages           = $dkimFailMessages
    DmarcAlignedMessages       = $alignedMessages
    DmarcNotAlignedMessages    = $notAlignedMessages
    WouldQuarantineMessages    = $wouldQuarantineMessages
}
