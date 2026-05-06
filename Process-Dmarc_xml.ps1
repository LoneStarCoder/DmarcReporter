
[array]$xmlfiles = Get-ChildItem -Path ".\dmarc_xml_exports" -Filter "*.xml"

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

function ConvertFrom-UnixEpochToUtcString {
    param(
        [AllowNull()]
        [object]$Value
    )

    $epoch = ConvertTo-NullableInt -Value $Value

    if ($null -eq $epoch) {
        return $null
    }

    return [DateTimeOffset]::FromUnixTimeSeconds($epoch).UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ")
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

    if ($parsedIp.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork) {
        return "IPv4"
    }

    if ($parsedIp.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetworkV6) {
        return "IPv6"
    }

    return $null
}

function Get-DkimAuthRecords {
    param(
        [AllowNull()]
        [object]$DkimNodes
    )

    $records = @()

    foreach ($dkimNode in (ConvertTo-Array -Value $DkimNodes)) {
        if ($null -eq $dkimNode) {
            continue
        }

        $records += [PSCustomObject]@{
            domain   = if ([string]::IsNullOrWhiteSpace([string]$dkimNode.domain)) { $null } else { [string]$dkimNode.domain }
            selector = if ([string]::IsNullOrWhiteSpace([string]$dkimNode.selector)) { $null } else { [string]$dkimNode.selector }
            result   = if ([string]::IsNullOrWhiteSpace([string]$dkimNode.result)) { $null } else { [string]$dkimNode.result }
        }
    }

    return @($records)
}

function Get-PolicyOverrideReasons {
    param(
        [AllowNull()]
        [object]$ReasonNodes
    )

    $reasons = @()

    foreach ($reasonNode in (ConvertTo-Array -Value $ReasonNodes)) {
        if ($null -eq $reasonNode) {
            continue
        }

        $reasons += [PSCustomObject]@{
            type    = if ([string]::IsNullOrWhiteSpace([string]$reasonNode.type)) { $null } else { [string]$reasonNode.type }
            comment = if ([string]::IsNullOrWhiteSpace([string]$reasonNode.comment)) { $null } else { [string]$reasonNode.comment }
        }
    }

    return @($reasons)
}

$xml = @()

foreach ($item in $xmlfiles) {
    if ($item.Extension -eq ".xml") {
        $xml += [xml](Get-Content -Path $item.FullName)
    }
}

$mastertable = @()

foreach ($record in $xml) {

    $reportMetadata = $record.feedback.report_metadata
    $policyPublished = $record.feedback.policy_published
    $reportDateBegin = ConvertTo-NullableInt -Value $reportMetadata.date_range.begin
    $reportDateEnd = ConvertTo-NullableInt -Value $reportMetadata.date_range.end
    $reportDateBeginUtc = ConvertFrom-UnixEpochToUtcString -Value $reportMetadata.date_range.begin
    $reportDateEndUtc = ConvertFrom-UnixEpochToUtcString -Value $reportMetadata.date_range.end
    $processDate = $reportDateBeginUtc

    foreach ($domainrecord in (ConvertTo-Array -Value $record.feedback.record)) {
        $hostname = $null
        $ip = $null
        $sleepcounter = $null

        $ip = $domainrecord.row.source_ip
        $dkimAuthRecords = Get-DkimAuthRecords -DkimNodes $domainrecord.auth_results.dkim
        $policyOverrideReasons = Get-PolicyOverrideReasons -ReasonNodes $domainrecord.row.policy_evaluated.reason
        $sourceIpCount = ConvertTo-NullableInt -Value $domainrecord.row.count
        $policyPct = ConvertTo-NullableInt -Value $policyPublished.pct

        $temptable = [PSCustomObject]@{
            processdate            = $processDate
            reportid               = if ([string]::IsNullOrWhiteSpace([string]$reportMetadata.report_id)) { $null } else { [string]$reportMetadata.report_id }
            reportemail            = if ([string]::IsNullOrWhiteSpace([string]$reportMetadata.email)) { $null } else { [string]$reportMetadata.email }
            reportdatebegin        = $reportDateBegin
            reportdateend          = $reportDateEnd
            reportdatebeginutc     = $reportDateBeginUtc
            reportdateendutc       = $reportDateEndUtc
            orgname                = if ([string]::IsNullOrWhiteSpace([string]$reportMetadata.org_name)) { $null } else { [string]$reportMetadata.org_name }
            dmarcdomain            = if ([string]::IsNullOrWhiteSpace([string]$policyPublished.domain)) { $null } else { [string]$policyPublished.domain }
            policyp                = if ([string]::IsNullOrWhiteSpace([string]$policyPublished.p)) { $null } else { [string]$policyPublished.p }
            policysp               = if ([string]::IsNullOrWhiteSpace([string]$policyPublished.sp)) { $null } else { [string]$policyPublished.sp }
            policypct              = $policyPct
            policyadkim            = if ([string]::IsNullOrWhiteSpace([string]$policyPublished.adkim)) { $null } else { [string]$policyPublished.adkim }
            policyaspf             = if ([string]::IsNullOrWhiteSpace([string]$policyPublished.aspf)) { $null } else { [string]$policyPublished.aspf }
            policyfo               = if ([string]::IsNullOrWhiteSpace([string]$policyPublished.fo)) { $null } else { [string]$policyPublished.fo }
            sourceip               = if ([string]::IsNullOrWhiteSpace([string]$domainrecord.row.source_ip)) { $null } else { [string]$domainrecord.row.source_ip }
            sourceipversion        = Get-IpVersion -IpAddress $domainrecord.row.source_ip
            sourceipcount          = $sourceIpCount
            dmarcdisposition       = if ([string]::IsNullOrWhiteSpace([string]$domainrecord.row.policy_evaluated.disposition)) { $null } else { [string]$domainrecord.row.policy_evaluated.disposition }
            dmarcspf               = if ([string]::IsNullOrWhiteSpace([string]$domainrecord.row.policy_evaluated.spf)) { $null } else { [string]$domainrecord.row.policy_evaluated.spf }
            dmarcdkim              = if ([string]::IsNullOrWhiteSpace([string]$domainrecord.row.policy_evaluated.dkim)) { $null } else { [string]$domainrecord.row.policy_evaluated.dkim }
            policyoverridetype     = @($policyOverrideReasons | ForEach-Object { $_.type } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            policyoverridecomment  = @($policyOverrideReasons | ForEach-Object { $_.comment } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            policyoverridereason   = @($policyOverrideReasons)
            envelopefrom           = if ([string]::IsNullOrWhiteSpace([string]$domainrecord.identifiers.envelope_from)) { $null } else { [string]$domainrecord.identifiers.envelope_from }
            headerfrom             = if ([string]::IsNullOrWhiteSpace([string]$domainrecord.identifiers.header_from)) { $null } else { [string]$domainrecord.identifiers.header_from }
            dkimauth               = @($dkimAuthRecords)
            dkimdomain             = @($dkimAuthRecords | ForEach-Object { $_.domain } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            dkimselector           = @($dkimAuthRecords | ForEach-Object { $_.selector } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            dkimresult             = @($dkimAuthRecords | ForEach-Object { $_.result } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            spfdomain              = if ([string]::IsNullOrWhiteSpace([string]$domainrecord.auth_results.spf.domain)) { $null } else { [string]$domainrecord.auth_results.spf.domain }
            spfscope               = if ([string]::IsNullOrWhiteSpace([string]$domainrecord.auth_results.spf.scope)) { $null } else { [string]$domainrecord.auth_results.spf.scope }
            spfresult              = if ([string]::IsNullOrWhiteSpace([string]$domainrecord.auth_results.spf.result)) { $null } else { [string]$domainrecord.auth_results.spf.result }
        }


        $mastertable += $temptable
    }
}

$mastertable | ConvertTo-Json -Depth 10 | Set-Content -Path '.\mastertable.json' -Encoding UTF8
