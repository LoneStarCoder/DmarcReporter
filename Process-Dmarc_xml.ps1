
[array]$xmlfiles = Get-ChildItem -Path ".\dmarc_xml_exports" -Filter "*.xml"

$xml = @()

    foreach ( $item in $xmlfiles  )
    {
        if ( $item.extension -eq ".xml")
        {
            $xml += [xml](Get-Content -Path $item.fullname)
        }
    }

$xml


$mastertable = @()

foreach ($record in $xml)
{
    #clear itterative variables
    $recordcount = $null
    $geouri = $null


    $recordcount = $record.feedback.record.count
    if (!($recordcount))
    {
        $recordcount = 1
    }

    foreach ($domainrecord in $record.feedback.record)
    {
        
        #clear itterative variables
        $hostname = $null
        $ip = $null
        $sleepcounter = $null

        #start loop
        $ip = $domainrecord.row.source_ip
<#        if ( $ip ) 
        {
            if ( $iptable.ContainsKey($ip))
            {
                [array]$hostname = $iptable.$ip
            }else
            {
                [array]$hostname = (Resolve-DnsName -name $domainrecord.row.source_ip -Type ptr -QuickTimeout -ErrorAction SilentlyContinue).namehost
                if ([string]::IsNullOrEmpty($hostname))
                {
                    [array]$hostname = "hostname is not resolved"
                }
                
                $iptable += @{ $ip = $($hostname[0])}
                
            }
        } #>

        $temptable = new-object psobject

        $temptable | Add-Member -Type NoteProperty -Name "processdate" -Value $date
        $temptable | Add-Member -Type NoteProperty -Name "orgname" -Value $record.feedback.report_metadata.org_name
        $temptable | Add-Member -Type NoteProperty -Name "dmarcdomain" -Value ([array]$record.feedback.policy_published.domain)[0]
        $temptable | Add-Member -Type NoteProperty -Name "sourceip" -Value $domainrecord.row.source_ip
        #$temptable | Add-Member -Type NoteProperty -Name "sourcedomain" -Value $($hostname[0])
        $temptable | Add-Member -Type NoteProperty -Name "sourceipcount" -Value $domainrecord.row.count
        $temptable | Add-Member -Type NoteProperty -Name "dmarcdisposition" -Value $domainrecord.row.policy_evaluated.disposition
        $temptable | Add-Member -Type NoteProperty -Name "dmarcspf" -Value $domainrecord.row.policy_evaluated.spf
        $temptable | Add-Member -Type NoteProperty -Name "dmarcdkim" -Value $domainrecord.row.policy_evaluated.dkim
        $temptable | Add-Member -Type NoteProperty -Name "headerfrom" -Value $domainrecord.identifiers.header_from
        $temptable | Add-Member -Type NoteProperty -Name "dkimdomain" -Value $domainrecord.auth_results.dkim.domain
        $temptable | Add-Member -Type NoteProperty -Name "dkimresult" -Value $domainrecord.auth_results.dkim.result
        $temptable | Add-Member -Type NoteProperty -Name "spfdomain" -Value $domainrecord.auth_results.spf.domain
        $temptable | Add-Member -Type NoteProperty -Name "spfscope" -Value $domainrecord.auth_results.spf.scope
        $temptable | Add-Member -Type NoteProperty -Name "spfresult" -Value $domainrecord.auth_results.spf.result
        
        if ($geolookupenabled.IsPresent)
        {
            #$geouri = "http://freegeoip.net/json/" + $domainrecord.row.source_ip This one is depricated and now needs APIKEYS
            $geouri = "http://ip-api.com/json/" + $domainrecord.row.source_ip
            
            try
            {
                #introduce sleep to reduce calls to below 150 per minute threshold now 120 due to delay
                start-sleep -Milliseconds 500
                $geodata = Invoke-RestMethod -Method Get -Uri $geouri
                
            }catch
            {
            
            }
             
            $temptable | Add-Member -Type NoteProperty -Name "latitude" -Value $geodata.lat
            $temptable | Add-Member -Type NoteProperty -Name "longitude" -Value $geodata.lon
            $temptable | Add-Member -Type NoteProperty -Name "country_name" -Value $geodata.country
            $temptable | Add-Member -Type NoteProperty -Name "region_name" -Value $geodata.regionname
            $temptable | Add-Member -Type NoteProperty -Name "city" -Value $geodata.city

        }
        
        $mastertable += $temptable

    }
}


 #$mastertable | ft -Property *

 $mastertable  | ConvertTo-Json -Depth 10 |  Set-Content -Path '.\mastertable.json' -Encoding UTF8