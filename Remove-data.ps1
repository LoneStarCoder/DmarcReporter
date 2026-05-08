#$Remove = Read-Host "Remove the following yes/no?
#rm .\dmarc_report_emails\*
#rm .\dmarc_xml_exports\*
#rm .\DmarcReport\*
#"

$Remove = "yes"
if ($Remove -eq "yes"){
    Write-host "Removing"
    rm .\dmarc_report_emails\*
    rm .\dmarc_xml_exports\*
    rm .\DmarcReport\*
} else {Write-host "Not Removing"}