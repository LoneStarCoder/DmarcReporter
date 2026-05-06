#region Parameters
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateRange(1, 365)]
        [int]$Days,

        [Parameter(Mandatory = $true)]
        [string]$EmailFolderPath,
        #'Inbox\Ignore\dmarcreports'

        [Parameter()]
        [ValidateSet('UnreadOnly', 'All')]
        [string]$MessageFilter = 'UnreadOnly',

        [Parameter()]
        [bool]$MarkAsRead
    )
#endregion

#region Do Work
Write-Host "Attempting to get Attachments from Emails" -ForegroundColor Green
.\Get-Dmarc_emails.ps1 -Days $Days -EmailFolderPath $EmailFolderPath -MessageFilter $MessageFilter -MarkAsRead $MarkAsRead

Write-Host "Extracting xml from Attachments" -ForegroundColor Green
.\Extract-Dmarc_reports.ps1

Write-Host "Processing xml Files"  -ForegroundColor Green
.\Process-Dmarc_xml.ps1

Write-Host "Generating Reports"
.\New-DMarcReport.ps1

#endregion