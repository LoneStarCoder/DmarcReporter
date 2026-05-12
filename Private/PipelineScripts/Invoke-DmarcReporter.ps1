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
        [bool]$MarkAsRead,

        [Parameter()]
        [ValidateSet('Skip', 'Overwrite')]
        [string]$ExistingFileAction = 'Skip',

        [Parameter()]
        [bool]$GenerateReports=$false
    )
#endregion

#region Do Work
Write-Information "Attempting to get attachments from emails"
$Get_Dmarc_emails_Results = .\Get-Dmarc_emails.ps1 -Days $Days -EmailFolderPath $EmailFolderPath -MessageFilter $MessageFilter -MarkAsRead $MarkAsRead -ExistingFileAction $ExistingFileAction
Write-Information "Email results"
$Get_Dmarc_emails_Results
if (($Get_Dmarc_emails_Results.SavedAttachments) -ge 1) {

    Write-Information "Extracting XML from attachments"
    .\Extract-Dmarc_reports.ps1

    Write-Information "Processing XML files"
    .\Process-Dmarc_xml.ps1

    if ($GenerateReports) {
     Write-Information "Updating GEO IP cache"
     .\Invoke-GEO_IP_Lookup.ps1

     Write-Information "Merging GEO IP data into master table"
     .\Merge-GEOIntoMasterTable.ps1

     Write-Information "Generating reports"
     .\New-DMarcReport.ps1
    }
}
#endregion
