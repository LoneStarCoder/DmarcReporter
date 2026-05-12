@{
    RootModule        = 'DmarcDashboard.psm1'
    ModuleVersion     = '0.1.0'
    GUID              = '4c8d6c5d-4f88-47c8-bcae-0a2181b763d6'
    Author            = 'TPC Group'
    CompanyName       = 'TPC Group'
    Copyright         = '(c) TPC Group. All rights reserved.'
    Description       = 'PowerShell tooling for collecting, normalizing, enriching, and reporting DMARC aggregate reports.'
    PowerShellVersion = '5.1'
    FunctionsToExport = @('Invoke-DmarcDashboard', 'New-DmarcDashboardConfig')
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
    PrivateData       = @{
        PSData = @{
            Tags       = @('DMARC', 'Email', 'Reporting', 'Dashboard')
            ProjectUri = ''
            LicenseUri = ''
        }
    }
}
