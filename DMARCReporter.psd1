@{
    RootModule        = 'DMARCReporter.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = '4c8d6c5d-4f88-47c8-bcae-0a2181b763d6'
    Author            = 'Brody Kilpatrick'
    CompanyName       = 'Brody Kilpatrick'
    Copyright         = '(c) LonestarCoder.'
    Description       = 'PowerShell tooling for collecting, normalizing, enriching, and reporting DMARC aggregate reports.'
    PowerShellVersion = '5.1'
    FunctionsToExport = @('Invoke-DMARCReporter', 'New-DMARCReporterConfig')
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
    PrivateData       = @{
        PSData = @{
            Tags       = @('DMARC', 'Email', 'Reporting', 'Dashboard')
            ProjectUri = 'https://github.com/LoneStarCoder/DmarcReporter'
            LicenseUri = 'https://github.com/LoneStarCoder/DmarcReporter/blob/main/LICENSE'
        }
    }
}
