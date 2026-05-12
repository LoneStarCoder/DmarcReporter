$repoRoot = Split-Path -Parent $PSScriptRoot
$modulePath = Join-Path $repoRoot 'DMARCReporter.psd1'

Import-Module $modulePath -Force

Describe 'DMARCReporter module' {
    It 'exports the public commands' {
        $commands = (Get-Command -Module DMARCReporter).Name
        ($commands -contains 'Invoke-DMARCReporter') | Should Be $true
        ($commands -contains 'New-DMARCReporterConfig') | Should Be $true
    }

    It 'creates a default config file' {
        $testRoot = Join-Path $env:TEMP ('DMARCReporterTests_' + [guid]::NewGuid().ToString('N'))
        New-Item -Path $testRoot -ItemType Directory -Force | Out-Null
        try {
            $configPath = Join-Path $testRoot 'dmarc.config.json'
            New-DMARCReporterConfig -Path $configPath -MailboxFolder 'Inbox\DMARC' -OutputRoot '.\Runs' | Out-Null
            $config = Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json
            $config.MailboxFolder | Should Be 'Inbox\DMARC'
            $config.OutputRoot | Should Be '.\Runs'
        }
        finally {
            Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'keeps GEO lookup disabled in example config' {
        $config = Get-Content -LiteralPath (Join-Path $repoRoot 'config\config.example.json') -Raw | ConvertFrom-Json
        $config.EnableGeoLookup | Should Be $false
    }
}

Describe 'DMARCReporter private helpers' {
    InModuleScope DMARCReporter {
        It 'normalizes null to an empty array' {
            @(ConvertTo-DmarcArray -Value $null).Count | Should Be 0
        }

        It 'reads optional properties safely' {
            $value = Get-DmarcOptionalPropertyValue -Object ([pscustomobject]@{ Name = 'value' }) -PropertyName 'Name'
            $value | Should Be 'value'
            $missing = Get-DmarcOptionalPropertyValue -Object ([pscustomobject]@{}) -PropertyName 'Missing'
            $missing | Should Be $null
        }
    }
}

Describe 'DMARC XML normalization script' {
    It 'converts fixture XML into normalized JSON' {
        $testRoot = Join-Path $env:TEMP ('DmarcNormalizeTests_' + [guid]::NewGuid().ToString('N'))
        $source = Join-Path $testRoot 'xml'
        $output = Join-Path $testRoot 'mastertable.json'
        New-Item -Path $source -ItemType Directory -Force | Out-Null
        Copy-Item -LiteralPath (Join-Path $repoRoot 'tests\fixtures\sample-dmarc.xml') -Destination (Join-Path $source 'sample-dmarc.xml')

        try {
            & (Join-Path $repoRoot 'Private\PipelineScripts\Process-Dmarc_xml.ps1') -SourceFolder $source -OutputJsonPath $output | Out-Null
            $records = Get-Content -LiteralPath $output -Raw | ConvertFrom-Json
            $records.sourceip | Should Be '192.0.2.10'
            $records.sourceipcount | Should Be 3
            $records.dmarcdomain | Should Be 'example.com'
        }
        finally {
            Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

Describe 'GEO merge script' {
    It 'merges GEO fields by source IP' {
        $testRoot = Join-Path $env:TEMP ('DmarcGeoMergeTests_' + [guid]::NewGuid().ToString('N'))
        New-Item -Path $testRoot -ItemType Directory -Force | Out-Null
        $master = Join-Path $testRoot 'mastertable.json'
        $geo = Join-Path $testRoot 'GEOIP.json'
        $output = Join-Path $testRoot 'mastertable.enriched.json'

        @(
            [pscustomobject]@{
                sourceip = '192.0.2.10'
                sourceipcount = 3
            }
        ) | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $master -Encoding UTF8

        @(
            [pscustomobject]@{
                ip = '192.0.2.10'
                lookup_status = 'ok'
                country = 'US'
                org = 'Example Org'
            }
        ) | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $geo -Encoding UTF8

        try {
            & (Join-Path $repoRoot 'Private\PipelineScripts\Merge-GEOIntoMasterTable.ps1') -MasterTableJsonPath $master -GeoIpJsonPath $geo -OutputJsonPath $output | Out-Null
            $records = Get-Content -LiteralPath $output -Raw | ConvertFrom-Json
            $records.geo_country | Should Be 'US'
            $records.geo_org | Should Be 'Example Org'
            $records.geo_lookup_status | Should Be 'ok'
        }
        finally {
            Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}
