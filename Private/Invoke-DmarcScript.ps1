function Invoke-DmarcScript {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ScriptName,

        [Parameter()]
        [hashtable]$Parameters = @{}
    )

    $scriptPath = Join-Path -Path (Join-Path -Path $script:ModuleRoot -ChildPath 'Private\PipelineScripts') -ChildPath $ScriptName
    if (-not (Test-Path -LiteralPath $scriptPath -PathType Leaf)) {
        throw "Pipeline script not found: $scriptPath"
    }

    & $scriptPath @Parameters
}
