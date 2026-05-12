function Get-DmarcGeoApiToken {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [string]$ExplicitToken,

        [AllowNull()]
        [string]$EnvironmentVariableName,

        [AllowNull()]
        [string]$SecretName
    )

    if (-not [string]::IsNullOrWhiteSpace($ExplicitToken)) {
        return $ExplicitToken
    }

    if (-not [string]::IsNullOrWhiteSpace($EnvironmentVariableName)) {
        $envValue = [System.Environment]::GetEnvironmentVariable($EnvironmentVariableName)
        if (-not [string]::IsNullOrWhiteSpace($envValue)) {
            return $envValue
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($SecretName) -and (Get-Command Get-Secret -ErrorAction SilentlyContinue)) {
        try {
            $secret = Get-Secret -Name $SecretName -AsPlainText -ErrorAction Stop
            if (-not [string]::IsNullOrWhiteSpace($secret)) {
                return $secret
            }
        }
        catch {
            Write-Verbose "Unable to read SecretManagement secret '$SecretName': $($_.Exception.Message)"
        }
    }

    return $null
}
