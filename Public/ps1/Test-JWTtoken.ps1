function Test-JWTtoken {
    <#
    .SYNOPSIS
    Parses a JWT token and returns the decoded payload as a PowerShell object.
    .DESCRIPTION
    Decodes the header and payload of a JWT (JSON Web Token) and returns the payload as a PowerShell object.
    Only works for access and ID tokens (not refresh tokens).
    .PARAMETER Token
    The JWT token string to parse.
    .EXAMPLE
    Test-JWTtoken -Token $jwt
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Token
    )

    # Validate token format
    if (-not ($Token.Contains('.') -and $Token.StartsWith('eyJ'))) {
        Write-Error 'Invalid token format. Token must be a JWT (header.payload.signature) and start with "eyJ".'
        return
    }

    # Helper function to decode Base64Url
    function DecodeBase64Url {
        param([string]$inURL)
        $b64 = $inURL.Replace('-', '+').Replace('_', '/')
        while ($b64.Length % 4) { $b64 += '=' }
        return [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($b64))
    }

    # Decode header
    $headerPart = $Token.Split('.')[0]
    Write-Verbose "Base64Url header: $headerPart"
    $headerJson = DecodeBase64Url $headerPart
    Write-Verbose "Decoded header: $headerJson"
    $headerObj = $headerJson | ConvertFrom-Json

    # Decode payload
    $payloadPart = $Token.Split('.')[1]
    Write-Verbose "Base64Url payload: $payloadPart"
    $payloadJson = DecodeBase64Url $payloadPart
    Write-Verbose "Decoded payload: $payloadJson"
    $payloadObj = $payloadJson | ConvertFrom-Json

    # Optionally output header as verbose
    Write-Verbose "JWT Header:`n$($headerObj | Format-List | Out-String)"
    Write-Verbose "JWT Payload:`n$($payloadObj | Format-List | Out-String)"

    return $payloadObj
}
