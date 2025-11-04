function Get-CertificateThumbprint {
    <#
    .SYNOPSIS
    Gets the thumbprint of a certificate based on subject pattern
    .DESCRIPTION
    Searches for certificates in the CurrentUser\My store and returns the thumbprint of the first certificate that matches the subject pattern
    .PARAMETER SubjectPattern
    The regex pattern to match against the certificate subject
    .PARAMETER Store
    The certificate store location (defaults to Cert:\CurrentUser\My)
    .EXAMPLE
    Get-CertificateThumbprint -SubjectPattern 'MyCertificate.*'
    .EXAMPLE
    Get-CertificateThumbprint -SubjectPattern 'MyCompany.*' -Store 'Cert:\LocalMachine\My'
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SubjectPattern,
        
        [Parameter()]
        [string]$Store = 'Cert:\CurrentUser\My'
    )
    
    try {
        $certificate = Get-ChildItem $Store | Where-Object { $_.subject -match $SubjectPattern } | Select-Object -First 1
        
        if ($certificate) {
            Write-Verbose "Found certificate with subject: $($certificate.Subject)"
            return $certificate.Thumbprint
        } else {
            Write-Error "No certificate found matching pattern '$SubjectPattern' in store '$Store'"
            return $null
        }
    } catch {
        Write-Error "Error retrieving certificate: $_"
        return $null
    }
}
