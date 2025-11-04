function New-SavedCred() {
    <#
    .SYNOPSIS
    Save a new credential
    .DESCRIPTION
    Saves credentials into files stored in the profile location. 
    .PARAMETER FileName
    defines the name of the file to be stored. Do not include the suffix
    .PARAMETER AdminName
    defines the user name to used when creating the credential
    .EXAMPLE
    New-SavedCred sample
    #>

    [CmdletBinding()]
    param(
        [string]$FileName,
        [string]$AdminName
    )
    $MyProfileLocation = Split-Path -Path $profile
    try {
        Get-Credential $AdminName | Export-Clixml "$MyProfileLocation\$FileName.xml"
        Write-Verbose 'Credentials Saved to Cache'
    } catch {
        Write-Error 'Could not save Admin Credentials'
    }
}
