function Get-SavedCred() {
    <#
    .SYNOPSIS
    Gets a saved credential
    .DESCRIPTION
    Imports credentials from files stored in the profile location
    .PARAMETER FileName
    defines the name of the file to be retrieved. Do not include the suffix
    .EXAMPLE
    Get-SavedCred sample
    #>

    [CmdletBinding()]
    param(
        [string]$FileName
    )
    $MyProfileLocation = Split-Path -Path $profile
    try {
        $credLoc = "$MyProfileLocation\$FileName.xml"
        $cred = Import-Clixml $credLoc
        Write-Verbose "Cached Credentials retrieved from $credLoc"
        return $cred
    } catch {
        Write-Error 'Could not get Admin Credentials. Ensure the credential file is present and then try again.'
    }
}
