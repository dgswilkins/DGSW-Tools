function Update-SourceAnchor {
    <#
    .SYNOPSIS
    Updates the mS-DS-ConsistencyGUID attribute for a user in Active Directory.
    .DESCRIPTION
    Converts a base64-encoded string to a binary GUID and sets it as the mS-DS-ConsistencyGUID attribute for the specified Active Directory user. This is typically used for source anchor updates in hybrid identity scenarios.
    .PARAMETER User
    The sAMAccountName or distinguished name of the Active Directory user to update.
    .PARAMETER B64String
    The base64-encoded string representing the GUID to set as the source anchor.
    .PARAMETER Credential
    The credentials to use for the Active Directory operation.
    .EXAMPLE
    Update-SourceAnchor -User "jdoe" -B64String "base64string==" -Credential (Get-Credential)
    .NOTES
    https://scripting.up-in-the.cloud/aadc/the-guid-conversion-carousel.html
    #>

    [CmdletBinding()]
    param(
        [Parameter()][string]$User,
        [Parameter()][string]$B64String,
        [Parameter()][pscredential]$Credential
    )
    $hexstring = ([system.convert]::FromBase64String($B64String) | ForEach-Object ToString X2) -join ' '
    $binary = [byte[]] ( -split (($hexstring -replace ' ', '') -replace '..', '0x$& '))
    Set-ADUser $User -Replace @{'mS-DS-ConsistencyGUID' = $binary } -Credential $Credential
}
