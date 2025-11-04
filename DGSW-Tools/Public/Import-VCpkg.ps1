function Import-VCpkg {
    <#
    .SYNOPSIS
    Imports the vcpkg PowerShell module for tab completion.
    .DESCRIPTION
    Loads the vcpkg PowerShell module to enable tab completion and other PowerShell integration
    features for vcpkg, if installed at the default location.
    #>

    [cmdletbinding()]
    param()

    try {
        Import-Module C:\Users\Public\Documents\vcpkg\scripts\posh-vcpkg
    } catch {
        Write-Error 'Could not import vcpkg tab completion'
    }

} #end function
