function Remove-OldModules {
    <#
    .SYNOPSIS
    Removes older versions of installed PowerShell modules.
    .DESCRIPTION
    Finds and uninstalls all versions of installed modules except for the latest version.
    Useful for cleaning up disk space and avoiding version conflicts.
    .PARAMETER Force
    If specified, forces removal of old module versions without confirmation.
    .EXAMPLE
    Remove-OldModules
    Prompts for confirmation before uninstalling old module versions.
    .EXAMPLE
    Remove-OldModules -Force
    Uninstalls all old module versions without prompting for confirmation.
    #>

    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param (
        [Switch]$Force
    )

    $Latest = Get-InstalledModule 
    foreach ($module in $Latest) { 
        Write-Verbose "Checking old versions of $($module.Name) [latest is $( $module.Version)]"
        $oldVers = Get-InstalledModule -Name $module.Name -AllVersions | Where-Object { $_.Version -ne $module.Version }
        foreach ($ver in $oldVers) {
            if ($Force -or $PSCmdlet.ShouldProcess("Version '$($ver.Version)' of module '$($ver.Name)'", 'Uninstall-Module')) {
                Write-Verbose "About to perform the operation 'Uninstall-Module' on target 'Version $($ver.Version) of module '$($ver.Name)'"
                Uninstall-Module $ver.Name -Force:$Force -RequiredVersion $ver.Version 
            }
        }
    }
}
