function Connect-VCenter {
    <#
    .SYNOPSIS
    Connects to a VMware vCenter server using PowerCLI.
    .DESCRIPTION
    Loads the VMware PowerCLI module if necessary, disconnects any existing vCenter
    sessions, and connects to the specified vCenter server.
    .PARAMETER VC_Server
    The hostname or IP address of the vCenter server to connect to.
    .EXAMPLE
    Connect-VCenter -VC_Server "vcenter01.domain.com"
    #>

    [CmdletBinding()]
    param(
        [string]$VC_Server
    )
    if ( !(Get-Module -Name VMware.VimAutomation.Core -ErrorAction SilentlyContinue) ) {
        . 'C:\Program Files (x86)\VMware\Infrastructure\vSphere PowerCLI\Scripts\Initialize-PowerCLIEnvironment.ps1'
    }
    if ($global:DefaultVIServers.Count -gt 0) {
        Disconnect-VIServer -Server * -Force -Confirm: $false
    }
    try {
        Connect-VIServer $VC_Server -WarningAction SilentlyContinue -ErrorAction stop | Out-Null
        Write-Verbose "Successfully conncted to vCenter $($global:DefaultVIServers.name)"
    } catch {
        Write-Error 'Failed to connect'
    }
}
