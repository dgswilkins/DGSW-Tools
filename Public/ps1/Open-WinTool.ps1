function Open-WinTool {
    <#
    .SYNOPSIS
    Opens common Windows administrative tools with saved credentials.
    .DESCRIPTION
    Launches various Windows administrative consoles (such as Active Directory Users and Computers, 
    DNS Manager, Group Policy Management Console, etc.) using saved credentials. Allows you to open
    one or more tools in a new process with the appropriate permissions.
    .PARAMETER ADUC
    Opens Active Directory Users and Computers (dsa.msc).
    .PARAMETER ADDT
    Opens Active Directory Domains and Trusts (domain.msc).
    .PARAMETER DNS
    Opens DNS Manager (dnsmgmt.msc).
    .PARAMETER DSAC
    Opens Active Directory Administrative Center (dsac.exe).
    .PARAMETER GPMC
    Opens Group Policy Management Console (gpmc.msc).
    .EXAMPLE
    Open-WinTool -ADUC
    Opens Active Directory Users and Computers with saved credentials.
    .EXAMPLE
    Open-WinTool -DNS -GPMC
    Opens both DNS Manager and Group Policy Management Console with saved credentials.
    #>

    [CmdletBinding()]
    param(
        [Parameter(ParameterSetName = 'IndividualService')][switch]$ADUC,
        [Parameter(ParameterSetName = 'IndividualService')][switch]$ADDT,
        [Parameter(ParameterSetName = 'IndividualService')][switch]$DNS,
        [Parameter(ParameterSetName = 'IndividualService')][switch]$DSAC,
        [Parameter(ParameterSetName = 'IndividualService')][switch]$GPMC
    )

    switch ($psCmdlet.ParameterSetName) {
        'IndividualService' {
            if (!($ADUC -or $DNS -or $DSAC -or $GPMC -or $ADDT)) {
                Write-Error 'Switch not recognized. Run Get-Help Open-WinTool to see available parameters'
                return
            }
        }

        default {
            Write-Error 'You must choose at least one service to connect to. Run Get-Help Open-WinTool to see available parameters'
            return
        }
    }

    #Gather credentials
    Write-Verbose 'Gathering installed modules. Please wait....'
    $Modules = Get-Module -ListAvailable -Refresh
    if (($Modules | Select-Object -ExpandProperty Name) -notcontains 'RunAs') {
        Write-Error 'RunAs module is not installed'
        return
    } else {
        Import-Module RunAs
    }
    $UserCredentials = Get-SavedCred -FileName 'sa'

    #Connect to ADUC
    if ($ADUC) {
        Write-Verbose 'Opening ADUC'
        try {
            RunAs -netOnly $UserCredentials -program 'mmc' -arguments 'C:\Windows\system32\dsa.msc'
            Write-Verbose 'Successfully opened ADUC'
        } catch {
            Write-Error 'Could not open ADUC.'
        }
    }
    if ($ADDT) {
        Write-Verbose 'Opening ADUC'
        try {
            RunAs -netOnly $UserCredentials -program 'mmc' -arguments 'C:\Windows\system32\domain.msc'
            Write-Verbose 'Successfully opened ADUC'
        } catch {
            Write-Error 'Could not open ADUC.'
        }
    }

    if ($DNS) {
        Write-Verbose 'Opening DNS'
        try {
            RunAs -netOnly $UserCredentials -program 'mmc' -arguments 'C:\Windows\system32\dnsmgmt.msc'
            Write-Verbose 'Successfully opened DNS'
        } catch {
            Write-Error 'Could not open DNS.'
        }
    }

    if ($DSAC) {
        Write-Verbose 'Opening DSAC'
        try {
            RunAs -netOnly $UserCredentials -program 'C:\Windows\system32\dsac.exe'
            Write-Verbose 'Successfully opened DSAC'
        } catch {
            Write-Error 'Could not open DSAC.'
        }
    }

    if ($GPMC) {
        Write-Verbose 'Opening GPMC'
        try {
            RunAs -netOnly $UserCredentials -program 'mmc' -arguments 'C:\Windows\system32\gpmc.msc'
            Write-Verbose 'Successfully opened GPMC'
        } catch {
            Write-Error 'Could not open GPMC.'
        }
    }
}
