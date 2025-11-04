function Disconnect-LocalExchange {
    <#
    .SYNOPSIS
    Disconnects from an on-premises Exchange server session.
    .DESCRIPTION
    Removes the current remote PowerShell session to the on-premises Exchange server
    and clears the session environment variable. If not connected, the function does nothing.
    .EXAMPLE
    Disconnect-LocalExchange
    Disconnects the current session from the on-premises Exchange server.
    #>

    #Disconnect from Exchange
    Write-Verbose 'Disconnecting Exchange'
    try {
        Remove-PSSession -Name $env:ExSession
        $env:ExSession = $null
    } catch {
        Write-Error 'Could not disconnect Exchange Online or not connected.'
    }

}

