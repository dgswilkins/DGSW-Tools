#connect to on premise services
function Connect-LocalExchange {
    <#
    .SYNOPSIS
    Connects to an on-premises Exchange server using saved credentials.
    .DESCRIPTION
    Establishes a remote PowerShell session to an on-premises Exchange server using credentials
    retrieved from a saved credential file. If already connected, the function does nothing.
    Requires the Exchange server hostname to be set in the $env:ExchHost environment variable.
    .EXAMPLE
    Connect-LocalExchange
    Connects to the Exchange server specified by $env:ExchHost using the saved 'sa' credentials.
    #>

    #Get credentials
    $UserCredentials = Get-SavedCred -FileName 'sa'

    # Connect to Exchange
    Write-Verbose "Connecting to Exchange Server $($env:ExchHost)"
    try {
        if ($env:ExSession) {
            Write-Verbose 'Already connected to Exchange'
        } else {
            # Note that https:// is not supported for on-premises Exchange connections
            # https://learn.microsoft.com/en-us/powershell/exchange/connect-to-exchange-servers-using-remote-powershell?view=exchange-ps
            if (!$UserCredentials) {
                Write-Error 'No credentials found. Please run New-SavedCred to create credentials.'
                return
            }
            $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$($env:ExchHost)/powershell" -Credential $UserCredentials -AllowRedirection -Name 'Exchange'
            Import-PSSession $session
            $env:ExSession = $session.Name
            Write-Verbose 'Successfully connected to Exchange.'
        }
    } catch {
        Write-Error 'Could not connect to the Exchange service. '
        Write-Error 'Ensure the credentials are correct and then try again.'
    }
}
