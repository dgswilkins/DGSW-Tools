function Get-MyPublicIP {
    <#
    .SYNOPSIS
    Retrieves the public IPv4 and IPv6 addresses of the current machine.
    .DESCRIPTION
    Uses external web services to determine the public-facing IPv4 and IPv6 addresses of the system. Returns the results as plain text.
    #>

    [cmdletbinding()]
    param()

    $paramHashIP4 = @{
        uri              = 'https://api.ipify.org'
        DisableKeepAlive = $True
        UseBasicParsing  = $True
        ErrorAction      = 'Stop'
    }
    $paramHashIP6 = @{
        uri              = 'https://api64.ipify.org'
        DisableKeepAlive = $True
        UseBasicParsing  = $True
        ErrorAction      = 'Stop'
    }
    try {
        $request = Invoke-WebRequest @paramHashIP4
        $request.Content.Trim()
        $request = Invoke-WebRequest @paramHashIP6
        $request.Content.Trim()
    } catch {
        throw $_
    }

} #end function
