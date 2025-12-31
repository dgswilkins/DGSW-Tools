#connect to Microsoft services using certificate based authentication
function Connect-CBAService {
    <#
    .SYNOPSIS
    Connects to Microsoft services using certificate based authentication
    .DESCRIPTION
    Connects to various Microsoft services using certificate based authentication
    The following modules are required for each connection type:
    Az: Install-Module -Name Az
    AzureAD: Install-Module -Name AzureAD
    Exchange Online: Install-Module -Name ExchangeOnlineManagement
    Graph: Install-Module -Name Microsoft.Graph
    PnpOnline: Install-Module -Name PnP.PowerShell
    
    .PARAMETER Az
    Connects to Azure using the Az PowerShell module and certificate-based authentication.
    .PARAMETER Subscription
    The Azure subscription ID or name to use for the Az connection.
    .PARAMETER AzureAD
    Connects to Azure Active Directory using certificate-based authentication.
    .PARAMETER Compliance
    Connects to Microsoft 365 Compliance Center using certificate-based authentication.
    .PARAMETER ExchangeOnline
    Connects to Exchange Online using certificate-based authentication.
    .PARAMETER CommandName
    Specifies a list of Exchange Online commands to load when connecting (comma-separated).
    .PARAMETER Graph
    Connects to Microsoft Graph using certificate-based authentication.
    .PARAMETER PnpOnline
    Connects to SharePoint Online (PnP PowerShell) using certificate-based authentication.
    .PARAMETER AllServices
    Connects to all supported Microsoft services using certificate-based authentication.
    .EXAMPLE
    Connect-CBAService -Az -AzureAD -Compliance -ExchangeOnline -Graph -PnpOnline
    #>
    [CmdletBinding()]
    param(
        [Parameter(ParameterSetName = 'AzSubscription')][switch]$Az,
        [Parameter(ParameterSetName = 'AzSubscription')][string]$Subscription,
        [Parameter(ParameterSetName = 'IndividualService')][switch]$AzureAD,
        [Parameter(ParameterSetName = 'IndividualService')][switch]$Compliance,
        [Parameter(ParameterSetName = 'ExoService')][switch]$ExchangeOnline,
        [Parameter(ParameterSetName = 'ExoService')]
        [ValidatePattern('^([a-zA-Z]+-[a-zA-Z]+)(,([a-zA-Z]+-[a-zA-Z]+))*$')]
        [array]$CommandName,
        [Parameter(ParameterSetName = 'IndividualService')][switch]$Graph,
        [Parameter(ParameterSetName = 'IndividualService')][switch]$PnpOnline,
        [Parameter(ParameterSetName = 'AllServices')][switch]$AllServices
    )

    $showBanner = $false
    $noWelcome = $true
    if ($PSBoundParameters.ContainsKey('Verbose') -or $VerbosePreference -eq 'Continue') {
        $showBanner = $true
        $noWelcome = $false 
    }

    switch ($psCmdlet.ParameterSetName) {
        'IndividualService' {
            if (!($Az -or $AzureAD -or $Compliance -or $ExchangeOnline -or $Graph -or $PnpOnline)) {
                throw('You must choose at least one service to connect to. Run Get-Help Connect-CBAService to see available parameters')
            }
        }

        'AzSubscription' {
            if (!($Az)) {
                throw('''Subscription'' can only be used with the ''Az'' parameter. Run Get-Help Connect-CBAService to see available parameters')
            }
        }
        'ExoService' {
            if (!($ExchangeOnline)) {
                throw('''CommandName'' can only be used with the ''ExchangeOnline'' parameter. Run Get-Help Connect-CBAService to see available parameters')
            }
        }

        'AllServices' {
        }

        default {
            throw('You must choose at least one service to connect to. Run Get-Help Connect-CBAService to see available parameters') 
        }
    }

    #Connect to Az
    if ($Az -or $AllServices) {
        Write-Verbose 'Connecting to Az'
        try {
            $AzAdminAppId = $env:AzAdminAppId
            $thumbprint = Get-CertificateThumbprint -SubjectPattern $env:AzAdminCert
            $tenant = $env:O365tenant
            if ($Subscription -eq 'True' -or !$Subscription) {
                Write-Verbose 'No subscription specified, using default subscription'
                $Subscription = $env:AzDefaultSubscription
            } else {
                $Subscription = $Subscription.Trim()
                Write-Verbose "Using subscription: $Subscription"
            }
            
            $connectAzParams = @{
                ApplicationId         = $AzAdminAppId
                CertificateThumbprint = $thumbprint
                Subscription          = $Subscription
                Tenant                = $tenant
            }
            Connect-AzAccount @connectAzParams -ErrorAction Stop
            Write-Verbose 'Successfully connected to Az.'
        } catch {
            Write-Error "Could not connect to Az service.[ $_ ]"
        }
    }
    #Connect to Azure AD
    if ($AzureAD -or $AllServices) {
        Write-Verbose 'Connecting to Azure AD'
        try {
            $tenantID = $env:O365tenantId
            $AzADappId = $env:AzADappId
            $thumbprint = Get-CertificateThumbprint -SubjectPattern $env:AzADCert
            $connectAzureADParams = @{
                ApplicationId         = $AzADappId
                CertificateThumbprint = $thumbprint
                TenantId              = $tenantID
            }
            Connect-AzureAD @connectAzureADParams -ErrorAction Stop
            Write-Verbose 'Successfully connected to Azure AD.'
        } catch {
            Write-Error "Could not connect to Azure AD service.[ $_ ]"
        }
    }
    
    #Connect to Compliance Center
    if ($Compliance -or $AllServices) {
        Write-Verbose 'Connecting to Compliance Center'
        try {
            #ensure the the exchangeonline module is available
            Import-Module ExchangeOnlineManagement -ErrorAction Stop
            $organization = $env:O365tenant
            $ExOappId = $env:ExOappId
            $thumbprint = Get-CertificateThumbprint -SubjectPattern $env:ExoCert
            $connectComplianceParams = @{
                AppId                 = $ExOappId
                CertificateThumbprint = $thumbprint
                Organization          = $organization
                ShowBanner            = $showBanner
            }
            Connect-IPPSSession @connectComplianceParams -ErrorAction Stop
            Write-Verbose 'Successfully connected to Compliance Center.'
        } catch {
            Write-Error "Could not connect to Compliance Center service.[ $_ ]"
        }
    }

    # Connect to Exchange Online
    if ($ExchangeOnline -or $AllServices) {
        if (-not (Get-ConnectionInformation)) {
            Write-Verbose 'Connecting to Exchange Online'
            try {
                #ensure the the exchangeonline module is available
                Import-Module ExchangeOnlineManagement -ErrorAction Stop
                $organization = $env:O365tenant
                $ExOappId = $env:ExOappId
                $thumbprint = Get-CertificateThumbprint -SubjectPattern $env:ExOCert
                if ($CommandName -eq 'True' -or !$CommandName) {
                    Write-Verbose 'No commands specified, loading all commands'
                    $connectExchangeOnlineParams = @{
                        AppId                 = $ExOappId
                        CertificateThumbprint = $thumbprint
                        Organization          = $organization
                        ShowBanner            = $showBanner
                    }
                } else {
                    [string]$CommandString = $CommandName -join ','
                    Write-Verbose "Only loading specified commands: [$CommandString]"
                    $connectExchangeOnlineParams = @{
                        AppId                 = $ExOappId
                        CertificateThumbprint = $thumbprint
                        Organization          = $organization
                        ShowBanner            = $showBanner
                        CommandName           = $CommandString
                    }
                }
                Connect-ExchangeOnline @connectExchangeOnlineParams -ErrorAction Stop
                Write-Verbose 'Successfully connected to Exchange Online.'
            } catch {
                # report the error message
                Write-Error "Could not connect to Exchange Online service.[ $_ ]"
            } 
        } else {
            Write-Verbose 'Already connected to Exchange Online.'
        }
    }   

    # Connect to Graph
    if ($Graph -or $AllServices) {
        if (-not (Get-MgContext)) {
            Write-Verbose 'Connecting to Graph'
            try {
                $thumbprint = Get-CertificateThumbprint -SubjectPattern $env:MSGraphCert
                $GraphAppId = $env:GraphAppId
                $tenantID = $env:O365tenant
                $connectGraphParams = @{
                    ApplicationId         = $GraphAppId
                    CertificateThumbprint = $thumbprint
                    Tenant                = $tenantID
                    NoWelcome             = $noWelcome
                }
                Connect-MgGraph @connectGraphParams -ErrorAction Stop
                Write-Verbose 'Successfully connected to Graph.'
            } catch {
                Write-Error "Could not connect to Graph service.[ $_ ]"
            }
        } else {
            Write-Verbose 'Already connected to Graph.'
        }
    }

    # Connect to PnpOnline
    if ($PnPOnline -or $AllServices) {
        Write-Verbose 'Connecting to PnP Online'
        try {
            $SPAAppId = $env:SPAAppId
            $thumbprint = Get-CertificateThumbprint -SubjectPattern $env:SPACert
            $tenantID = $env:O365tenant
            $siteURL = "https://$($env:O365tenantPrefix)-admin.sharepoint.com"
            $connectPnPOnlineParams = @{
                ClientId   = $SPAAppId
                Tenant     = $tenantID
                Thumbprint = $thumbprint
                Url        = $siteURL
            }
            Connect-PnPOnline @connectPnPOnlineParams -ErrorAction Stop
            Write-Verbose 'Successfully connected to PnP Online.'
        } catch {
            Write-Error "Could not connect to PnP Online service.[ $_ ]"
        }
    }
}
