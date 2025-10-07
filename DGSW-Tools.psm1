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

function New-SavedCred() {
    <#
    .SYNOPSIS
    Save a new credential
    .DESCRIPTION
    Saves credentials into files stored in the profile location. 
    .PARAMETER FileName
    defines the name of the file to be stored. Do not include the suffix
    .PARAMETER AdminName
    defines the user name to used when creating the credential
    .EXAMPLE
    New-SavedCred sample
    #>

    [CmdletBinding()]
    param(
        [string]$FileName,
        [string]$AdminName
    )
    $MyProfileLocation = Split-Path -Path $profile
    try {
        Get-Credential $AdminName | Export-Clixml "$MyProfileLocation\$FileName.xml"
        Write-Verbose 'Credentials Saved to Cache'
    } catch {
        Write-Error 'Could not save Admin Credentials'
    }
}

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

function Get-StringHash {
    <#
    .SYNOPSIS
    Compute a hash of a string using the specified algorithm.
    .DESCRIPTION
    Supports SHA256, SHA512, SHA3_256, and SHA3_512. Assumes that the string is UTF-8 encoded.
    .PARAMETER String
    The input string to hash.
    .PARAMETER HashName
    The hash algorithm to use: SHA256, SHA512, SHA3_256, or SHA3_512.
    .EXAMPLE
    Get-StringHash -String "hello" -HashName SHA256
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$String,
        [Parameter()][ValidateSet('SHA256', 'SHA512', 'SHA3_256', 'SHA3_512')][string]$HashName = 'SHA256'
    )

    switch ($HashName.ToUpper()) {
        'SHA256' { 
            $algo = [System.Security.Cryptography.SHA256]::Create() 
            Write-Verbose 'Using SHA256 hash algorithm'
        }
        'SHA512' { 
            $algo = [System.Security.Cryptography.SHA512]::Create() 
            Write-Verbose 'Using SHA512 hash algorithm'
        }
        'SHA3_256' {
            try {
                $algo = [System.Security.Cryptography.SHA3_256]::Create()
                Write-Verbose 'Using SHA3_256 hash algorithm'
            } catch {
                throw 'SHA3_256 is not supported on this system. Please install a compatible .NET implementation (8 or later).'
            }
        }
        'SHA3_512' {
            try {
                $algo = [System.Security.Cryptography.SHA3_512]::Create()
                Write-Verbose 'Using SHA3_512 hash algorithm'
            } catch {
                throw 'SHA3_512 is not supported on this system. Please install a compatible .NET implementation (8 or later).'
            }
        }
        default { throw "Unsupported hash algorithm: $HashName" }
    }

    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($String)
        Write-Verbose "Computing hash for string: [$String]"
        $hash = $algo.ComputeHash($bytes)
        return ($hash | ForEach-Object { $_.ToString('x2') }) -join ''
    } finally {
        if ($algo) { $algo.Dispose() }
    }
}

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
                ShowBanner            = $false
            }
            Connect-IPPSSession @connectComplianceParams -ErrorAction Stop
            Write-Verbose 'Successfully connected to Compliance Center.'
        } catch {
            Write-Error "Could not connect to Compliance Center service.[ $_ ]"
        }
    }

    # Connect to Exchange Online
    if ($ExchangeOnline -or $AllServices) {
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
                    ShowBanner            = $false
                }
            } else {
                [string]$CommandString = $CommandName -join ','
                Write-Verbose "Only loading specified commands: [$CommandString]"
                $connectExchangeOnlineParams = @{
                    AppId                 = $ExOappId
                    CertificateThumbprint = $thumbprint
                    Organization          = $organization
                    ShowBanner            = $false
                    CommandName           = $CommandString
                }
            }
            Connect-ExchangeOnline @connectExchangeOnlineParams -ErrorAction Stop
            Write-Verbose 'Successfully connected to Exchange Online.'
        } catch {
            # report the error message
            Write-Error "Could not connect to Exchange Online service.[ $_ ]"
        }
    }   

    # Connect to Graph
    if ($Graph -or $AllServices) {
        Write-Verbose 'Connecting to Graph'
        try {
            $thumbprint = Get-CertificateThumbprint -SubjectPattern $env:MSGraphCert
            $GraphAppId = $env:GraphAppId
            $tenantID = $env:O365tenant
            $connectGraphParams = @{
                ApplicationId         = $GraphAppId
                CertificateThumbprint = $thumbprint
                Tenant                = $tenantID
            }
            Connect-MgGraph @connectGraphParams -ErrorAction Stop
            Write-Verbose 'Successfully connected to Graph.'
        } catch {
            Write-Error "Could not connect to Graph service.[ $_ ]"
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

function New-IsoFile {
    <#
    .SYNOPSIS 
    Creates a new .iso file 
    .DESCRIPTION
    The New-IsoFile cmdlet creates a new .iso file containing content from chosen folders 
    .PARAMETER Source
    Specifies the files or folders to include in the ISO image.
    .PARAMETER Path
    The output path for the new ISO file.
    .PARAMETER BootFile
    The path to a boot image file to make the ISO bootable (optional).
    .PARAMETER Media
    The media type for the ISO (e.g., DVDPLUSRW, BDR, etc.).
    .PARAMETER Title
    The volume label/title for the ISO image.
    .PARAMETER Force
    Overwrites the target ISO file if it already exists.
    .PARAMETER FromClipboard
    If specified, uses files/folders currently on the clipboard as the source.
    .EXAMPLE
    New-IsoFile "c:\tools","c:Downloads\utils"
    This command creates a .iso file in $env:temp folder (default location) that contains
    c:\tools and c:\downloads\utils folders. The folders themselves are included at the root of the .iso image. 
    .EXAMPLE
    New-IsoFile -FromClipboard -Verbose
    Before running this command, select and copy (Ctrl-C) files/folders in Explorer first. 
    .EXAMPLE
    dir c:\WinPE | New-IsoFile -Path c:\temp\WinPE.iso -BootFile "${env:ProgramFiles(x86)}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\efisys.bin" -Media DVDPLUSR -Title "WinPE"
    This command creates a bootable .iso file containing the content from c:\WinPE folder, but the
    folder itself isn't included. Boot file etfsboot.com can be found in Windows ADK. 
    Refer to IMAPI_MEDIA_PHYSICAL_TYPE enumeration for possible media types:
    https://learn.microsoft.com/en-us/windows/win32/api/imapi2/ne-imapi2-imapi_media_physical_type 
    #>

    [CmdletBinding(DefaultParameterSetName = 'Source')]
    param(
        [parameter(Position = 1, Mandatory = $true, ValueFromPipeline = $true, ParameterSetName = 'Source')]$Source,
        [parameter(Position = 2)][string]$Path = "$env:temp\$((Get-Date).ToString('yyyyMMdd-HHmmss.ffff')).iso",
        [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })][string]$BootFile = $null,
        [ValidateSet('CDR', 'CDRW', 'DVDRAM', 'DVDPLUSR', 'DVDPLUSRW', 'DVDPLUSR_DUALLAYER', 'DVDDASHR', 'DVDDASHRW', 'DVDDASHR_DUALLAYER', 'DISK', 'DVDPLUSRW_DUALLAYER', 'BDR', 'BDRE')][string] $Media = 'DVDPLUSRW_DUALLAYER',
        [string]$Title = (Get-Date).ToString('yyyyMMdd-HHmmss.ffff'),
        [switch]$Force,
        [parameter(ParameterSetName = 'Clipboard')][switch]$FromClipboard
    )

    begin {
        ($cp = New-Object System.CodeDom.Compiler.CompilerParameters).CompilerOptions = '/unsafe'
        if (!('ISOFile' -as [type])) {
            Add-Type -CompilerParameters $cp -TypeDefinition @'
public class ISOFile
{
  public unsafe static void Create(string Path, object Stream, int BlockSize, int TotalBlocks)
  {
    int bytes = 0;
    byte[] buf = new byte[BlockSize];
    var ptr = (System.IntPtr)(&bytes);
    var o = System.IO.File.OpenWrite(Path);
    var i = Stream as System.Runtime.InteropServices.ComTypes.IStream;

    if (o != null) {
      while (TotalBlocks-- > 0) {
        i.Read(buf, BlockSize, ptr); o.Write(buf, 0, bytes);
      }
      o.Flush(); o.Close();
    }
  }
}
'@
        }

        if ($BootFile) {
            if ('BDR', 'BDRE' -contains $Media) { 
                Write-Warning "Bootable image doesn't seem to work with media type $Media"
            }
            ($Stream = New-Object -ComObject ADODB.Stream -Property @{Type = 1 }).Open()  # adFileTypeBinary
            $Stream.LoadFromFile((Get-Item -LiteralPath $BootFile).Fullname)
            ($Boot = New-Object -ComObject IMAPI2FS.BootOptions).AssignBootImage($Stream)
        }

        $MediaType = @('UNKNOWN', 'CDROM', 'CDR', 'CDRW', 'DVDROM', 'DVDRAM', 'DVDPLUSR', 'DVDPLUSRW', 'DVDPLUSR_DUALLAYER', 'DVDDASHR', 'DVDDASHRW', 'DVDDASHR_DUALLAYER', 'DISK', 'DVDPLUSRW_DUALLAYER', 'HDDVDROM', 'HDDVDR', 'HDDVDRAM', 'BDROM', 'BDR', 'BDRE')

        Write-Verbose -Message "Selected media type is $Media with value $($MediaType.IndexOf($Media))"
        ($Image = New-Object -com IMAPI2FS.MsftFileSystemImage -Property @{VolumeName = $Title }).ChooseImageDefaultsForMediaType($MediaType.IndexOf($Media))

        if (!($Target = New-Item -Path $Path -ItemType File -Force:$Force -ErrorAction SilentlyContinue)) {
            Write-Error -Message "Cannot create file $Path. Use -Force parameter to overwrite if the target file already exists."; break 
        }
    }

    process {
        if ($FromClipboard) {
            if ($PSVersionTable.PSVersion.Major -lt 5) {
                Write-Error -Message 'The -FromClipboard parameter is only supported on PowerShell v5 or higher'; break 
            }
            $Source = Get-Clipboard -Format FileDropList
        }

        foreach ($item in $Source) {
            if ($item -isnot [System.IO.FileInfo] -and $item -isnot [System.IO.DirectoryInfo]) {
                $item = Get-Item -LiteralPath $item
            }

            if ($item) {
                Write-Verbose -Message "Adding item to the target image: $($item.FullName)"
                try {
                    $Image.Root.AddTree($item.FullName, $true) 
                } catch {
                    Write-Error -Message ($_.Exception.Message.Trim() + ' Try a different media type.') 
                }
            }
        }
    }

    end {
        if ($Boot) {
            $Image.BootImageOptions = $Boot 
        }
        $Result = $Image.CreateResultImage()
        [ISOFile]::Create($Target.FullName, $Result.ImageStream, $Result.BlockSize, $Result.TotalBlocks)
        Write-Verbose -Message "Target image ($($Target.FullName)) has been created"
        $Target
    }
}

function Optimize-MyVM {
    <#
    .SYNOPSIS
    Optimizes Hyper-V Disks 
    .DESCRIPTION
    The Optimize-MyVM cmdlet removes all snapshots from the Vm and then compacts the Disks
    .PARAMETER VMName
    The name of the Hyper-V virtual machine to optimize.
    .PARAMETER NoCompact
    If specified, skips the disk compaction step and only removes snapshots.
    .EXAMPLE
    Optimize-MyVM -VMName testVM
    #>

    [CmdletBinding()]
    param (
        [string]$VMName,
        [switch]$NoCompact
    )

    try {
        gsudo cache on
    } catch {
        Write-Warning 'Gsudo not cached!'
        return
    }

    try {
        Hyper-V\Get-VM $VMName -ErrorAction Stop
        Write-Verbose "Operating on $VMName"
    } catch {
        Write-Warning "Virtual machine $VMName not found"
        return
    }

    $snapshots = Hyper-V\Get-VMSnapshot -VMName $VMName
    if ($snapshots) {
        Write-Progress -Activity 'Waiting for all checkpoints to be deleted...'
        $snapshots | Hyper-V\Remove-VMSnapshot
        $counter = 0
        do {
            $status = (Hyper-V\Get-VM -Name $VMName).status
            $counter += 1
            if (($counter % 10) -eq 1) {
                Write-Debug "Stage 1 [$status] [$counter]"
                Write-Progress -Activity 'Waiting for checkpoint to be deleted...' -PercentComplete ($counter % 100)
                $Disks = Hyper-V\Get-VMHardDiskDrive -VMName $VMName
                if ($Disks) {
                    $disksMerged = $true
                    foreach ($disk in $disks) {
                        $DName = $disk.path
                        Write-Debug "Disk: [$DName]"
                        if ($DName -like '*.avhdx') {
                            Write-Debug 'Stage 1 - disk not merged'
                            $disksMerged = $false
                        }
                    }
                    if ($disksMerged) {
                        Write-Debug 'Stage 1 disks merged'
                        break
                    }
                }
            }
        } while (($counter -le 1500) -and ($status -eq 'Operating normally'))
        Write-Debug "Stage 2 [$counter]"
        $counter = 0
        do {
            $status = (Hyper-V\Get-VM -Name $VMName).status
            $counter += 1
            if (($counter % 10) -eq 0) {
                Write-Debug "Stage 2 [$status] [$counter]"
                Write-Progress -Activity "Removing snapshot Phase 2 [$status]" -PercentComplete ($counter % 100)
            }
        } while ($status -ne 'Operating normally')
    } else {
        Write-Debug 'No Snapshots'
    }
    if ($NoCompact) {
        Write-Debug 'Compaction disabled'
    } else {
        Write-Progress -Activity 'Compacting'
        $Disks = Hyper-V\Get-VMHardDiskDrive -VMName $VMName
        foreach ($disk in $disks) {
            $DName = $disk.path
            Write-Progress -Activity 'Compacting' -Status "Disk $DName"
            $drive = gsudo "(Hyper-V\Mount-VHD '$DName' -PassThru | Get-Disk | Get-Partition | Get-Volume).DriveLetter"
            gsudo "Optimize-Volume -DriveLetter $drive -Analyze -Defrag | Out-Null"
            gsudo "Optimize-Volume -DriveLetter $drive -SlabConsolidate | Out-Null"
            gsudo "Optimize-Volume -DriveLetter $drive -Analyze -Retrim | Out-Null"
            gsudo "Hyper-V\Dismount-VHD '$DName'"
        }
    }
    $Disks = Hyper-V\Get-VMHardDiskDrive -VMName $VMName
    foreach ($disk in $disks) {
        Write-Progress -Activity 'Optimizing' -Status "Disk $DName"
        $DName = $disk.path
        gsudo "Hyper-V\Mount-VHD '$DName' -ReadOnly"
        gsudo "Hyper-V\Optimize-VHD '$DName' -Mode Full"
        gsudo "Hyper-V\Dismount-VHD '$DName'"
    }
    Write-Progress -Activity 'Checkpointing'
    Hyper-V\Checkpoint-VM -Name $VMName -SnapshotName Clean
    Write-Progress -Activity 'Done' -Completed
    #gsudo cache off
}

function wslcompact {
    <#
    .SYNOPSIS
    Compacts the virtual disk of a WSL distribution.
    .DESCRIPTION
    Exports and re-imports the specified (or all) WSL distributions to optimize and
    reduce the size of their ext4.vhdx virtual disk files.
    .PARAMETER distro
    The name of the WSL distribution to compact. If omitted, all distributions are processed.
    .EXAMPLE
    wslcompact -distro Ubuntu
    #>

    [CmdletBinding()]
    param([string]$distro)

    $tmp_folder = "$Env:TEMP\wslcompact"
    mkdir "$tmp_folder" -ErrorAction SilentlyContinue | Out-Null
    Get-ChildItem HKCU:\Software\Microsoft\Windows\CurrentVersion\Lxss\`{* | ForEach-Object {
        $wsl_ = Get-ItemProperty $_.PSPath
        $wsl_distro = $wsl_.DistributionName
        $wsl_path = if ($wsl_.BasePath.StartsWith('\\')) {
            $wsl_.BasePath.Substring(4)
        } else {
            $wsl_.BasePath
        }
        if ( !$distro -or ($distro -eq $wsl_distro) ) {
            Write-Output "Creating optimized $wsl_distro image."
            $size1 = (Get-Item -Path "$wsl_path\ext4.vhdx").Length / 1MB
            wsl --shutdown
            cmd /c "wsl --export ""$wsl_distro"" - | wsl --import wslclean ""$tmp_folder"" -" 
            wsl --shutdown
            Move-Item "$tmp_folder/ext4.vhdx" "$wsl_path" -Force
            wsl --unregister wslclean | Out-Null
            $size2 = (Get-Item -Path "$wsl_path\ext4.vhdx").Length / 1MB
            Write-Verbose "$wsl_distro image file: $wsl_path\ext4.vhdx"
            Write-Verbose "Compacted from $size1 MB to $size2 MB"
        }
    }
    Remove-Item -Recurse -Force "$tmp_folder"
}

function Connect-VMConsole {
    <#
    .SYNOPSIS
    Opens a Hyper-V VM console window for the specified virtual machine.
    .DESCRIPTION
    Launches the Hyper-V VMConnect console for a given virtual machine by name, ID, or
    input object, optionally on a remote computer. Can also start the VM if it is currently
    off. Supports connecting by VM name, VM ID (GUID), or by passing a VM object directly.
    .PARAMETER ComputerName
    The name of the Hyper-V host computer. Defaults to the local computer.
    .PARAMETER Name
    The name of the virtual machine to connect to.
    .PARAMETER Id
    The GUID of the virtual machine to connect to.
    .PARAMETER InputObject
    A Microsoft.HyperV.PowerShell.VirtualMachine object representing the VM.
    .PARAMETER StartVM
    If specified, starts the VM if it is currently off before connecting.
    .EXAMPLE
    Connect-VMConsole -Name "TestVM"
    Opens a console window for the virtual machine named "TestVM" on the local computer.
    .EXAMPLE
    Connect-VMConsole -ComputerName "HV01" -Name "TestVM"
    Opens a console window for "TestVM" on the remote Hyper-V host "HV01".
    .EXAMPLE
    Get-VM -Name "TestVM" | Connect-VMConsole
    Opens a console window for the VM object returned by Get-VM.
    .EXAMPLE
    Connect-VMConsole -Name "TestVM" -StartVM
    Starts "TestVM" if it is off, then opens a console window.
    #>

    [CmdletBinding(DefaultParameterSetName = 'name')]
    param(
        [Parameter(ParameterSetName = 'name')]
        [Alias('cn')]
        [System.String[]]$ComputerName = $env:COMPUTERNAME,
        [Parameter(Position = 0,
            Mandatory, ValueFromPipelineByPropertyName,
            ValueFromPipeline, ParameterSetName = 'name')]
        [Alias('VMName')]
        [System.String]$Name,

        [Parameter(Position = 0,
            Mandatory, ValueFromPipelineByPropertyName,
            ValueFromPipeline, ParameterSetName = 'id')]
        [Alias('VMId', 'Guid')]
        [System.Guid]$Id,

        [Parameter(Position = 0, Mandatory,
            ValueFromPipeline, ParameterSetName = 'inputObject')]
        [Microsoft.HyperV.PowerShell.VirtualMachine]$InputObject,

        [switch]$StartVM
    )

    begin {
        Write-Verbose 'Initializing InstanceCount, InstanceCount = 0'
        $InstanceCount = 0
    }

    process {
        try {
            foreach ($computer in $ComputerName) {
                Write-Verbose "ParameterSetName is '$($PSCmdlet.ParameterSetName)'"
                if ($PSCmdlet.ParameterSetName -eq 'name') {
                    # Get the VM by Id if Name can convert to a guid
                    if ($Name -as [guid]) {
                        Write-Verbose 'Incoming value can cast to guid'
                        $vm = Hyper-V\Get-VM -Id $Name -ErrorAction SilentlyContinue
                    } else {
                        $vm = Hyper-V\Get-VM -Name $Name -ErrorAction SilentlyContinue
                    }
                } elseif ($PSCmdlet.ParameterSetName -eq 'id') {
                    $vm = Hyper-V\Get-VM -Id $Id -ErrorAction SilentlyContinue
                } else {
                    $vm = $InputObject
                }

                if ($vm) {
                    Write-Verbose "Executing 'vmconnect.exe $computer $($vm.Name) -G $($vm.Id) -C $InstanceCount'"
                    vmconnect.exe $computer $vm.Name -G $vm.Id -C $InstanceCount
                } else {
                    Write-Error "Cannot find vm: '$Name'"
                }

                if ($StartVM -and $vm) {
                    if ($vm.State -eq 'off') {
                        Write-Verbose "StartVM was specified and VM state is 'off'. Starting VM '$($vm.Name)'"
                        Start-VM -VM $vm
                    } else {
                        Write-Verbose "Starting VM '$($vm.Name)'. Skipping, VM is not not in 'off' state."
                    }
                }

                $InstanceCount += 1
                Write-Verbose "InstanceCount = $InstanceCount"
            }
        } catch {
            Write-Error $_
        }
    }
}

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

function Show-TerminalColors {
    <#
    .SYNOPSIS
    Displays all possible foreground and background color combinations in the terminal.
    .DESCRIPTION
    Iterates through all available console colors and prints sample text for each
    foreground/background combination, allowing you to preview how colors appear in your terminal.
    #>

    [cmdletbinding()]
    param()

    $colors = [enum]::GetValues([System.ConsoleColor]) |
        Select-Object @{N = 'ColorObject'; E = { $_ } }, @{N = 'ColorName'; E = { if ($_.ToString().substring(0, 3) -eq 'Dar' ) {
                    $_.ToString().Substring(4) + 'DARK' 
                } else {
                    $_.ToString() 
                } } 
        } | Sort-Object Colorname
    foreach ($bgcolor in $colors.ColorObject) { 
        foreach ($fgcolor in $colors.ColorObject) { 
            Write-Host "$fgcolor|" -ForegroundColor $fgcolor -BackgroundColor $bgcolor -NoNewline 
        } 
        Write-Host " on $bgcolor" 
    }
}


function Clear-RecentFiles {
    <#
    .SYNOPSIS
    Clears the list of recent files and unpinned folders from Quick Access in Windows Explorer.
    .DESCRIPTION
    Removes all files from the user's Recent Files, AutomaticDestinations, and CustomDestinations
    folders (except system files), and unpins all folders from Quick Access that are not pinned.
    This helps maintain privacy and declutter the Quick Access and Recent Files lists in Windows Explorer.
    .NOTES
    This was from code found in the comments in this post:
    https://social.technet.microsoft.com/Forums/windows/en-US/8ad4210c-6ca7-48bd-b218-0676bbf8600a/empty-recent-files-list-from-explorer-by-powershell-or-registry-means

    #>

    [cmdletbinding()]
    param()

    Write-Verbose 'Clearing recent files'

    Get-ChildItem $env:APPDATA\Microsoft\Windows\Recent\* -File -Force -Exclude desktop.ini | 
        Remove-Item -Force -ErrorAction SilentlyContinue
    Get-ChildItem $env:APPDATA\Microsoft\Windows\Recent\AutomaticDestinations\* -File -Force -Exclude desktop.ini, f01b4d95cf55d32a.automaticDestinations-ms | 
        Remove-Item -Force -ErrorAction SilentlyContinue
    Get-ChildItem $env:APPDATA\Microsoft\Windows\Recent\CustomDestinations\* -File -Force -Exclude desktop.ini | 
        Remove-Item -Force -ErrorAction SilentlyContinue
    
    # Clear unpinned folders from Quick Access, using the Verbs() method
    $UnpinnedQAFolders = (0, 0)
    while ($UnpinnedQAFolders) {
        $UnpinnedQAFolders = (((New-Object -ComObject Shell.Application).Namespace('shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}').Items() | 
                    Where-Object IsFolder -EQ $true).Verbs() | Where-Object Name -Match 'Remove from Quick access')
        if ($UnpinnedQAFolders) {
            $UnpinnedQAFolders.DoIt() 
        }
    }
    
    Write-Verbose 'Done!'
    #Stop-Process -Name explorer -Force
    
    Remove-Variable UnpinnedQAFolders    
}

function Get-wslLocation {
    <#
    .SYNOPSIS
    Lists installed WSL distributions and their filesystem locations.
    .DESCRIPTION
    Enumerates all Windows Subsystem for Linux (WSL) distributions registered for the
    current user and returns their names and filesystem paths.
    #>

    [cmdletbinding()]
    param()

    Write-Verbose 'Getting WSL installations info'
    [System.Collections.ArrayList]$return = @() 
    $foundItems = Get-ChildItem 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Lxss' -Recurse
    foreach ($item in $foundItems) {
        $properties = Get-ItemProperty "Registry::$($item.name)"
        $object = [PSCustomObject]@{
            Name = $properties.DistributionName
            Path = $properties.BasePath
        }
        $return.Add($object) | Out-Null
    }
    
    Write-Verbose 'Done!'    
    return $return
}

function Switch-DarkModeState {
    <#
    .SYNOPSIS
    Toggles between Windows dark and light theme modes.
    .DESCRIPTION
    Checks the current Windows theme setting and switches between dark and light modes
    by applying the corresponding theme file. Also attempts to close the System Settings
    window if it appears.
    #>

    [cmdletbinding()]
    param()

    $xpath = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize'
    $xname = 'AppsUseLightTheme'
    if ((Get-ItemProperty $xpath -Name AppsUseLightTheme).$xname -eq 1) {
        Start-Process 'C:\Windows\Resources\Themes\dark.theme' -Wait # set the dark theme of choice
    } else {
        Start-Process 'C:\Windows\Resources\Themes\aero.theme' -Wait # set the light theme of choice
    }

    for ( $i = 0; $i -lt 100; $i++ ) {
        $systemSetting = Get-Process | Where-Object { $_.ProcessName -eq 'SystemSettings' }
        if ( $systemSetting ) {
            $systemSetting | ForEach-Object { $_.Kill() } # The SystemSetting window may be displayed for seconds before being killed.
            $i = 100
        }
        Start-Sleep -Milliseconds 50
    } # Does anyone know a better way to wait for the Process to start?
}

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

function Get-CertificateThumbprint {
    <#
    .SYNOPSIS
    Gets the thumbprint of a certificate based on subject pattern
    .DESCRIPTION
    Searches for certificates in the CurrentUser\My store and returns the thumbprint of the first certificate that matches the subject pattern
    .PARAMETER SubjectPattern
    The regex pattern to match against the certificate subject
    .PARAMETER Store
    The certificate store location (defaults to Cert:\CurrentUser\My)
    .EXAMPLE
    Get-CertificateThumbprint -SubjectPattern 'MyCertificate.*'
    .EXAMPLE
    Get-CertificateThumbprint -SubjectPattern 'MyCompany.*' -Store 'Cert:\LocalMachine\My'
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SubjectPattern,
        
        [Parameter()]
        [string]$Store = 'Cert:\CurrentUser\My'
    )
    
    try {
        $certificate = Get-ChildItem $Store | Where-Object { $_.subject -match $SubjectPattern } | Select-Object -First 1
        
        if ($certificate) {
            Write-Verbose "Found certificate with subject: $($certificate.Subject)"
            return $certificate.Thumbprint
        } else {
            Write-Error "No certificate found matching pattern '$SubjectPattern' in store '$Store'"
            return $null
        }
    } catch {
        Write-Error "Error retrieving certificate: $_"
        return $null
    }
}

function Find-OutlookEmails {
    <#
    .SYNOPSIS
    Finds emails older than a specified date in Outlook folders and subfolders and optionally, delete them
    .DESCRIPTION
    Uses Outlook Interop to search through a specified folder and all subfolders
    to find emails older than the specified date.
    .PARAMETER FolderName
    The name of the Outlook folder to search (e.g., "Inbox", "Sent Items")
    .PARAMETER DaysBack
    The number of days back to search for old emails
    .PARAMETER OutputPath
    Optional path to export results to CSV
    .PARAMETER Delete
    If specified, will delete the found emails instead of just reporting them
    .PARAMETER SummaryOnly
    If specified, will only output a summary of the number of emails found
    .EXAMPLE
    .\Find-OldOutlookEmails.ps1 -FolderName "Inbox" -DaysBack 365
    .EXAMPLE
    .\Find-OldOutlookEmails.ps1 -FolderName "Sent Items" -DaysBack 30 -OutputPath "d:\Reports\OldEmails.csv"
    .EXAMPLE
    .\Find-OldOutlookEmails.ps1 -FolderName "Archive" -DaysBack 90 -Delete
    .EXAMPLE
    .\Find-OldOutlookEmails.ps1 -FolderName "Archive" -DaysBack 90 -SummaryOnly
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FolderName,
    
        [Parameter(Mandatory = $true)]
        [int]$DaysBack,
    
        [Parameter(Mandatory = $false)]
        [string]$OutputPath,

        [Parameter(Mandatory = $false)]
        [switch]$Delete,

        [Parameter(Mandatory = $false)]
        [switch]$SummaryOnly
    )

    # Calculate cutoff date from DaysBack
    $CutoffDate = (Get-Date).AddDays(-$DaysBack)

    # Initialize results array
    $oldEmails = [System.Collections.Generic.List[object]]::new()
    $global:totalEmailsProcessed = 0

    try {
        $interop_assembly_location = (Get-ChildItem -Recurse C:\Windows\assembly Microsoft.Office.Interop.Outlook.dll)
        if (-not $interop_assembly_location) {
            Write-Error 'Microsoft.Office.Interop.Outlook.dll not found. Ensure Outlook is installed.'
            exit 1
        }
        Add-Type -Path $interop_assembly_location -ReferencedAssemblies Microsoft.Office.Interop.Outlook.dll
        Write-Verbose 'Microsoft.Office.Interop.Outlook.dll loaded successfully'
        # Create Outlook Application object
        Write-Verbose 'Connecting to Outlook...'
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace('MAPI')
    
        # Get the default store (mailbox)
        $defaultStore = $namespace.DefaultStore
        Write-Verbose "Connected to mailbox: $($defaultStore.DisplayName)"

        # Find the specified folder
        $rootFolder = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox).Parent
        $targetFolder = $null
    
        # Search for the folder by name
        foreach ($folder in $rootFolder.Folders) {
            if ($folder.Name -eq $FolderName) {
                $targetFolder = $folder
                break
            }
        }
    
        if (-not $targetFolder) {
            Write-Error "Folder '$FolderName' not found in mailbox"
            exit 1
        }

        Write-Verbose "Found folder: $($targetFolder.Name)"

        # Function to recursively search folders
        function Search-Folder {
            param($Folder, $Depth = 0)
        
            $indent = '  ' * $Depth
            Write-Verbose "$indent Searching folder: $($Folder.Name) ($($Folder.Items.Count) items)"

            # Search items in current folder
            foreach ($item in $Folder.Items) {
                $global:totalEmailsProcessed++
            
                # Only process mail items
                if ($item.Class -eq 43) {
                    # olMail = 43
                    try {
                        if ($item.ReceivedTime -lt $CutoffDate) {
                            $oldEmails.Add([PSCustomObject]@{
                                    Subject        = $item.Subject
                                    Sender         = $item.SenderName
                                    SenderEmail    = $item.SenderEmailAddress
                                    ReceivedTime   = $item.ReceivedTime
                                    Size           = $item.Size
                                    FolderPath     = $Folder.FolderPath
                                    HasAttachments = $item.Attachments.Count -gt 0
                                    Importance     = $item.Importance
                                    Categories     = $item.Categories
                                })
                        }
                    } catch {
                        Write-Warning "Error processing email on line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"
                    }
                }
            
                # Show progress every 100 items
                if ($global:totalEmailsProcessed % 100 -eq 0) {
                    Write-Verbose "$indent   Processed $global:totalEmailsProcessed emails..."
                }
            }
        
            # Recursively search subfolders
            foreach ($subFolder in $Folder.Folders) {
                Search-Folder -Folder $subFolder -Depth ($Depth + 1)
            }
        }
    
        # Start the search
        Write-Verbose "`nStarting search for emails older than $($CutoffDate.ToString('yyyy-MM-dd'))..."
        Search-Folder -Folder $targetFolder

        # Delete emails if requested
        if ($Delete -and $oldEmails.Count -gt 0) {
            Write-Verbose "`nDeleting $($oldEmails.Count) emails..."
            # Recursive helper to search a folder tree for a folder matching the target FolderPath
            function Find-FolderByPath {
                param(
                    [Parameter(Mandatory = $true)] $ParentFolder,
                    [Parameter(Mandatory = $true)][string] $TargetPath
                )
                Write-Debug "Searching folder: $($ParentFolder.Name)"
                try {
                    Write-Debug "Checking folder: $($ParentFolder.FolderPath)"
                    if ($ParentFolder.FolderPath -eq $TargetPath) {
                        return $ParentFolder
                    }
                } catch {
                    # Some folder objects may not expose FolderPath; ignore and continue
                }

                foreach ($sub in $ParentFolder.Folders) {
                    $found = Find-FolderByPath -ParentFolder $sub -TargetPath $TargetPath
                    if ($found) { return $found }
                }
                return $null
            }

            foreach ($email in $oldEmails) {
                try {
                    # Use recursive search starting from the mailbox root (folderObj)
                    $foundFolder = Find-FolderByPath -ParentFolder $targetFolder -TargetPath $email.FolderPath
                    if ($foundFolder) {
                        foreach ($item in $foundFolder.Items) {
                            if ($item.Class -eq 43 -and $item.Subject -eq $email.Subject -and $item.ReceivedTime -eq $email.ReceivedTime) {
                                try {
                                    Write-Debug "Deleting email: '$($item.Subject)' from folder '$($email.FolderPath)'"
                                    $item.Delete()
                                } catch {
                                    Write-Warning "Failed to delete item: $($_.Exception.Message)"
                                }
                                break
                            }
                        }
                    } else {
                        Write-Verbose "Folder not found for path: $($email.FolderPath)"
                    }
                } catch {
                    Write-Warning "Failed to delete email: $($_.Exception.Message)"
                }
                Write-Verbose "Deleted $($oldEmails.IndexOf($email) + 1) of $($oldEmails.Count) emails..."
            }
        }
    
        # Display results
        if ($SummaryOnly) {
            Write-Output '=== SUMMARY ==='
            Write-Output "Total emails processed: $global:totalEmailsProcessed"
            Write-Output "Emails older than $($CutoffDate.ToString('yyyy-MM-dd')): $($oldEmails.Count)"
        }

        if ($oldEmails.Count -gt 0) {
            # Calculate total size
            if ($SummaryOnly) {
                $totalSize = ($oldEmails | Measure-Object -Property Size -Sum).Sum
                $totalSizeMB = [math]::Round($totalSize / 1MB, 2)

                Write-Output "Total size of old emails: $totalSizeMB MB"
            }
            # Export to CSV if path specified
            if ($OutputPath) {
                $oldEmails | Sort-Object ReceivedTime | Export-Csv -Path $OutputPath -NoTypeInformation
                Write-Verbose "Results exported to: $OutputPath"
            }
    
            # Show summary by folder
            if ($SummaryOnly) {
                Write-Output "`nEmails by folder:"
                $oldEmails | Group-Object FolderPath | Sort-Object Count -Descending | ForEach-Object {
                    Write-Output "  $($_.Name): $($_.Count) emails"
                }
            }
        } else {
            Write-Verbose 'No emails found older than the specified date.'
        }
    } catch {
        Write-Error "Error accessing Outlook: $($_.Exception.Message)"
        Write-Error 'Make sure Outlook is installed and you have appropriate permissions.'
    } finally {
        # Clean up COM objects
        if ($outlook) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }

    Write-Verbose 'Function completed successfully'
    if (-not $SummaryOnly) {
        return $oldEmails
    }
}

function Test-JWTtoken {
    <#
    .SYNOPSIS
    Parses a JWT token and returns the decoded payload as a PowerShell object.
    .DESCRIPTION
    Decodes the header and payload of a JWT (JSON Web Token) and returns the payload as a PowerShell object.
    Only works for access and ID tokens (not refresh tokens).
    .PARAMETER Token
    The JWT token string to parse.
    .EXAMPLE
    Test-JWTtoken -Token $jwt
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Token
    )

    # Validate token format
    if (-not ($Token.Contains('.') -and $Token.StartsWith('eyJ'))) {
        Write-Error 'Invalid token format. Token must be a JWT (header.payload.signature) and start with "eyJ".'
        return
    }

    # Helper function to decode Base64Url
    function DecodeBase64Url {
        param([string]$inURL)
        $b64 = $inURL.Replace('-', '+').Replace('_', '/')
        while ($b64.Length % 4) { $b64 += '=' }
        return [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($b64))
    }

    # Decode header
    $headerPart = $Token.Split('.')[0]
    Write-Verbose "Base64Url header: $headerPart"
    $headerJson = DecodeBase64Url $headerPart
    Write-Verbose "Decoded header: $headerJson"
    $headerObj = $headerJson | ConvertFrom-Json

    # Decode payload
    $payloadPart = $Token.Split('.')[1]
    Write-Verbose "Base64Url payload: $payloadPart"
    $payloadJson = DecodeBase64Url $payloadPart
    Write-Verbose "Decoded payload: $payloadJson"
    $payloadObj = $payloadJson | ConvertFrom-Json

    # Optionally output header as verbose
    Write-Verbose "JWT Header:`n$($headerObj | Format-List | Out-String)"
    Write-Verbose "JWT Payload:`n$($payloadObj | Format-List | Out-String)"

    return $payloadObj
}

function New-myVM {
    <#
    .SYNOPSIS
    Creates a new Hyper-V virtual machine configured for Windows 11 with TPM and Secure Boot requirements.
    .DESCRIPTION
    This function automates the creation of a Generation 2 Hyper-V VM, configures memory, processors, TPM, Secure Boot, and attaches two DVD drives for installation ISOs. It sets up the VM for Windows 11 requirements and disables automatic checkpoints.
    .PARAMETER VMName
    The name of the new virtual machine to create.
    .PARAMETER ISO
    The filename of the Windows 11 ISO to attach to the second DVD drive.
    .EXAMPLE
    New-myVM -VMName 'Win11Test' -ISO 'Win11_Insider.iso'
    Creates a new VM named 'Win11Test' and attaches the specified ISO to the second DVD drive.
    .NOTES
    - Requires Hyper-V and Windows 11 compatible hardware.
    - Assumes WinPE ISO is located at E:\Downloads\Microsoft\OS\WinPE_amd64-2507.iso
    - Assumes Windows 11 ISO is located at E:\Downloads\Microsoft\OS\Windows11\Insiders\$ISO
    - Requires HGS Guardian and Key Protector for TPM configuration.
    #>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory)][string] $VMName,
        [Parameter(Mandatory)][string] $ISO
    )

    # Define VM parameters
    $VMParams = @{
        MemoryStartupBytes = 4GB
        NewVHDPath         = "C:\ProgramData\Microsoft\Windows\Virtual Hard Disks\$VMName.vhdx"
        NewVHDSizeBytes    = 128GB
        Generation         = 2
        SwitchName         = 'Default Switch'
    }

    # Create the VM
    Hyper-V\New-VM -Name $VMName @VMParams

    # Add processors to the VM
    Hyper-V\Set-VMProcessor -VMName $VMName -Count 2

    # Configure dynamic memory
    Hyper-V\Set-VMMemory -VMName $VMName -DynamicMemoryEnabled $true -MinimumBytes 512MB -MaximumBytes 8192MB

    # Make sure the pre-reqs for Win 11 are set
    $HGOwner = Get-HgsGuardian UntrustedGuardian
    $KeyProtector = New-HgsKeyProtector -Owner $HGOwner -AllowUntrustedRoot
    Hyper-V\Set-VMKeyProtector -VMName $VMName -KeyProtector $KeyProtector.RawData
    Hyper-V\Enable-VMTPM -VMName $VMName
    
    # Attach 2 DVD drives to the VM
    Hyper-V\Add-VMDvdDrive -VMName $VMName
    Hyper-V\Add-VMDvdDrive -VMName $VMName

    # Connect the DVD drive to the ISO image
    $dvds = Hyper-V\Get-VMDvdDrive -VMName $VMName
    $dvds[0] | Hyper-V\Set-VMDvdDrive -Path 'E:\Downloads\Microsoft\OS\WinPE_amd64-2507.iso'
    $dvds[1] | Hyper-V\Set-VMDvdDrive -Path "E:\Downloads\Microsoft\OS\Windows11\Insiders\$ISO"

    # Set the boot order
    $disks = Hyper-V\Get-VMHardDiskDrive -VMName $VMName
    #Set-VMFirmware -VMName $VMName -FirstBootDevice ($dvds[0])
    Hyper-V\Set-VMFirmware -VMName $VMName -BootOrder @($dvds[0], $disks[0])

    Hyper-V\Set-VM -Name $VMName -CheckpointType Standard -AutomaticCheckpointsEnabled $false
}
# End of New-myVM function