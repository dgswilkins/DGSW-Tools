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