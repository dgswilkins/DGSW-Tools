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
