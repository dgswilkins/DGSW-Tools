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
