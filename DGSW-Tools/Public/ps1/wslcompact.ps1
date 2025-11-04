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
