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
