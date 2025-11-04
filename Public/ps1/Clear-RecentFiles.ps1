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
