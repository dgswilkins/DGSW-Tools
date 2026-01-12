function Find-GraphEmails {
    <#
    .SYNOPSIS
    Finds emails older than a specified date in folders and subfolders and optionally, delete them
    .DESCRIPTION
    Uses Microsoft Graph API to search through a specified folder and all subfolders
    to find emails older than the specified date.
    .PARAMETER FolderName
    The name of the folder to search (e.g., "Inbox", "Sent Items")
    .PARAMETER DaysBack
    The number of days back to search for old emails
    .PARAMETER OutputPath
    Optional path to export results to CSV
    .PARAMETER Recurse
    If specified, will search all subfolders of the specified folder
    .PARAMETER Delete
    If specified, will delete the found emails instead of just reporting them
    .PARAMETER SummaryOnly
    If specified, will only output a summary of the number of emails found
    .EXAMPLE
    .\Find-GraphEmails.ps1 -FolderName "Inbox" -DaysBack 365
    .EXAMPLE
    .\Find-GraphEmails.ps1 -FolderName "Sent Items" -DaysBack 30 -OutputPath "d:\Reports\OldEmails.csv"
    .EXAMPLE
    .\Find-GraphEmails.ps1 -FolderName "Archive" -DaysBack 90 -Delete
    .EXAMPLE
    .\Find-GraphEmails.ps1 -FolderName "Archive" -DaysBack 90 -SummaryOnly
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Mailbox,

        [Parameter(Mandatory = $false)]
        [string]$FolderName,

        [Parameter(Mandatory = $false)]
        [switch]$Recurse,

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
    $CutoffDate = (Get-Date -Hour 0 -Minute 0 -Second 0 -Millisecond 0).AddDays(-$DaysBack)

    # Ensure Microsoft.Graph module is available
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
        Write-Error 'Microsoft.Graph PowerShell module is not installed. Install it with: Install-Module Microsoft.Graph -Scope CurrentUser'
        return
    }

    # Connect to Graph if not already connected
    $graphConnected = $false
    try {
        if (Get-MgContext -ErrorAction SilentlyContinue) {
            Write-Verbose 'Already connected to Microsoft Graph.'
        } else {
            Write-Verbose 'Connecting to Microsoft Graph...'
            connect-CBAService -Graph -ErrorAction Stop
            $graphConnected = $true
        }
    } catch {
        Write-Error "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
        return
    }

    $oldEmails = [System.Collections.Generic.List[object]]::new()

    # Helper: get full folder path from folder id
    function Get-MailFolderPath {
        param(
            [Parameter(Mandatory = $true)] [string]$UserId,
            [Parameter(Mandatory = $true)] [string]$FolderId
        )
        $parts = @()
        $currentId = $FolderId
        while ($currentId) {
            try {
                $f = Get-MgUserMailFolder -UserId $UserId -MailFolderId $currentId -ErrorAction Stop
            } catch {
                break
            }
            $parts += $f.DisplayName
            $currentId = $f.ParentFolderId
        }
        if ($parts.Count -eq 0) { return '' }
        return ($parts[-1..0] -join '\\')
    }

    # Build OData filter for receivedDateTime
    $cutoffUtc = $CutoffDate.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    $odataFilter = "receivedDateTime lt $cutoffUtc"

    # If FolderName specified, find the folder id(s) and prepare to recurse
    $targetFolderIds = @()
    try {
        Write-Verbose "Enumerating mail folders for mailbox '$Mailbox'..."
        $folders = Get-MgUserMailFolder -UserId $Mailbox -All -ErrorAction Stop
        if ($FolderName) {
            $matched = $folders | Where-Object { $_.DisplayName -eq $FolderName }
            if ($matched.Count -gt 0) {
                $targetFolderIds = $matched | ForEach-Object { [PSCustomObject]@{ Id = $_.Id; DisplayName = $_.DisplayName } }
            } else {
                Write-Warning "Folder '$FolderName' not found for mailbox '$Mailbox'. Searching all folders by name is case-sensitive."
            }
        }
    } catch {
        Write-Warning "Could not enumerate mail folders: $($_.Exception.Message)"
    }

    # Helper: get descendant folder ids (breadth-first) using Graph child-folder calls
    function Get-DescendantFolderIds {
        param(
            [Parameter(Mandatory = $true)][string]$UserId,
            [Parameter(Mandatory = $true)][PSCustomObject[]]$RootIds
        )
        $result = [System.Collections.Generic.List[object]]::new()
        $queue = [System.Collections.Generic.Queue[PSCustomObject[]]]::new()
        foreach ($r in $RootIds) {
            Write-Verbose "Enqueuing root folder: $($r.DisplayName)"
            Write-Verbose "Finding descendants of folder id: $($r.Id)"
            $queue.Enqueue($r)
            $result.Add($r)
        }
        while ($queue.Count -gt 0) {
            $current = $queue.Dequeue()
            try {
                $children = Get-MgUserMailFolderChildFolder -UserId $UserId -MailFolderId $current.Id -All -ErrorAction Stop
            } catch {
                Write-Warning "Failed to enumerate child folders of folder id $($current.Id). : $($_.Exception.Message)"
                continue
            }
            foreach ($c in $children) {
                if (-not ($result | Where-Object { $_.Id -eq $c.Id })) {
                    Write-Verbose "Found child folder: $($c.DisplayName)"
                    $result.Add([PSCustomObject]@{ Id = $c.Id; DisplayName = $c.DisplayName })
                    $queue.Enqueue([PSCustomObject]@{ Id = $c.Id; DisplayName = $c.DisplayName })
                }
            }
        }
        return $result
    }

    # Retrieve messages
    try {
        if ($targetFolderIds.Count -gt 0) {
            if ($Recurse) {
                Write-Verbose "Searching folder '$FolderName' and all subfolders for messages older than $($CutoffDate.ToString('yyyy-MM-dd'))"
            $folderIds = Get-DescendantFolderIds -UserId $Mailbox -RootIds $targetFolderIds
            } else {
                Write-Verbose "Searching folder '$FolderName' for messages older than $($CutoffDate.ToString('yyyy-MM-dd'))"
                $folderIds = $targetFolderIds
            }
            # Get all descendant folder ids for each matched root
            foreach ($f in $folderIds) {
                $fid = $f.Id
                $fdisplay = $f.DisplayName
                try {
                    $parms = @{
                        UserId       = $Mailbox
                        MailFolderId = $fid
                        Filter       = $odataFilter
                        Property     = 'From,Id,ReceivedDateTime,Subject'
                    }
                    $msgs = Get-MgUserMailFolderMessage @parms -all -ErrorAction Stop
                    Write-Verbose "Retrieved $($msgs.Count) messages from folder '$fdisplay' older than $($CutoffDate.ToString('yyyy-MM-dd'))"
                } catch {
                    Write-Warning "Failed to retrieve messages from folder id $fid : $($_.Exception.Message)"
                    continue
                }

                foreach ($m in $msgs) {
                    $oldEmails.Add([PSCustomObject]@{
                            Subject      = $m.Subject
                            Sender       = ($m.From -and $m.From.EmailAddress) ? $m.From.EmailAddress.Name : ''
                            SenderEmail  = ($m.From -and $m.From.EmailAddress) ? $m.From.EmailAddress.Address : ''
                            ReceivedTime = $m.ReceivedDateTime
                            FolderPath   = $fdisplay
                            Id           = $m.Id
                        })
                }
            }
        } else {
            # No specific folder - retrieve all mailbox messages matching filter
            Write-Verbose "Searching entire mailbox '$Mailbox' for messages older than $($CutoffDate.ToString('yyyy-MM-dd'))"
            $msgs = Get-MgUserMessage -UserId $Mailbox -Filter $odataFilter -All -ErrorAction Stop
            Write-Verbose "Retrieved $($msgs.Count) messages from mailbox '$Mailbox' older than $($CutoffDate.ToString('yyyy-MM-dd'))"
            foreach ($m in $msgs) {
                $oldEmails.Add([PSCustomObject]@{
                        Subject      = $m.Subject
                        Sender       = ($m.From -and $m.From.EmailAddress) ? $m.From.EmailAddress.Name : ''
                        SenderEmail  = ($m.From -and $m.From.EmailAddress) ? $m.From.EmailAddress.Address : ''
                        ReceivedTime = $m.ReceivedDateTime
                        FolderPath   = $fid.DisplayName
                        Id           = $m.Id
                    })
            }
        }
    } catch {
        Write-Error "Failed to retrieve messages: $($_.Exception.Message)"
        return
    }

    # Deletion via Graph (permanent delete)
    if ($Delete -and $oldEmails.Count -gt 0) {
        Write-Verbose "Deleting $($oldEmails.Count) messages via Graph..."
        $i = 0
        $showProgress = $false
        if ($PSBoundParameters.ContainsKey('Verbose') -or $VerbosePreference -eq 'Continue') { $showProgress = $true }
        if ($showProgress) { $activity = "Deleting $($oldEmails.Count) messages from $Mailbox" }

        foreach ($e in $oldEmails) {
            $i++
            try {
                if ($e.Id) {
                    if ($showProgress) {
                        $percent = if ($oldEmails.Count -gt 0) { [int](($i / $oldEmails.Count) * 100) } else { 100 }
                        Write-Progress -Activity $activity -Status "Deleting $i of $($oldEmails.Count): $($e.Subject)" -PercentComplete $percent
                    } 
                    Remove-MgUserMessage -UserId $Mailbox -MessageId $e.Id -ErrorAction Stop
                }
            } catch {
                Write-Warning "Failed to delete message Id $($e.Id): $($_.Exception.Message)"
            }
        }

        if ($showProgress) { Write-Progress -Activity $activity -Completed }
    }

    # Output / Summary
    if ($SummaryOnly) {
        Write-Output '=== SUMMARY ==='
        Write-Output "Mailbox: $Mailbox"
        Write-Output "Total messages found older than $($CutoffDate.ToString('yyyy-MM-dd')): $($oldEmails.Count)"
        $totalSize = ($oldEmails | Measure-Object -Property Size -Sum).Sum
        $totalSizeMB = [math]::Round(($totalSize / 1MB), 2)
        Write-Output "Total size: $totalSizeMB MB"
        if ($oldEmails.Count -gt 0) {
            Write-Output 'Messages by folder:'
            $oldEmails | Group-Object FolderPath | Sort-Object Count -Descending | ForEach-Object {
                Write-Output "  $($_.Name): $($_.Count) messages"
            }
        }
    }

    if ($oldEmails.Count -gt 0 -and $OutputPath) {
        $oldEmails | Sort-Object ReceivedTime | Select-Object Subject, Sender, SenderEmail, ReceivedTime, Size, FolderPath, HasAttachments, Importance | Export-Csv -Path $OutputPath -NoTypeInformation
        Write-Verbose "Results exported to: $OutputPath"
    }

    Write-Verbose 'Function completed successfully (Graph)'
    if (-not $SummaryOnly) { return $oldEmails }
    if ($graphConnected) {
        Write-Verbose 'Disconnecting from Microsoft Graph...'
        Disconnect-MgGraph | Out-Null
    }
}
