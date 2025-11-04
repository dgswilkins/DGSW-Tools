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
