    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateRange(1, 365)]
        [int]$Days,

        [Parameter(Mandatory = $true)]
        [string]$EmailFolderPath,

        [Parameter()]
        [string]$DestinationFolder = '.\dmarc_report_emails',

        [Parameter()]
        [ValidateSet('UnreadOnly', 'All')]
        [string]$MessageFilter = 'UnreadOnly',

        [Parameter()]
        [bool]$MarkAsRead=$false
    )

    begin {
        function Get-SafeFileName {
            param(
                [Parameter(Mandatory = $true)]
                [string]$FileName
            )

            $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
            $safeName = $FileName

            foreach ($char in $invalidChars) {
                $safeName = $safeName.Replace($char, '_')
            }

            return $safeName
        }

        function Get-UniqueFilePath {
            param(
                [Parameter(Mandatory = $true)]
                [string]$Directory,

                [Parameter(Mandatory = $true)]
                [string]$FileName
            )

            $safeFileName = Get-SafeFileName -FileName $FileName
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($safeFileName)
            $extension = [System.IO.Path]::GetExtension($safeFileName)

            $candidatePath = Join-Path -Path $Directory -ChildPath $safeFileName
            $counter = 1

            while (Test-Path -LiteralPath $candidatePath) {
                $candidateName = '{0}_{1}{2}' -f $baseName, $counter, $extension
                $candidatePath = Join-Path -Path $Directory -ChildPath $candidateName
                $counter++
            }

            return $candidatePath
        }

        function Get-OutlookFolderByPath {
            param(
                [Parameter(Mandatory = $true)]
                [object]$NameSpace,

                [Parameter(Mandatory = $true)]
                [string]$Path
            )

            $parts = $Path -split '\\'

            if ($parts[0] -ieq 'Inbox') {
                $folder = $NameSpace.GetDefaultFolder(6) # 6 = olFolderInbox
                $startIndex = 1
            }
            else {
                throw "This function currently expects the folder path to start with 'Inbox'. Path received: $Path"
            }

            for ($i = $startIndex; $i -lt $parts.Count; $i++) {
                $childName = $parts[$i]
                $nextFolder = $null

                foreach ($child in $folder.Folders) {
                    if ($child.Name -ieq $childName) {
                        $nextFolder = $child
                        break
                    }
                }

                if ($null -eq $nextFolder) {
                    throw "Outlook folder not found: $($parts[0..$i] -join '\')"
                }

                $folder = $nextFolder
            }

            return $folder
        }
    }

    process {
        $resolvedDestination = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($DestinationFolder)

        if (-not (Test-Path -LiteralPath $resolvedDestination)) {
            New-Item -Path $resolvedDestination -ItemType Directory -Force | Out-Null
        }

        $outlook = $null
        $namespace = $null

        try {
            $outlook = New-Object -ComObject Outlook.Application
            $namespace = $outlook.GetNamespace('MAPI')

            $targetFolder = Get-OutlookFolderByPath -NameSpace $namespace -Path $EmailFolderPath

            $since = (Get-Date).AddDays(-$Days)

            # Outlook Restrict date parsing is locale-sensitive.
            # This format is commonly accepted by the Outlook COM object model on Windows.
            $sinceFilterDate = $since.ToString('g', [System.Globalization.CultureInfo]::CurrentCulture)

            $filterParts = @(
                "[ReceivedTime] >= '$sinceFilterDate'"
            )

            if ($MessageFilter -eq 'UnreadOnly') {
                $filterParts = @('[UnRead] = True') + $filterParts
            }

            $filter = $filterParts -join ' AND '

            $items = $targetFolder.Items
            $items.Sort('[ReceivedTime]', $true)

            $filteredItems = $items.Restrict($filter)

            $savedFiles = New-Object System.Collections.Generic.List[object]
            $mailEntryIdsToMarkRead = New-Object System.Collections.Generic.List[string]
            $processedMailCount = 0
            $savedAttachmentCount = 0

            foreach ($item in @($filteredItems)) {
                # 43 = olMail
                if ($null -eq $item -or $item.Class -ne 43) {
                    continue
                }

                if ($MessageFilter -eq 'UnreadOnly' -and -not $item.UnRead) {
                    continue
                }

                $processedMailCount++

                if ($item.Attachments.Count -gt 0) {
                    for ($i = 1; $i -le $item.Attachments.Count; $i++) {
                        $attachment = $item.Attachments.Item($i)

                        $targetPath = Get-UniqueFilePath `
                            -Directory $resolvedDestination `
                            -FileName $attachment.FileName

                        $attachment.SaveAsFile($targetPath)
                        $savedAttachmentCount++

                        $savedFiles.Add([pscustomobject]@{
                            ReceivedTime = $item.ReceivedTime
                            Subject      = $item.Subject
                            Sender       = $item.SenderEmailAddress
                            Attachment   = $attachment.FileName
                            SavedTo      = $targetPath
                        }) | Out-Null

                        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($attachment) | Out-Null
                    }
                }

                if ($MarkAsRead -and $item.UnRead -and -not [string]::IsNullOrWhiteSpace($item.EntryID)) {
                    $mailEntryIdsToMarkRead.Add($item.EntryID) | Out-Null
                }
            }

            foreach ($entryId in $mailEntryIdsToMarkRead) {
                $mailItemToMarkRead = $null

                try {
                    $mailItemToMarkRead = $namespace.GetItemFromID($entryId)

                    if ($null -ne $mailItemToMarkRead -and $mailItemToMarkRead.Class -eq 43 -and $mailItemToMarkRead.UnRead) {
                        $mailItemToMarkRead.UnRead = $false
                        $mailItemToMarkRead.Save()
                    }
                }
                finally {
                    if ($null -ne $mailItemToMarkRead) {
                        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($mailItemToMarkRead) | Out-Null
                    }
                }
            }

            [pscustomobject]@{
                Folder             = $EmailFolderPath
                Since              = $since
                DestinationFolder  = $resolvedDestination
                MessageFilter      = $MessageFilter
                ProcessedMailCount = $processedMailCount
                SavedAttachments   = $savedAttachmentCount
                SavedFiles         = $savedFiles
            }
        }
        finally {
            if ($null -ne $filteredItems) {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($filteredItems) | Out-Null
            }

            if ($null -ne $items) {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($items) | Out-Null
            }

            if ($null -ne $targetFolder) {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($targetFolder) | Out-Null
            }

            if ($null -ne $namespace) {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($namespace) | Out-Null
            }

            if ($null -ne $outlook) {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null
            }

            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    }
