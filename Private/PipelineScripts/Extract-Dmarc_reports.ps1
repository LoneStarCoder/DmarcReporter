    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$SourceFolder = '.\dmarc_report_emails',

        [Parameter()]
        [string]$DestinationFolder = '.\dmarc_xml_exports',

        [Parameter()]
        [switch]$Recurse,

        [Parameter()]
        [switch]$Overwrite
    )

function Export-DmarcMsgAttachment {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$SourceFolder = '.\dmarc_report_emails',

        [Parameter()]
        [switch]$Recurse,

        [Parameter()]
        [switch]$Overwrite,

        [Parameter()]
        [switch]$DeleteMsgAfterExtract
    )

    begin {
        function Get-SafeFileName {
            [CmdletBinding()]
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
            [CmdletBinding()]
            param(
                [Parameter(Mandatory = $true)]
                [string]$Directory,

                [Parameter(Mandatory = $true)]
                [string]$FileName,

                [Parameter()]
                [switch]$Overwrite
            )

            $safeFileName = Get-SafeFileName -FileName $FileName
            $candidatePath = Join-Path -Path $Directory -ChildPath $safeFileName

            if ($Overwrite -or -not (Test-Path -LiteralPath $candidatePath)) {
                return $candidatePath
            }

            return $null
        }
    }

    process {
        $resolvedSourceFolder = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($SourceFolder)

        if (-not (Test-Path -LiteralPath $resolvedSourceFolder -PathType Container)) {
            throw "Source folder not found: $resolvedSourceFolder"
        }

        $getChildItemParams = @{
            LiteralPath = $resolvedSourceFolder
            File        = $true
            Filter      = '*.msg'
        }

        if ($Recurse) {
            $getChildItemParams.Recurse = $true
        }

        $msgFiles = Get-ChildItem @getChildItemParams

        $outlook = $null
        $results = New-Object System.Collections.Generic.List[object]

        try {
            $outlook = New-Object -ComObject Outlook.Application

            foreach ($msgFile in $msgFiles) {
                $mailItem = $null

                try {
                    $mailItem = $outlook.CreateItemFromTemplate($msgFile.FullName)

                    if ($null -eq $mailItem) {
                        $results.Add([pscustomobject]@{
                            MsgFile        = $msgFile.FullName
                            Attachment     = $null
                            SavedTo        = $null
                            Extracted      = $false
                            Message        = 'Outlook returned a null item for this .msg file.'
                        }) | Out-Null

                        continue
                    }

                    if ($mailItem.Attachments.Count -eq 0) {
                        $results.Add([pscustomobject]@{
                            MsgFile        = $msgFile.FullName
                            Attachment     = $null
                            SavedTo        = $null
                            Extracted      = $false
                            Message        = 'No attachments found in .msg file.'
                        }) | Out-Null

                        continue
                    }

                    for ($i = 1; $i -le $mailItem.Attachments.Count; $i++) {
                        $attachment = $null

                        try {
                            $attachment = $mailItem.Attachments.Item($i)

                            if ([string]::IsNullOrWhiteSpace($attachment.FileName)) {
                                $attachmentFileName = "attachment_$i"
                            }
                            else {
                                $attachmentFileName = $attachment.FileName
                            }

                            $destinationPath = Get-UniqueFilePath `
                                -Directory $resolvedSourceFolder `
                                -FileName $attachmentFileName `
                                -Overwrite:$Overwrite

                            if ([string]::IsNullOrWhiteSpace($destinationPath)) {
                                $results.Add([pscustomobject]@{
                                    MsgFile        = $msgFile.FullName
                                    Attachment     = $attachmentFileName
                                    SavedTo        = $null
                                    Extracted      = $false
                                    Message        = 'Skipped because the destination file already exists and overwrite is disabled.'
                                }) | Out-Null

                                continue
                            }

                            $attachment.SaveAsFile($destinationPath)

                            $results.Add([pscustomobject]@{
                                MsgFile        = $msgFile.FullName
                                Attachment     = $attachmentFileName
                                SavedTo        = $destinationPath
                                Extracted      = $true
                                Message        = $null
                            }) | Out-Null
                        }
                        catch {
                            $results.Add([pscustomobject]@{
                                MsgFile        = $msgFile.FullName
                                Attachment     = $attachmentFileName
                                SavedTo        = $null
                                Extracted      = $false
                                Message        = $_.Exception.Message
                            }) | Out-Null
                        }
                        finally {
                            if ($null -ne $attachment) {
                                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($attachment) | Out-Null
                            }
                        }
                    }

                    if ($DeleteMsgAfterExtract) {
                        Remove-Item -LiteralPath $msgFile.FullName -Force
                    }
                }
                catch {
                    $results.Add([pscustomobject]@{
                        MsgFile        = $msgFile.FullName
                        Attachment     = $null
                        SavedTo        = $null
                        Extracted      = $false
                        Message        = $_.Exception.Message
                    }) | Out-Null
                }
                finally {
                    if ($null -ne $mailItem) {
                        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($mailItem) | Out-Null
                    }
                }
            }

            [pscustomobject]@{
                SourceFolder         = $resolvedSourceFolder
                MsgFilesFound        = $msgFiles.Count
                AttachmentsExtracted = @($results | Where-Object { $_.Extracted }).Count
                FailedExtractions    = @($results | Where-Object { -not $_.Extracted }).Count
                Results              = $results
            }
        }
        finally {
            if ($null -ne $outlook) {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null
            }

            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    }
}


function Export-DmarcXmlReport {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$SourceFolder = '.\dmarc_report_emails',

        [Parameter()]
        [string]$DestinationFolder = '.\dmarc_xml_exports',

        [Parameter()]
        [switch]$Recurse,

        [Parameter()]
        [switch]$Overwrite
    )

    begin {
        function Get-SafeFileName {
            [CmdletBinding()]
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
            [CmdletBinding()]
            param(
                [Parameter(Mandatory = $true)]
                [string]$Directory,

                [Parameter(Mandatory = $true)]
                [string]$FileName,

                [Parameter()]
                [switch]$Overwrite
            )

            $safeFileName = Get-SafeFileName -FileName $FileName
            $candidatePath = Join-Path -Path $Directory -ChildPath $safeFileName

            if ($Overwrite -or -not (Test-Path -LiteralPath $candidatePath)) {
                return $candidatePath
            }

            return $null
        }

        function Export-GZipXmlFile {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory = $true)]
                [System.IO.FileInfo]$File,

                [Parameter(Mandatory = $true)]
                [string]$DestinationDirectory,

                [Parameter()]
                [switch]$Overwrite
            )

            $outputFileName = $File.Name

            if ($outputFileName.EndsWith('.xml.gz', [System.StringComparison]::OrdinalIgnoreCase)) {
                $outputFileName = $outputFileName.Substring(0, $outputFileName.Length - 3)
            }
            elseif ($outputFileName.EndsWith('.gz', [System.StringComparison]::OrdinalIgnoreCase)) {
                $outputFileName = $outputFileName.Substring(0, $outputFileName.Length - 3)

                if (-not $outputFileName.EndsWith('.xml', [System.StringComparison]::OrdinalIgnoreCase)) {
                    $outputFileName = "$outputFileName.xml"
                }
            }

            $destinationPath = Get-UniqueFilePath `
                -Directory $DestinationDirectory `
                -FileName $outputFileName `
                -Overwrite:$Overwrite

            if ([string]::IsNullOrWhiteSpace($destinationPath)) {
                return [pscustomobject]@{
                    SourceFile      = $File.FullName
                    SourceType      = 'GZip'
                    XmlFile         = $null
                    Exported        = $false
                    Message         = 'Skipped because the destination file already exists and overwrite is disabled.'
                }
            }

            $inputStream = $null
            $gzipStream = $null
            $outputStream = $null

            try {
                $inputStream = [System.IO.File]::OpenRead($File.FullName)
                $gzipStream = [System.IO.Compression.GZipStream]::new(
                    $inputStream,
                    [System.IO.Compression.CompressionMode]::Decompress
                )
                $outputStream = [System.IO.File]::Create($destinationPath)

                $gzipStream.CopyTo($outputStream)

                return [pscustomobject]@{
                    SourceFile      = $File.FullName
                    SourceType      = 'GZip'
                    XmlFile         = $destinationPath
                    Exported        = $true
                    Message         = $null
                }
            }
            catch {
                if (Test-Path -LiteralPath $destinationPath) {
                    Remove-Item -LiteralPath $destinationPath -Force -ErrorAction SilentlyContinue
                }

                return [pscustomobject]@{
                    SourceFile      = $File.FullName
                    SourceType      = 'GZip'
                    XmlFile         = $null
                    Exported        = $false
                    Message         = $_.Exception.Message
                }
            }
            finally {
                if ($null -ne $outputStream) {
                    $outputStream.Dispose()
                }

                if ($null -ne $gzipStream) {
                    $gzipStream.Dispose()
                }

                if ($null -ne $inputStream) {
                    $inputStream.Dispose()
                }
            }
        }

        function Export-ZipXmlFile {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory = $true)]
                [System.IO.FileInfo]$File,

                [Parameter(Mandatory = $true)]
                [string]$DestinationDirectory,

                [Parameter()]
                [switch]$Overwrite
            )

            $results = New-Object System.Collections.Generic.List[object]
            $zipArchive = $null

            try {
                $zipArchive = [System.IO.Compression.ZipFile]::OpenRead($File.FullName)

                foreach ($entry in $zipArchive.Entries) {
                    if ([string]::IsNullOrWhiteSpace($entry.Name)) {
                        continue
                    }

                    if (-not $entry.Name.EndsWith('.xml', [System.StringComparison]::OrdinalIgnoreCase)) {
                        continue
                    }

                    $destinationFileName = Get-SafeFileName -FileName $entry.Name
                    $destinationPath = Get-UniqueFilePath `
                        -Directory $DestinationDirectory `
                        -FileName $destinationFileName `
                        -Overwrite:$Overwrite

                    if ([string]::IsNullOrWhiteSpace($destinationPath)) {
                        $results.Add([pscustomobject]@{
                            SourceFile      = $File.FullName
                            SourceType      = 'Zip'
                            ZipEntry        = $entry.FullName
                            XmlFile         = $null
                            Exported        = $false
                            Message         = 'Skipped because the destination file already exists and overwrite is disabled.'
                        }) | Out-Null

                        continue
                    }

                    try {
                        [System.IO.Compression.ZipFileExtensions]::ExtractToFile(
                            $entry,
                            $destinationPath,
                            [bool]$Overwrite
                        )

                        $results.Add([pscustomobject]@{
                            SourceFile      = $File.FullName
                            SourceType      = 'Zip'
                            ZipEntry        = $entry.FullName
                            XmlFile         = $destinationPath
                            Exported        = $true
                            Message         = $null
                        }) | Out-Null
                    }
                    catch {
                        $results.Add([pscustomobject]@{
                            SourceFile      = $File.FullName
                            SourceType      = 'Zip'
                            ZipEntry        = $entry.FullName
                            XmlFile         = $null
                            Exported        = $false
                            Message         = $_.Exception.Message
                        }) | Out-Null
                    }
                }

                if ($results.Count -eq 0) {
                    $results.Add([pscustomobject]@{
                        SourceFile      = $File.FullName
                        SourceType      = 'Zip'
                        ZipEntry        = $null
                        XmlFile         = $null
                        Exported        = $false
                        Message         = 'No .xml files found inside zip archive.'
                    }) | Out-Null
                }

                return $results
            }
            catch {
                return [pscustomobject]@{
                    SourceFile      = $File.FullName
                    SourceType      = 'Zip'
                    ZipEntry        = $null
                    XmlFile         = $null
                    Exported        = $false
                    Message         = $_.Exception.Message
                }
            }
            finally {
                if ($null -ne $zipArchive) {
                    $zipArchive.Dispose()
                }
            }
        }
    }

    process {
        Add-Type -AssemblyName System.IO.Compression.FileSystem

        $resolvedSourceFolder = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($SourceFolder)
        $resolvedDestinationFolder = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($DestinationFolder)

        if (-not (Test-Path -LiteralPath $resolvedSourceFolder -PathType Container)) {
            throw "Source folder not found: $resolvedSourceFolder"
        }

        if (-not (Test-Path -LiteralPath $resolvedDestinationFolder -PathType Container)) {
            New-Item -Path $resolvedDestinationFolder -ItemType Directory -Force | Out-Null
        }

        $getChildItemParams = @{
            LiteralPath = $resolvedSourceFolder
            File        = $true
        }

        if ($Recurse) {
            $getChildItemParams.Recurse = $true
        }

        $files = Get-ChildItem @getChildItemParams | Where-Object {
            $_.Name.EndsWith('.zip', [System.StringComparison]::OrdinalIgnoreCase) -or
            $_.Name.EndsWith('.gz', [System.StringComparison]::OrdinalIgnoreCase)
        }

        $results = New-Object System.Collections.Generic.List[object]

        foreach ($file in $files) {
            if ($file.Name.EndsWith('.zip', [System.StringComparison]::OrdinalIgnoreCase)) {
                $zipResults = Export-ZipXmlFile `
                    -File $file `
                    -DestinationDirectory $resolvedDestinationFolder `
                    -Overwrite:$Overwrite

                foreach ($result in @($zipResults)) {
                    $results.Add($result) | Out-Null
                }

                continue
            }

            if ($file.Name.EndsWith('.gz', [System.StringComparison]::OrdinalIgnoreCase)) {
                $gzipResult = Export-GZipXmlFile `
                    -File $file `
                    -DestinationDirectory $resolvedDestinationFolder `
                    -Overwrite:$Overwrite

                $results.Add($gzipResult) | Out-Null

                continue
            }
        }

        [pscustomobject]@{
            SourceFolder       = $resolvedSourceFolder
            DestinationFolder  = $resolvedDestinationFolder
            ArchiveFilesFound  = $files.Count
            XmlFilesExported   = @($results | Where-Object { $_.Exported }).Count
            FailedExports      = @($results | Where-Object { -not $_.Exported }).Count
            Results            = $results
        }
    }
}

$msgResult = Export-DmarcMsgAttachment `
    -SourceFolder $SourceFolder `
    -Recurse:$Recurse `
    -Overwrite:$Overwrite

$xmlResult = Export-DmarcXmlReport `
    -SourceFolder $SourceFolder `
    -DestinationFolder $DestinationFolder `
    -Recurse:$Recurse `
    -Overwrite:$Overwrite

[pscustomobject]@{
    SourceFolder      = $SourceFolder
    DestinationFolder = $DestinationFolder
    MsgExtraction     = $msgResult
    XmlExtraction     = $xmlResult
}
