<#
.SYNOPSIS
    Phase 3 - upload converted PDFs back to the SharePoint folder the source
    .pub came from.

.DESCRIPTION
    Implements brief section 6.

    Safety rules, all enforced here rather than left to the menu:
      * No upload runs without an explicit typed confirmation showing the count.
      * The default is to place the PDF ALONGSIDE the original .pub. The source
        file is never touched unless the operator opts in separately.
      * A name collision defaults to versioning (SharePoint renames to
        'Name 1.pdf'); overwriting is opt-in.
      * Deleting the source .pub is off by default and needs its own second
        confirmation, typed in full.
#>

Set-StrictMode -Version 2.0

foreach ($dependency in @('Logging', 'Config', 'Graph', 'Discovery')) {
    if (-not (Get-Module -Name $dependency)) {
        Import-Module (Join-Path $PSScriptRoot ("{0}.psm1" -f $dependency)) -Force -DisableNameChecking
    }
}

# Graph requires upload session chunks to be a multiple of 320 KiB.
$script:ChunkSize        = 3276800
$script:SimpleUploadMax  = 4MB

function Get-PubConflictBehavior {
    <#
    .SYNOPSIS
        Maps the operator's setting to Graph's conflictBehavior value.
    #>
    [CmdletBinding()]
    param(
        [ValidateSet('Skip', 'Overwrite', 'Version')]
        [string] $Action = 'Version'
    )

    switch ($Action) {
        'Overwrite' { return 'replace' }
        'Skip'      { return 'fail' }
        default     { return 'rename' }
    }
}

function Send-PubFileToSharePoint {
    <#
    .SYNOPSIS
        Uploads one local file into a drive folder.

    .PARAMETER DriveId
        Target document library drive id.

    .PARAMETER ParentItemId
        The folder item to upload into - the source .pub's parent, so the PDF
        lands beside it.

    .PARAMETER ConflictBehavior
        rename | replace | fail.

    .OUTPUTS
        The created driveItem.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $LocalPath,
        [Parameter(Mandatory)] [string] $DriveId,
        [Parameter(Mandatory)] [string] $ParentItemId,
        [Parameter(Mandatory)] [string] $FileName,

        [ValidateSet('rename', 'replace', 'fail')]
        [string] $ConflictBehavior = 'rename'
    )

    $file       = Get-Item -LiteralPath $LocalPath -ErrorAction Stop
    $encodedName = [uri]::EscapeDataString($FileName)

    if ($file.Length -lt $script:SimpleUploadMax) {
        $uri   = 'drives/{0}/items/{1}:/{2}:/content?@microsoft.graph.conflictBehavior={3}' -f $DriveId, $ParentItemId, $encodedName, $ConflictBehavior
        $bytes = [System.IO.File]::ReadAllBytes($file.FullName)

        return Invoke-PubGraph -Uri $uri -Method PUT -Body $bytes -ContentType 'application/pdf'
    }

    # ---- large file: resumable upload session ----
    $sessionUri  = 'drives/{0}/items/{1}:/{2}:/createUploadSession' -f $DriveId, $ParentItemId, $encodedName
    $sessionBody = @{
        item = @{
            '@microsoft.graph.conflictBehavior' = $ConflictBehavior
            name                                = $FileName
        }
    }

    $session = Invoke-PubGraph -Uri $sessionUri -Method POST -Body $sessionBody -ContentType 'application/json'
    if (-not $session -or -not $session.uploadUrl) { throw 'Graph did not return an upload URL.' }

    $stream = [System.IO.File]::OpenRead($file.FullName)
    try {
        $buffer    = New-Object byte[] $script:ChunkSize
        $position  = 0L
        $total     = $file.Length
        $response  = $null

        while ($position -lt $total) {
            $read = $stream.Read($buffer, 0, $buffer.Length)
            if ($read -le 0) { break }

            $chunk = $buffer
            if ($read -ne $buffer.Length) {
                $chunk = New-Object byte[] $read
                [Array]::Copy($buffer, 0, $chunk, 0, $read)
            }

            $rangeHeader = 'bytes {0}-{1}/{2}' -f $position, ($position + $read - 1), $total

            $attempt = 0
            while ($true) {
                $attempt++
                try {
                    # The upload URL is pre-authenticated - no Authorization header.
                    $response = Invoke-WebRequest -Uri $session.uploadUrl `
                                                  -Method Put `
                                                  -Headers @{ 'Content-Range' = $rangeHeader } `
                                                  -Body $chunk `
                                                  -ContentType 'application/octet-stream' `
                                                  -UseBasicParsing `
                                                  -ErrorAction Stop
                    break
                } catch {
                    $status = Get-PubHttpStatusCode -ErrorRecord $_
                    if ($attempt -ge 4 -or @(429, 500, 502, 503, 504) -notcontains $status) { throw }

                    $wait = [math]::Pow(2, $attempt)
                    Write-PubLog -Level Warn -Message ('Chunk upload got HTTP {0} - retrying in {1}s.' -f $status, [int] $wait)
                    Start-Sleep -Seconds ([int] $wait)
                }
            }

            $position += $read
        }

        if ($response -and $response.Content) {
            try { return ($response.Content | ConvertFrom-Json) } catch { return $null }
        }
        return $null
    } finally {
        $stream.Dispose()
    }
}

function Remove-PubSourceFile {
    <#
    .SYNOPSIS
        Deletes the original .pub from SharePoint (recoverable from the site
        recycle bin) after its PDF uploaded successfully.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $DriveId,
        [Parameter(Mandatory)] [string] $ItemId
    )

    Invoke-PubGraph -Uri ('drives/{0}/items/{1}' -f $DriveId, $ItemId) -Method DELETE | Out-Null
    return $true
}

function Invoke-PubUpload {
    <#
    .SYNOPSIS
        Menu option 9 - uploads converted PDFs to their original location.

    .DESCRIPTION
        Only rows with Status = Converted and a PDF on disk are eligible.
        Rows are updated to Uploaded or Failed, and the result is written to a
        NEW timestamped CSV so the original inventory stays as the audit
        baseline.

    .PARAMETER Confirmed
        Set by the menu once the operator has confirmed at the prompt. Without
        it this function asks for confirmation itself - there is no path that
        writes to SharePoint unconfirmed.

    .PARAMETER RemoveSource
        Also delete the original .pub after a successful upload. Off by
        default; requires its own typed confirmation.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        $Rows,

        $Config,
        [switch] $Confirmed,
        [switch] $RemoveSource,

        [ValidateSet('Skip', 'Overwrite', 'Version')]
        [string] $ConflictAction
    )

    if (-not $Config) { $Config = Get-PubConfig }
    if ([string]::IsNullOrWhiteSpace($ConflictAction)) {
        $ConflictAction = [string] $Config['UploadConflictAction']
        if ([string]::IsNullOrWhiteSpace($ConflictAction)) { $ConflictAction = 'Version' }
    }

    $rows = @($Rows)

    $eligible = New-Object System.Collections.Generic.List[object]
    foreach ($row in $rows) {
        if ([string] $row.Status -ne 'Converted') { continue }

        $pdfPath = [string] $row.PdfPath
        if ([string]::IsNullOrWhiteSpace($pdfPath) -or -not (Test-Path -LiteralPath $pdfPath)) {
            Set-PubInventoryStatus -Row $row -Status Failed -Notes 'Converted PDF is missing from the working folder - re-run the conversion.' | Out-Null
            Write-PubFileResult -Outcome Failed -FileName ([string] $row.FileName) -Detail 'PDF missing locally'
            continue
        }

        $eligible.Add($row)
    }

    if ($eligible.Count -eq 0) {
        Write-PubLog -Level Warn -Message 'No converted PDFs are ready to upload. Run options 7 and 8 first.'
        return $rows
    }

    # ---- confirmation gate (brief section 6) ----
    Write-Host ''
    Write-Host ('{0} file(s) ready to upload.' -f $eligible.Count) -ForegroundColor Cyan
    Write-Host ('  Destination      : the original SharePoint folder of each .pub') -ForegroundColor Gray
    Write-Host ('  Name collisions  : {0}' -f $ConflictAction) -ForegroundColor Gray
    if ($RemoveSource) {
        Write-Host  '  Source .pub      : WILL BE DELETED after a successful upload' -ForegroundColor Red
    } else {
        Write-Host  '  Source .pub      : left in place' -ForegroundColor Gray
    }
    Write-Host ''

    if (-not $Confirmed) {
        $answer = Read-Host ('Upload {0} PDF(s) to SharePoint? (Y/N)' -f $eligible.Count)
        if ($answer -notmatch '^[Yy]') {
            Write-PubLog -Level Warn -Message 'Upload cancelled by the operator.'
            return $rows
        }
    }

    if ($RemoveSource) {
        $typed = Read-Host 'Type DELETE to confirm the original .pub files should be removed after upload'
        if ($typed -cne 'DELETE') {
            Write-PubLog -Level Warn -Message 'Source deletion not confirmed - uploading only, originals will be left in place.'
            $RemoveSource = $false
        }
    }

    if (-not (Connect-PubGraphApp -Config $Config)) { return $rows }

    Write-PubLog -Level Info -Message ('=== Phase 3: Upload ({0} file(s)) ===' -f $eligible.Count)

    $conflictBehavior = Get-PubConflictBehavior -Action $ConflictAction
    $uploaded = 0
    $failed   = 0
    $skipped  = 0
    $deleted  = 0
    $index    = 0

    foreach ($row in $eligible) {
        $index++
        Write-Progress -Activity 'Uploading converted PDFs' `
                       -Status ('{0} of {1}' -f $index, $eligible.Count) `
                       -CurrentOperation ([string] $row.FileName) `
                       -PercentComplete ([int] (($index / $eligible.Count) * 100))

        $pdfName = [System.IO.Path]::ChangeExtension([string] $row.FileName, '.pdf')

        $parentItemId = [string] $row.ParentItemId
        if ([string]::IsNullOrWhiteSpace($parentItemId)) {
            $failed++
            Set-PubInventoryStatus -Row $row -Status Failed -Notes 'Row has no ParentItemId - re-run discovery for this file.' | Out-Null
            Write-PubFileResult -Outcome Failed -FileName $pdfName -Detail 'missing ParentItemId'
            continue
        }

        try {
            $item = Send-PubFileToSharePoint -LocalPath ([string] $row.PdfPath) `
                                             -DriveId ([string] $row.DriveId) `
                                             -ParentItemId $parentItemId `
                                             -FileName $pdfName `
                                             -ConflictBehavior $conflictBehavior

            $uploadedName = $pdfName
            if ($item -and $item.PSObject.Properties['name'] -and $item.name) { $uploadedName = $item.name }

            $note = ''
            if ($uploadedName -ne $pdfName) { $note = 'Name collision - uploaded as "{0}".' -f $uploadedName }

            Set-PubInventoryStatus -Row $row -Status Uploaded -Notes $note | Out-Null
            $uploaded++
            Write-PubFileResult -Outcome Uploaded -FileName $uploadedName -Detail ('{0} / {1}{2}' -f $row.SiteUrl, $row.LibraryName, $row.FolderPath)

            if ($RemoveSource) {
                try {
                    Remove-PubSourceFile -DriveId ([string] $row.DriveId) -ItemId ([string] $row.ItemId) | Out-Null
                    $deleted++
                    Write-PubFileResult -Outcome Deleted -FileName ([string] $row.FileName) -Detail 'source removed (recoverable from the site recycle bin)'
                } catch {
                    $message = Get-PubGraphErrorMessage -ErrorRecord $_
                    Set-PubInventoryStatus -Row $row -Status Uploaded -Notes ('PDF uploaded, but the source .pub could not be deleted: {0}' -f $message) | Out-Null
                    Write-PubLog -Level Warn -Message ('Could not delete {0}: {1}' -f $row.FileName, $message)
                }
            }
        } catch {
            $message = Get-PubGraphErrorMessage -ErrorRecord $_
            $status  = Get-PubHttpStatusCode -ErrorRecord $_

            if ($status -eq 409 -and $ConflictAction -eq 'Skip') {
                $skipped++
                Set-PubInventoryStatus -Row $row -Status Skipped -Notes 'A PDF with this name already exists and the collision setting is Skip.' | Out-Null
                Write-PubFileResult -Outcome Skipped -FileName $pdfName -Detail 'already present in SharePoint'
            } else {
                $failed++
                Set-PubInventoryStatus -Row $row -Status Failed -Notes ('Upload failed: {0}' -f $message) | Out-Null
                Write-PubFileResult -Outcome Failed -FileName $pdfName -Detail $message
            }
        }
    }

    Write-Progress -Activity 'Uploading converted PDFs' -Completed

    # New timestamped CSV - the original inventory is left untouched as the
    # audit baseline (brief section 6).
    $inventoryFolder = Get-PubWorkingFolder -SubFolder 'inventory' -Config $Config
    $resultCsv       = Join-Path $inventoryFolder ('PublisherFileInventory_{0}_uploaded.csv' -f (Get-Date -Format 'yyyy-MM-dd_HHmm'))
    Export-PubInventory -Rows $rows -Path $resultCsv -Config $Config | Out-Null

    $extraLines = @(('Updated CSV : {0}' -f $resultCsv))
    if ($deleted -gt 0) { $extraLines += ('Sources deleted : {0}' -f $deleted) }

    Write-PubPhaseSummary -Phase 'Upload' -Attempted $eligible.Count -Succeeded $uploaded -Failed $failed -Skipped $skipped -ExtraLines $extraLines

    return $rows
}

Export-ModuleMember -Function @(
    'Get-PubConflictBehavior'
    'Send-PubFileToSharePoint'
    'Remove-PubSourceFile'
    'Invoke-PubUpload'
)
