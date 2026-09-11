<#
.SYNOPSIS
    Phase 2 - download the selected .pub files and drive the PDF conversion.

.DESCRIPTION
    Implements brief section 5. This module does the Graph half (selection and
    download) in whichever PowerShell the menu is running under, then hands the
    COM half to Convert.Publisher.ps1 in a Windows PowerShell 5.1 child
    process.

    Why out-of-process: early-bound Interop
    (Add-Type -AssemblyName Microsoft.Office.Interop.Publisher, resolved from
    the GAC) does not load reliably under PowerShell 7's .NET Core runtime.
    Running the conversion under 5.1 keeps Tom's working COM code working, and
    the process boundary has a second benefit - a .pub file that hangs
    Publisher on a password prompt kills one child process, not the whole run.

    Hard prerequisite: a Windows host with Microsoft Publisher installed. This
    is checked before any conversion starts, with a clear error if Publisher is
    absent.
#>

Set-StrictMode -Version 2.0

foreach ($dependency in @('Logging', 'Config', 'Graph', 'Discovery')) {
    if (-not (Get-Module -Name $dependency)) {
        Import-Module (Join-Path $PSScriptRoot ("{0}.psm1" -f $dependency)) -Force -DisableNameChecking
    }
}

$script:ConverterScript      = Join-Path $PSScriptRoot 'Convert.Publisher.ps1'
$script:PublisherProbeResult = $null

function Get-PubWindowsPowerShellPath {
    <#
    .SYNOPSIS
        Locates powershell.exe (Windows PowerShell 5.1). $null if unavailable.
    #>
    [CmdletBinding()]
    param()

    if (-not (Test-PubIsWindows)) { return $null }

    $candidates = @(
        (Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe')
        (Join-Path $env:SystemRoot 'Sysnative\WindowsPowerShell\v1.0\powershell.exe')
    )

    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path -LiteralPath $candidate)) { return $candidate }
    }

    $command = Get-Command -Name 'powershell.exe' -ErrorAction SilentlyContinue
    if ($command) { return $command.Source }

    return $null
}

function Test-PubPublisherAvailable {
    <#
    .SYNOPSIS
        Confirms Microsoft Publisher can actually be automated on this host.

    .DESCRIPTION
        Runs Convert.Publisher.ps1 -TestOnly under Windows PowerShell 5.1. The
        result is cached for the session; pass -Refresh to re-probe.
    #>
    [CmdletBinding()]
    param(
        [switch] $Refresh,
        [switch] $Quiet
    )

    if ($null -ne $script:PublisherProbeResult -and -not $Refresh) {
        return $script:PublisherProbeResult
    }

    if (-not (Test-PubIsWindows)) {
        if (-not $Quiet) {
            Write-PubLog -Level Error -Message 'Conversion requires a Windows host with Microsoft Publisher installed - this host is not Windows.'
        }
        $script:PublisherProbeResult = $false
        return $false
    }

    $powerShell = Get-PubWindowsPowerShellPath
    if (-not $powerShell) {
        if (-not $Quiet) { Write-PubLog -Level Error -Message 'Windows PowerShell 5.1 (powershell.exe) was not found on this host.' }
        $script:PublisherProbeResult = $false
        return $false
    }

    try {
        $output = & $powerShell -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $script:ConverterScript -TestOnly 2>&1
        $ok     = ($LASTEXITCODE -eq 0)

        if ($ok) {
            if (-not $Quiet) { Write-PubLog -Level Success -Message 'Microsoft Publisher is installed and automatable on this host.' }
        } else {
            if (-not $Quiet) {
                Write-PubLog -Level Error -Message ('Microsoft Publisher is not available: {0}' -f ($output -join ' '))
                Write-PubLog -Level Info  -Message 'The download and upload phases still work here - run the conversion phase on a host with Publisher installed.'
            }
        }

        $script:PublisherProbeResult = $ok
        return $ok
    } catch {
        if (-not $Quiet) { Write-PubLog -Level Error -Message ('Publisher check failed: {0}' -f $_.Exception.Message) }
        $script:PublisherProbeResult = $false
        return $false
    }
}

function ConvertTo-PubSafeSegment {
    <#
    .SYNOPSIS
        Makes one path segment safe for the local file system.
    #>
    [CmdletBinding()]
    param(
        [AllowEmptyString()]
        [string] $Segment
    )

    if ([string]::IsNullOrWhiteSpace($Segment)) { return '_' }

    # The Windows invalid-character set is used explicitly rather than
    # [System.IO.Path]::GetInvalidFileNameChars(), which on Linux/macOS only
    # covers '/' - a path built on PowerShell 7 must still be valid when the
    # conversion step reads it on Windows.
    $invalid = @('<', '>', ':', '"', '/', '', '|', '?', '*') + [System.IO.Path]::GetInvalidFileNameChars()
    $builder = New-Object System.Text.StringBuilder

    foreach ($character in $Segment.ToCharArray()) {
        if ($invalid -contains $character -or [int] $character -lt 32) { [void] $builder.Append('_') }
        else { [void] $builder.Append($character) }
    }

    $safe = $builder.ToString().Trim().TrimEnd('.')
    if ([string]::IsNullOrWhiteSpace($safe)) { $safe = '_' }
    if ($safe.Length -gt 60) { $safe = $safe.Substring(0, 60) }

    return $safe
}

function Get-PubLocalPath {
    <#
    .SYNOPSIS
        Builds the local path for one inventory row.

    .DESCRIPTION
        Mirrors SiteUrl / LibraryName / FolderPath under the working folder so
        files from different sites cannot collide, and so the upload phase can
        find its way back (brief section 5.2).

    .PARAMETER Kind
        'Original' for the downloaded .pub, 'Pdf' for the converted PDF.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Row,

        [ValidateSet('Original', 'Pdf')]
        [string] $Kind = 'Original',

        $Config
    )

    if (-not $Config) { $Config = Get-PubConfig }

    $subFolder = 'originals'
    if ($Kind -eq 'Pdf') { $subFolder = 'converted' }

    $root = Get-PubWorkingFolder -SubFolder $subFolder -Config $Config

    $siteSegment = 'site'
    if ($Row.SiteUrl) {
        try {
            $siteUri     = [uri] $Row.SiteUrl
            $siteSegment = '{0}{1}' -f $siteUri.Host, ($siteUri.AbsolutePath -replace '/', '_')
        } catch {
            $siteSegment = [string] $Row.SiteUrl
        }
    }

    $path = Join-Path $root (ConvertTo-PubSafeSegment -Segment $siteSegment)
    $path = Join-Path $path (ConvertTo-PubSafeSegment -Segment ([string] $Row.LibraryName))

    $folderPath = [string] $Row.FolderPath
    if ($folderPath -and $folderPath -ne '/') {
        foreach ($segment in ($folderPath -split '/')) {
            if ([string]::IsNullOrWhiteSpace($segment)) { continue }
            $path = Join-Path $path (ConvertTo-PubSafeSegment -Segment $segment)
        }
    }

    $fileName = [string] $Row.FileName
    if ($Kind -eq 'Pdf') { $fileName = [System.IO.Path]::ChangeExtension($fileName, '.pdf') }

    return (Join-Path $path $fileName)
}

function Select-PubInventoryRow {
    <#
    .SYNOPSIS
        Filters inventory rows for the 'select files to process' sub-menu.

    .PARAMETER SiteFilter
        Wildcard matched against SiteUrl, e.g. *marketing*.

    .PARAMETER FolderFilter
        Wildcard matched against LibraryName + FolderPath.

    .PARAMETER NameFilter
        Wildcard matched against FileName.

    .PARAMETER Status
        Keep only rows in these statuses.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        $Rows,

        [string]   $SiteFilter,
        [string]   $FolderFilter,
        [string]   $NameFilter,
        [string[]] $Status
    )

    $selected = @($Rows)

    if (-not [string]::IsNullOrWhiteSpace($SiteFilter)) {
        $selected = @($selected | Where-Object { [string] $_.SiteUrl -like $SiteFilter })
    }
    if (-not [string]::IsNullOrWhiteSpace($FolderFilter)) {
        $selected = @($selected | Where-Object { ('{0}{1}' -f $_.LibraryName, $_.FolderPath) -like $FolderFilter })
    }
    if (-not [string]::IsNullOrWhiteSpace($NameFilter)) {
        $selected = @($selected | Where-Object { [string] $_.FileName -like $NameFilter })
    }
    if ($Status -and $Status.Count -gt 0) {
        $selected = @($selected | Where-Object { $Status -contains ([string] $_.Status) })
    }

    return $selected
}

function Invoke-PubDownload {
    <#
    .SYNOPSIS
        Menu option 7 - downloads the selected files to the working folder.

    .DESCRIPTION
        Already-downloaded files are left alone unless -Force is supplied, so
        re-running after an interruption does not re-fetch what is already on
        disk (brief section 2, resumability).

        Rows are updated in place: Status becomes Downloaded or Failed, with
        the reason in Notes.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        $Rows,

        $Config,
        [switch] $Force
    )

    if (-not $Config) { $Config = Get-PubConfig }

    $rows = @($Rows)
    if ($rows.Count -eq 0) {
        Write-PubLog -Level Warn -Message 'No rows selected - nothing to download.'
        return $rows
    }

    if (-not (Connect-PubGraphApp -Config $Config)) { return $rows }

    Write-PubLog -Level Info -Message ('=== Phase 2a: Download ({0} file(s)) ===' -f $rows.Count)

    $succeeded = 0
    $failed    = 0
    $skipped   = 0
    $index     = 0

    foreach ($row in $rows) {
        $index++
        Write-Progress -Activity 'Downloading Publisher files' `
                       -Status ('{0} of {1}' -f $index, $rows.Count) `
                       -CurrentOperation ([string] $row.FileName) `
                       -PercentComplete ([int] (($index / $rows.Count) * 100))

        $localPath = Get-PubLocalPath -Row $row -Kind Original -Config $Config

        if ((Test-Path -LiteralPath $localPath) -and -not $Force) {
            $row.LocalPath = $localPath
            if ([string] $row.Status -eq 'Pending') { Set-PubInventoryStatus -Row $row -Status Downloaded -Notes 'Already present locally.' | Out-Null }
            $skipped++
            Write-PubFileResult -Outcome Skipped -FileName ([string] $row.FileName) -Detail 'already downloaded'
            continue
        }

        if ([string]::IsNullOrWhiteSpace([string] $row.DriveId) -or [string]::IsNullOrWhiteSpace([string] $row.ItemId)) {
            $failed++
            Set-PubInventoryStatus -Row $row -Status Failed -Notes 'Row has no DriveId/ItemId - re-run discovery for this file.' | Out-Null
            Write-PubFileResult -Outcome Failed -FileName ([string] $row.FileName) -Detail 'missing DriveId/ItemId'
            continue
        }

        try {
            $folder = Split-Path -Parent $localPath
            if (-not (Test-Path -LiteralPath $folder)) { New-Item -Path $folder -ItemType Directory -Force | Out-Null }

            if ($localPath.Length -ge 250) {
                Write-PubLog -Level Warn -Message ('Local path is {0} characters - long-path support may be needed: {1}' -f $localPath.Length, $localPath)
            }

            $uri = 'drives/{0}/items/{1}/content' -f $row.DriveId, $row.ItemId
            Invoke-PubGraph -Uri $uri -Method GET -OutputFilePath $localPath | Out-Null

            if (-not (Test-Path -LiteralPath $localPath)) { throw 'Graph reported success but no file was written.' }

            $row.LocalPath = $localPath
            Set-PubInventoryStatus -Row $row -Status Downloaded -Notes '' | Out-Null
            $succeeded++
            Write-PubFileResult -Outcome Downloaded -FileName ([string] $row.FileName) -Detail $localPath
        } catch {
            $failed++
            $message = Get-PubGraphErrorMessage -ErrorRecord $_
            Set-PubInventoryStatus -Row $row -Status Failed -Notes ('Download failed: {0}' -f $message) | Out-Null
            Write-PubFileResult -Outcome Failed -FileName ([string] $row.FileName) -Detail $message
        }
    }

    Write-Progress -Activity 'Downloading Publisher files' -Completed

    Write-PubPhaseSummary -Phase 'Download' -Attempted $rows.Count -Succeeded $succeeded -Failed $failed -Skipped $skipped `
                          -ExtraLines @(('Working folder : {0}' -f (Get-PubWorkingFolder -SubFolder 'originals' -Config $Config)))

    return $rows
}

function Invoke-PubConvertBatch {
    <#
    .SYNOPSIS
        Runs one batch of conversions in a Windows PowerShell 5.1 child process.

    .DESCRIPTION
        Results are read back from a JSONL file the child appends to as each
        file finishes, so a batch that has to be killed still reports
        everything that completed before the hang.

    .OUTPUTS
        The parsed result objects for the batch.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Jobs,
        [Parameter(Mandatory)] [string] $ExistingPdfAction,
        [int] $PerFileTimeoutSeconds = 300
    )

    $jobs = @($Jobs)
    if ($jobs.Count -eq 0) { return @() }

    $temp       = Get-PubWorkingFolder -SubFolder 'converted'
    $stamp      = (Get-Date -Format 'yyyyMMdd_HHmmss_fff')
    $jobFile    = Join-Path $temp ('.convert_jobs_{0}.json'    -f $stamp)
    $resultFile = Join-Path $temp ('.convert_results_{0}.jsonl' -f $stamp)

    $powerShell = Get-PubWindowsPowerShellPath
    if (-not $powerShell) {
        Write-PubLog -Level Error -Message 'Windows PowerShell 5.1 was not found - cannot run the COM conversion.'
        return @()
    }

    try {
        ,$jobs | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $jobFile -Encoding UTF8
        New-Item -Path $resultFile -ItemType File -Force | Out-Null

        $arguments = @(
            '-NoProfile'
            '-NonInteractive'
            '-ExecutionPolicy', 'Bypass'
            '-File', ('"{0}"' -f $script:ConverterScript)
            '-JobFile', ('"{0}"' -f $jobFile)
            '-ResultFile', ('"{0}"' -f $resultFile)
            '-ExistingPdfAction', $ExistingPdfAction
        )

        $logFile = Get-PubLogFile
        if ($logFile) { $arguments += @('-LogFile', ('"{0}"' -f $logFile)) }

        $timeoutMs = $PerFileTimeoutSeconds * 1000 * $jobs.Count
        Write-PubLog -Level Debug -Message ('Starting conversion child process for {0} file(s).' -f $jobs.Count)

        $process = Start-Process -FilePath $powerShell -ArgumentList $arguments -PassThru -WindowStyle Hidden
        if (-not $process.WaitForExit($timeoutMs)) {
            Write-PubLog -Level Error -Message ('Conversion batch exceeded {0}s - stopping it. A corrupt or password-protected file may have hung Publisher.' -f [int] ($timeoutMs / 1000))
            try { $process.Kill() } catch { }
            Start-Sleep -Seconds 2
            Write-PubLog -Level Warn -Message 'Check Task Manager for an orphaned MSPUB.EXE if the next batch also stalls.'
        } elseif ($process.ExitCode -ne 0) {
            Write-PubLog -Level Warn -Message ('Conversion child process exited with code {0} - per-file results below.' -f $process.ExitCode)
        }

        $results = New-Object System.Collections.Generic.List[object]
        if (Test-Path -LiteralPath $resultFile) {
            foreach ($line in (Get-Content -LiteralPath $resultFile)) {
                if ([string]::IsNullOrWhiteSpace($line)) { continue }
                try { $results.Add(($line | ConvertFrom-Json)) }
                catch { Write-PubLog -Level Debug -Message ('Unreadable result line skipped: {0}' -f $line) }
            }
        }

        return $results.ToArray()
    } finally {
        foreach ($file in @($jobFile, $resultFile)) {
            if (Test-Path -LiteralPath $file) { Remove-Item -LiteralPath $file -Force -ErrorAction SilentlyContinue }
        }
    }
}

function Invoke-PubConvert {
    <#
    .SYNOPSIS
        Menu option 8 - converts downloaded .pub files to PDF.

    .DESCRIPTION
        Splits the work into batches (default 25) so one hung file costs one
        batch rather than the whole run, then feeds each per-file result back
        into the row's Status, Notes and PdfPath.

        The existing-PDF rule comes from config (Skip / Overwrite / Version),
        set from the menu, not hardcoded (brief section 5.4).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        $Rows,

        $Config,
        [ValidateSet('Skip', 'Overwrite', 'Version')]
        [string] $ExistingPdfAction,
        [int]    $BatchSize = 25,
        [int]    $PerFileTimeoutSeconds = 300
    )

    if (-not $Config) { $Config = Get-PubConfig }
    if ([string]::IsNullOrWhiteSpace($ExistingPdfAction)) {
        $ExistingPdfAction = [string] $Config['ExistingPdfAction']
        if ([string]::IsNullOrWhiteSpace($ExistingPdfAction)) { $ExistingPdfAction = 'Version' }
    }

    $rows = @($Rows)
    if ($rows.Count -eq 0) {
        Write-PubLog -Level Warn -Message 'No rows selected - nothing to convert.'
        return $rows
    }

    if (-not (Test-PubPublisherAvailable)) {
        Write-PubLog -Level Error -Message 'Conversion cannot run on this host. This is a hard prerequisite - see README section "Where each phase can run".'
        return $rows
    }

    # Only rows whose .pub is actually on disk can be converted.
    $convertible = New-Object System.Collections.Generic.List[object]
    $notReady    = 0

    foreach ($row in $rows) {
        $localPath = [string] $row.LocalPath
        if ([string]::IsNullOrWhiteSpace($localPath)) {
            $localPath = Get-PubLocalPath -Row $row -Kind Original -Config $Config
        }

        if (Test-Path -LiteralPath $localPath) {
            $row.LocalPath = $localPath
            $convertible.Add($row)
        } else {
            $notReady++
            Set-PubInventoryStatus -Row $row -Status Failed -Notes 'Not downloaded - run the download step (option 7) first.' | Out-Null
            Write-PubFileResult -Outcome Failed -FileName ([string] $row.FileName) -Detail 'no local copy'
        }
    }

    if ($convertible.Count -eq 0) {
        Write-PubPhaseSummary -Phase 'Convert' -Attempted $rows.Count -Succeeded 0 -Failed $notReady
        return $rows
    }

    Write-PubLog -Level Info -Message ('=== Phase 2b: Convert ({0} file(s), existing PDF = {1}) ===' -f $convertible.Count, $ExistingPdfAction)

    $byUniqueId = @{}
    $jobs       = New-Object System.Collections.Generic.List[object]

    foreach ($row in $convertible) {
        $key = [string] $row.UniqueId
        if ([string]::IsNullOrWhiteSpace($key)) { $key = [string] $row.ItemId }
        $byUniqueId[$key] = $row

        $jobs.Add([pscustomobject] @{
            UniqueId    = $key
            SourcePath  = [string] $row.LocalPath
            PdfPath     = (Get-PubLocalPath -Row $row -Kind Pdf -Config $Config)
            FileName    = [string] $row.FileName
            SiteUrl     = [string] $row.SiteUrl
            LibraryName = [string] $row.LibraryName
            FolderPath  = [string] $row.FolderPath
        })
    }

    $converted = 0
    $failed    = $notReady
    $skipped   = 0
    $processed = 0

    $jobArray = $jobs.ToArray()

    for ($offset = 0; $offset -lt $jobArray.Count; $offset += $BatchSize) {
        $last  = [math]::Min($offset + $BatchSize - 1, $jobArray.Count - 1)
        $batch = @($jobArray[$offset..$last])

        Write-Progress -Activity 'Converting Publisher files to PDF' `
                       -Status ('{0} of {1}' -f $processed, $jobArray.Count) `
                       -PercentComplete ([int] (($processed / $jobArray.Count) * 100))

        $results = Invoke-PubConvertBatch -Jobs $batch -ExistingPdfAction $ExistingPdfAction -PerFileTimeoutSeconds $PerFileTimeoutSeconds

        $reported = @{}
        foreach ($result in $results) {
            $key = [string] $result.UniqueId
            $reported[$key] = $true

            if (-not $byUniqueId.ContainsKey($key)) { continue }
            $row = $byUniqueId[$key]

            switch ([string] $result.Status) {
                'Converted' {
                    $row.PdfPath = [string] $result.PdfPath
                    Set-PubInventoryStatus -Row $row -Status Converted -Notes ([string] $result.Message) | Out-Null
                    $converted++
                    Write-PubFileResult -Outcome Converted -FileName ([string] $row.FileName) -Detail ([string] $result.PdfPath)
                }
                'Skipped' {
                    $row.PdfPath = [string] $result.PdfPath
                    Set-PubInventoryStatus -Row $row -Status Skipped -Notes ([string] $result.Message) | Out-Null
                    $skipped++
                    Write-PubFileResult -Outcome Skipped -FileName ([string] $row.FileName) -Detail ([string] $result.Message)
                }
                default {
                    Set-PubInventoryStatus -Row $row -Status Failed -Notes ([string] $result.Message) | Out-Null
                    $failed++
                    Write-PubFileResult -Outcome Failed -FileName ([string] $row.FileName) -Detail ([string] $result.Message)
                }
            }
        }

        # Anything in the batch with no result line was cut short - almost
        # always the file that hung Publisher.
        foreach ($job in $batch) {
            if ($reported.ContainsKey([string] $job.UniqueId)) { continue }
            if (-not $byUniqueId.ContainsKey([string] $job.UniqueId)) { continue }

            $row = $byUniqueId[[string] $job.UniqueId]
            Set-PubInventoryStatus -Row $row -Status Failed -Notes 'No result returned - the conversion process was stopped (possible corrupt or password-protected file).' | Out-Null
            $failed++
            Write-PubFileResult -Outcome Failed -FileName ([string] $row.FileName) -Detail 'conversion process stopped before this file reported'
        }

        $processed += $batch.Count
    }

    Write-Progress -Activity 'Converting Publisher files to PDF' -Completed

    Write-PubPhaseSummary -Phase 'Convert' -Attempted $rows.Count -Succeeded $converted -Failed $failed -Skipped $skipped `
                          -ExtraLines @(('PDF folder : {0}' -f (Get-PubWorkingFolder -SubFolder 'converted' -Config $Config)))

    return $rows
}

Export-ModuleMember -Function @(
    'Get-PubWindowsPowerShellPath'
    'Test-PubPublisherAvailable'
    'ConvertTo-PubSafeSegment'
    'Get-PubLocalPath'
    'Select-PubInventoryRow'
    'Invoke-PubDownload'
    'Invoke-PubConvertBatch'
    'Invoke-PubConvert'
)
