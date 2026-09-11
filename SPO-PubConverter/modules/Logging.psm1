<#
.SYNOPSIS
    Shared logging helper for the SharePoint Publisher File Converter.

.DESCRIPTION
    Every run writes a timestamped log file into /logs. Console output is
    colourised by level; the log file gets the same lines without colour so the
    file is greppable after the fact.

    A PowerShell transcript is also started alongside the structured log where
    the host supports it, so anything written outside Write-PubLog (native
    command output, unexpected exceptions) is still captured.

    Compatible with Windows PowerShell 5.1 and PowerShell 7.
#>

Set-StrictMode -Version 2.0

$script:LogFile        = $null
$script:TranscriptFile = $null
$script:TranscriptOn   = $false
$script:LogRoot        = $null

function Get-PubLogRoot {
    <#
    .SYNOPSIS
        Returns the folder logs are written to, creating it if required.
    #>
    [CmdletBinding()]
    param(
        [string] $LogRoot
    )

    if ($LogRoot) {
        $root = $LogRoot
    } elseif ($script:LogRoot) {
        $root = $script:LogRoot
    } else {
        # modules\Logging.psm1 -> project root -> \logs
        $root = Join-Path (Split-Path -Parent $PSScriptRoot) 'logs'
    }

    if (-not (Test-Path -LiteralPath $root)) {
        New-Item -Path $root -ItemType Directory -Force | Out-Null
    }

    return (Resolve-Path -LiteralPath $root).Path
}

function Initialize-PubLogging {
    <#
    .SYNOPSIS
        Opens a new timestamped log (and transcript) for this run.

    .PARAMETER Name
        Short name used in the log file name, e.g. 'Discovery'.

    .PARAMETER LogRoot
        Override the default <project>\logs folder.

    .PARAMETER NoTranscript
        Skip Start-Transcript (useful when a transcript is already running).
    #>
    [CmdletBinding()]
    param(
        [string] $Name = 'Session',
        [string] $LogRoot,
        [switch] $NoTranscript
    )

    $script:LogRoot = Get-PubLogRoot -LogRoot $LogRoot
    $stamp          = Get-Date -Format 'yyyy-MM-dd_HHmm'
    $safeName       = ($Name -replace '[^A-Za-z0-9_\-]', '_')

    $script:LogFile = Join-Path $script:LogRoot ("PubConverter_{0}_{1}.log" -f $safeName, $stamp)

    if (-not (Test-Path -LiteralPath $script:LogFile)) {
        New-Item -Path $script:LogFile -ItemType File -Force | Out-Null
    }

    if (-not $NoTranscript -and -not $script:TranscriptOn) {
        $script:TranscriptFile = Join-Path $script:LogRoot ("Transcript_{0}_{1}.log" -f $safeName, $stamp)
        try {
            Start-Transcript -Path $script:TranscriptFile -Append -ErrorAction Stop | Out-Null
            $script:TranscriptOn = $true
        } catch {
            # Some hosts (ISE, constrained runspaces) do not support transcripts.
            # The structured log is the source of truth, so this is not fatal.
            $script:TranscriptFile = $null
        }
    }

    Write-PubLog -Level Info -Message ("Log started: {0}" -f $script:LogFile) -NoConsole
    Write-PubLog -Level Info -Message ("Host: {0} / PowerShell {1} / {2}" -f $env:COMPUTERNAME, $PSVersionTable.PSVersion, $PSVersionTable.PSEdition) -NoConsole

    return $script:LogFile
}

function Stop-PubLogging {
    <#
    .SYNOPSIS
        Closes the transcript opened by Initialize-PubLogging.
    #>
    [CmdletBinding()]
    param()

    if ($script:TranscriptOn) {
        try { Stop-Transcript | Out-Null } catch { }
        $script:TranscriptOn = $false
    }
}

function Get-PubLogFile {
    <#
    .SYNOPSIS
        Returns the path of the log file for the current run (may be $null).
    #>
    [CmdletBinding()]
    param()
    return $script:LogFile
}

function Write-PubLog {
    <#
    .SYNOPSIS
        Writes one line to the console and to the run log.

    .PARAMETER Level
        Info | Success | Warn | Error | Debug.

    .PARAMETER NoConsole
        Write to the log file only.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [AllowEmptyString()]
        [string] $Message,

        [ValidateSet('Info', 'Success', 'Warn', 'Error', 'Debug')]
        [string] $Level = 'Info',

        [switch] $NoConsole
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line      = '{0} [{1}] {2}' -f $timestamp, $Level.ToUpper().PadRight(7), $Message

    if (-not $script:LogFile) {
        # Logging was never initialised (module used standalone) - open a default log.
        Initialize-PubLogging -Name 'Session' -NoTranscript | Out-Null
    }

    try {
        Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8 -ErrorAction Stop
    } catch {
        # Never let a logging failure take down a batch.
        Write-Host ("(log write failed: {0})" -f $_.Exception.Message) -ForegroundColor DarkGray
    }

    if ($NoConsole) { return }

    switch ($Level) {
        'Success' { Write-Host $Message -ForegroundColor Green }
        'Warn'    { Write-Host $Message -ForegroundColor Yellow }
        'Error'   { Write-Host $Message -ForegroundColor Red }
        'Debug'   { Write-Verbose $Message }
        default   { Write-Host $Message }
    }
}

function Write-PubFileResult {
    <#
    .SYNOPSIS
        Writes the standard one-line-per-file summary required by the brief.

    .DESCRIPTION
        Used by every phase so the log reads consistently:
            2026-09-11 10:31:02 [INFO   ] CONVERTED | Newsletter.pub | -> C:\...\converted\Newsletter.pdf
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Found', 'Downloaded', 'Converted', 'Uploaded', 'Skipped', 'Deleted', 'Failed')]
        [string] $Outcome,

        [Parameter(Mandatory)]
        [string] $FileName,

        [string] $Detail
    )

    $level = 'Info'
    if ($Outcome -eq 'Failed')  { $level = 'Error' }
    if ($Outcome -eq 'Skipped') { $level = 'Warn' }

    $message = '{0} | {1}' -f $Outcome.ToUpper(), $FileName
    if ($Detail) { $message = '{0} | {1}' -f $message, $Detail }

    Write-PubLog -Level $level -Message $message
}

function Write-PubPhaseSummary {
    <#
    .SYNOPSIS
        Prints the end-of-phase run summary required by section 9 of the brief.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string] $Phase,

        [int] $Attempted = 0,
        [int] $Succeeded = 0,
        [int] $Failed    = 0,
        [int] $Skipped   = 0,

        [string[]] $ExtraLines = @()
    )

    Write-PubLog -Level Info -Message ''
    Write-PubLog -Level Info -Message ('--- {0} summary ---' -f $Phase)
    Write-PubLog -Level Info -Message ('  Attempted : {0}' -f $Attempted)
    Write-PubLog -Level Success -Message ('  Succeeded : {0}' -f $Succeeded)
    if ($Skipped -gt 0) { Write-PubLog -Level Warn -Message ('  Skipped   : {0}' -f $Skipped) }
    if ($Failed -gt 0)  { Write-PubLog -Level Error -Message ('  Failed    : {0}' -f $Failed) }
    else                { Write-PubLog -Level Info -Message  ('  Failed    : 0') }

    foreach ($line in $ExtraLines) {
        if ($line) { Write-PubLog -Level Info -Message ('  {0}' -f $line) }
    }

    if ($script:LogFile) {
        Write-PubLog -Level Info -Message ('  Log       : {0}' -f $script:LogFile)
    }
    Write-PubLog -Level Info -Message ''
}

function Get-PubRecentLogFile {
    <#
    .SYNOPSIS
        Returns the most recently written log file, or $null if there are none.
    #>
    [CmdletBinding()]
    param(
        [string] $LogRoot
    )

    $root = Get-PubLogRoot -LogRoot $LogRoot
    return Get-ChildItem -LiteralPath $root -Filter 'PubConverter_*.log' -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
}

function Show-PubRecentLog {
    <#
    .SYNOPSIS
        Tails the most recent log file in-console (menu option 11).
    #>
    [CmdletBinding()]
    param(
        [int]    $Lines = 60,
        [string] $LogRoot
    )

    $log = Get-PubRecentLogFile -LogRoot $LogRoot
    if (-not $log) {
        Write-Host 'No log files found yet.' -ForegroundColor Yellow
        return
    }

    Write-Host ''
    Write-Host ("Last {0} lines of {1}" -f $Lines, $log.FullName) -ForegroundColor Cyan
    Write-Host ('-' * 78) -ForegroundColor DarkGray

    Get-Content -LiteralPath $log.FullName -Tail $Lines | ForEach-Object {
        $colour = 'Gray'
        if ($_ -match '\[ERROR\s*\]')   { $colour = 'Red' }
        elseif ($_ -match '\[WARN\s*\]') { $colour = 'Yellow' }
        elseif ($_ -match '\[SUCCESS\]') { $colour = 'Green' }
        Write-Host $_ -ForegroundColor $colour
    }

    Write-Host ('-' * 78) -ForegroundColor DarkGray
}

Export-ModuleMember -Function @(
    'Initialize-PubLogging'
    'Stop-PubLogging'
    'Get-PubLogFile'
    'Get-PubLogRoot'
    'Write-PubLog'
    'Write-PubFileResult'
    'Write-PubPhaseSummary'
    'Get-PubRecentLogFile'
    'Show-PubRecentLog'
)
