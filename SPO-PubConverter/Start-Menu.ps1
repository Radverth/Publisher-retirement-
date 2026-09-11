<#
.SYNOPSIS
    SharePoint Publisher File Converter - tenant-wide .pub discovery, PDF
    conversion and re-upload, driven from one numbered console menu.

.DESCRIPTION
    Entry point for the whole tool. There are no flags to remember: run this
    script, pick numbered options in the order a first-time operator would
    naturally use them, and the menu redisplays after each action so several
    phases can be run back to back in one session.

        1-3   Setup      - app registration, certificate, connection test
        4-5   Discovery  - crawl for .pub files, export the CSV inventory
        6-8   Conversion - load the CSV, download, convert to PDF
        9     Publish    - upload the PDFs back to SharePoint
        10-13 Utilities  - full pipeline, log, working folder, settings

    Every read/scan action runs without confirmation. Every write action
    (upload, overwrite, deleting a source .pub) requires an explicit typed
    confirmation.

.PARAMETER WorkingFolder
    Override the working folder for this session (also settable at option 13).

.PARAMETER NoRun
    Load the functions without showing the menu. Used by tests/Run-Tests.ps1 to
    render and assert on the menu without going interactive.

.EXAMPLE
    .\Start-Menu.ps1

.NOTES
    Runs under Windows PowerShell 5.1 or PowerShell 7. The PDF conversion step
    always runs as a Windows PowerShell 5.1 child process on a host with
    Microsoft Publisher installed - see the README.
#>
[CmdletBinding()]
param(
    [string] $WorkingFolder,
    [switch] $NoRun
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$script:ModuleFolder = Join-Path $PSScriptRoot 'modules'

foreach ($module in @('Logging', 'Config', 'Graph', 'AppRegistration', 'PnP', 'Discovery', 'Convert', 'Upload')) {
    Import-Module (Join-Path $script:ModuleFolder ("{0}.psm1" -f $module)) -Force -DisableNameChecking
}

# ---------------------------------------------------------------------------
# Session state
# ---------------------------------------------------------------------------
$script:Config        = Get-PubConfig
$script:Inventory     = @()
$script:InventoryPath = ''
$script:Selection     = @()
$script:Version       = 'v1.0'
$script:VerifiedRoute = ''   # set by option 3 once the route has really been tested

if ($WorkingFolder) {
    $script:Config['DefaultWorkingFolder'] = $WorkingFolder
    Save-PubConfig -Config $script:Config | Out-Null
}

# ---------------------------------------------------------------------------
# Console helpers
# ---------------------------------------------------------------------------
function Wait-PubKeyPress {
    [CmdletBinding()]
    param(
        [string] $Message = 'Press any key to return to the menu...'
    )

    Write-Host ''
    Write-Host $Message -ForegroundColor DarkGray

    try {
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    } catch {
        # Hosts without a raw UI (ISE, redirected input) fall back to Enter.
        $null = Read-Host
    }
}

function Read-PubMenuChoice {
    <#
    .SYNOPSIS
        Prompts until the operator types one of the valid option numbers.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string[]] $Valid,
        [string] $Prompt = ' Select an option'
    )

    while ($true) {
        $answer = Read-Host $Prompt
        if ($null -ne $answer) { $answer = $answer.Trim().TrimEnd(')') }

        if ($Valid -contains $answer) { return $answer }

        # Technicians type these out of habit - treat them as "0" rather than
        # scolding, since 0 always means back or exit.
        if ($Valid -contains '0' -and $answer -match '^(q|quit|exit|b|back|x)$') { return '0' }

        if ([string]::IsNullOrWhiteSpace($answer)) {
            Write-Host ' Enter one of the numbers listed (0 goes back).' -ForegroundColor Yellow
        } else {
            Write-Host (' "{0}" is not an option here. Valid choices: {1}' -f $answer, ($Valid -join ', ')) -ForegroundColor Yellow
        }
    }
}

function Confirm-PubAction {
    <#
    .SYNOPSIS
        Y/N confirmation for a write action. Defaults to No.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Question
    )

    $answer = Read-Host ('{0} (Y/N)' -f $Question)
    return ($answer -match '^[Yy]')
}

function Get-PubMenuWidth {
    [CmdletBinding()]
    param()
    return 78
}

function Write-PubMenuRule {
    [CmdletBinding()]
    param([string] $Character = '-', [string] $Colour = 'DarkGray')
    Write-Host ($Character * (Get-PubMenuWidth)) -ForegroundColor $Colour
}

function Format-PubMenuStatusLine {
    <#
    .SYNOPSIS
        Composes one 'Label : value' line of the state block.

    .DESCRIPTION
        Formatting is separated from colouring so the composed line can be
        width-checked by the tests - a console menu that wraps is unreadable,
        and the colour segments hide the true line length otherwise.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Label,
        [Parameter(Mandatory)] [AllowEmptyString()] [string] $Value,
        [switch] $Continuation
    )

    return ((Get-PubMenuStatusPrefix -Label $Label -Continuation:$Continuation) + $Value)
}

function Get-PubMenuStatusPrefix {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Label,
        [switch] $Continuation
    )

    if ($Continuation) { return (' ' * 19) }
    return ('  {0} : ' -f $Label.PadRight(14))
}

function Write-PubMenuStatusLine {
    <#
    .SYNOPSIS
        One aligned 'Label : value' line in the menu's state block.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Label,
        [Parameter(Mandatory)] [AllowEmptyString()] [string] $Value,
        [string] $Colour = 'Gray',
        [switch] $Continuation
    )

    Write-Host (Get-PubMenuStatusPrefix -Label $Label -Continuation:$Continuation) -NoNewline
    Write-Host $Value -ForegroundColor $Colour
}

function Write-PubMenuWrappedValue {
    <#
    .SYNOPSIS
        Writes a 'Label : a  |  b  |  c' line, wrapping onto aligned
        continuation lines so it never runs past 80 columns.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Label,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [string[]] $Parts,
        [string] $Colour = 'Gray',
        [int]    $MaxWidth = 57,
        [switch] $Continuation
    )

    if (-not $Parts -or $Parts.Count -eq 0) { return }

    $line  = ''
    $first = -not $Continuation

    foreach ($part in $Parts) {
        $candidate = $part
        if ($line) { $candidate = '{0}  |  {1}' -f $line, $part }

        if ($candidate.Length -gt $MaxWidth -and $line) {
            if ($first) { Write-PubMenuStatusLine -Label $Label -Value $line -Colour $Colour; $first = $false }
            else        { Write-PubMenuStatusLine -Label ' ' -Value $line -Colour $Colour -Continuation }
            $line = $part
        } else {
            $line = $candidate
        }
    }

    if ($line) {
        if ($first) { Write-PubMenuStatusLine -Label $Label -Value $line -Colour $Colour }
        else        { Write-PubMenuStatusLine -Label ' ' -Value $line -Colour $Colour -Continuation }
    }
}

function Get-PubMenuNoteColumn {
    [CmdletBinding()]
    param()
    # 62 clears the longest label (option 9) by two spaces; notes are kept to
    # 14 characters or fewer so the whole line fits an 80-column console.
    return 62
}

function Format-PubMenuOptionLine {
    <#
    .SYNOPSIS
        Composes one option line, with its note in the second column.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Number,
        [Parameter(Mandatory)] [string] $Label,
        [string] $Note
    )

    $text = '  {0}) {1}' -f $Number.PadLeft(3), $Label
    if ([string]::IsNullOrWhiteSpace($Note)) { return $text }

    return ($text.PadRight((Get-PubMenuNoteColumn)) + $Note)
}

function Write-PubMenuOption {
    <#
    .SYNOPSIS
        One menu option, with its status note aligned in a second column.

    .PARAMETER Note
        Short status for this option - what is done, what is waiting, or why
        it cannot run yet. This is what stops the operator guessing which step
        they are on.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Number,
        [Parameter(Mandatory)] [string] $Label,
        [string] $Note,
        [ValidateSet('Normal', 'Done', 'Next', 'Waiting', 'Blocked')]
        [string] $State = 'Normal'
    )

    $line   = Format-PubMenuOptionLine -Number $Number -Label $Label -Note $Note
    $column = Get-PubMenuNoteColumn

    $labelColour = 'White'
    if ($State -eq 'Blocked') { $labelColour = 'DarkGray' }
    if ($State -eq 'Next')    { $labelColour = 'Cyan' }

    if ([string]::IsNullOrWhiteSpace($Note) -or $line.Length -le $column) {
        Write-Host $line -ForegroundColor $labelColour
        return
    }

    $noteColour = 'DarkGray'
    switch ($State) {
        'Done'    { $noteColour = 'Green' }
        'Next'    { $noteColour = 'Cyan' }
        'Waiting' { $noteColour = 'Yellow' }
        'Blocked' { $noteColour = 'Red' }
    }

    Write-Host $line.Substring(0, $column) -ForegroundColor $labelColour -NoNewline
    Write-Host $line.Substring($column) -ForegroundColor $noteColour
}

function Get-PubMenuState {
    <#
    .SYNOPSIS
        Everything the menu needs to describe itself, computed once per render.

    .DESCRIPTION
        Deliberately cheap: no network calls. The discovery route is inferred
        from configuration and the cached PnP probe, and replaced with the
        verified result once option 3 has actually tested it.
    #>
    [CmdletBinding()]
    param()

    $config = $script:Config

    $state = @{
        HasApp        = -not [string]::IsNullOrWhiteSpace([string] $config['AppId'])
        HasCert       = -not [string]::IsNullOrWhiteSpace([string] $config['CertificateThumbprint'])
        CertState     = 'Unknown'
        CertMessage   = ''
        HasInventory  = ($script:Inventory -and @($script:Inventory).Count -gt 0)
        Stats         = $null
        TotalRows     = @($script:Inventory).Count
        RowsToProcess = 0
        Pending       = 0
        Downloaded    = 0
        Converted     = 0
        Uploaded      = 0
        Failed        = 0
        Publisher     = (Test-PubPublisherAvailable -Quiet)
    }

    $expiry              = Test-PubCertificateExpiry -Config $config -Quiet
    $state['CertState']  = $expiry.State
    $state['CertMessage'] = $expiry.Message

    if ($state['HasInventory']) {
        $rows  = Get-PubRowsToProcess
        $stats = Get-PubInventoryStatistic -Rows $rows

        $state['Stats']         = $stats
        $state['RowsToProcess'] = $stats.Total
        $state['Pending']       = $stats.Pending
        $state['Downloaded']    = $stats.Downloaded
        $state['Converted']     = $stats.Converted
        $state['Uploaded']      = $stats.Uploaded
        $state['Failed']        = $stats.Failed
    }

    return $state
}

function Get-PubDiscoveryRouteLabel {
    <#
    .SYNOPSIS
        Describes how discovery will find sites, in plain words.
    #>
    [CmdletBinding()]
    param()

    if ($script:VerifiedRoute) { return $script:VerifiedRoute }

    if ([string]::IsNullOrWhiteSpace([string] $script:Config['AppId'])) {
        return 'set up tenant access first (option 1)'
    }

    $method = [string] $script:Config['EnumerationMethod']
    if ([string]::IsNullOrWhiteSpace($method)) { $method = 'Auto' }

    if ([string] $script:Config['AuthMethod'] -eq 'SitesSelected') {
        return 'only the sites granted to this app'
    }

    if ($method -ne 'Graph' -and ([string] $script:Config['AuthMethod'] -eq 'TenantAdmin') -and (Test-PubPnPAvailable -Quiet)) {
        return 'tenant admin list - finds every site (test with option 3)'
    }

    return 'Graph search - may miss some sites (see README)'
}

function Get-PubNextStep {
    <#
    .SYNOPSIS
        Works out the one thing the operator should do next.

    .OUTPUTS
        A hashtable with Option and Text, or Option = '' when there is nothing
        outstanding.
    #>
    [CmdletBinding()]
    param($State)

    if (-not $State) { $State = Get-PubMenuState }

    if (-not $State['HasApp']) {
        return @{ Option = '1'; Text = 'set up access to the tenant (first time here)' }
    }
    if (-not $State['HasCert']) {
        return @{ Option = '2'; Text = 'create the certificate this tool signs in with' }
    }
    if ($State['CertState'] -eq 'Expired') {
        return @{ Option = '2'; Text = 'certificate expired - replace it before anything else' }
    }

    if (-not $State['HasInventory']) {
        $last = [string] $script:Config['LastInventoryCsv']
        if (-not [string]::IsNullOrWhiteSpace($last) -and (Test-Path -LiteralPath $last)) {
            return @{ Option = '6'; Text = 'load the CSV from your last run and carry on' }
        }
        return @{ Option = '4'; Text = 'scan SharePoint to find the .pub files' }
    }

    if ($State['Pending'] -gt 0)    { return @{ Option = '7'; Text = ('download the {0} file(s) not yet on this machine' -f $State['Pending']) } }
    if ($State['Downloaded'] -gt 0) { return @{ Option = '8'; Text = ('convert the {0} downloaded file(s) to PDF' -f $State['Downloaded']) } }
    if ($State['Converted'] -gt 0)  { return @{ Option = '9'; Text = ('upload the {0} finished PDF(s) back to SharePoint' -f $State['Converted']) } }

    if ($State['Failed'] -gt 0) {
        return @{ Option = '6'; Text = ('everything else is done - {0} row(s) failed, select them to retry' -f $State['Failed']) }
    }

    return @{ Option = ''; Text = 'all rows in this inventory are uploaded - nothing outstanding' }
}

function Get-PubInventoryLabel {
    <#
    .SYNOPSIS
        The inventory line for the menu header.
    #>
    [CmdletBinding()]
    param()

    if ($script:Inventory -and @($script:Inventory).Count -gt 0) {
        $name = Split-Path -Leaf $script:InventoryPath
        if ([string]::IsNullOrWhiteSpace($name)) { $name = 'unsaved scan results' }
        return ('{0} ({1} files)' -f $name, @($script:Inventory).Count)
    }

    $last = [string] $script:Config['LastInventoryCsv']
    if (-not [string]::IsNullOrWhiteSpace($last)) {
        if (Test-Path -LiteralPath $last) { return ('{0} (not loaded yet - option 6)' -f (Split-Path -Leaf $last)) }
        return ('{0} (file no longer present)' -f (Split-Path -Leaf $last))
    }

    return 'none yet - run option 4 to scan'
}

function Show-PubMainMenu {
    <#
    .SYNOPSIS
        Renders the main menu: current state, the next step, then the options.
    #>
    [CmdletBinding()]
    param(
        [switch] $NoClear
    )

    $state = Get-PubMenuState
    $next  = Get-PubNextStep -State $state

    if (-not $NoClear) {
        try { Clear-Host } catch { }
    }

    $width = Get-PubMenuWidth

    Write-PubMenuRule -Character '=' -Colour Cyan
    Write-Host ('   SharePoint Publisher File Converter'.PadRight($width - 6) + $script:Version) -ForegroundColor Cyan
    Write-PubMenuRule -Character '=' -Colour Cyan

    # ---- where things stand ----
    $tenant = [string] $script:Config['TenantDomain']
    if ([string]::IsNullOrWhiteSpace($tenant)) { $tenant = [string] $script:Config['TenantId'] }
    if ([string]::IsNullOrWhiteSpace($tenant)) { $tenant = 'not set up yet' }
    Write-PubMenuStatusLine -Label 'Tenant' -Value $tenant

    $accessValue  = 'not set up yet - start at option 1'
    $accessColour = 'Yellow'
    if ($state['HasApp'] -and $state['HasCert']) {
        switch ($state['CertState']) {
            'Valid'    { $accessValue = ('ready - certificate expires {0}' -f $script:Config['CertificateExpiry']); $accessColour = 'Green' }
            'Expiring' { $accessValue = ('certificate expires {0} - renew soon (option 2)' -f $script:Config['CertificateExpiry']); $accessColour = 'Yellow' }
            'Expired'  { $accessValue = ('CERTIFICATE EXPIRED {0} - run option 2' -f $script:Config['CertificateExpiry']); $accessColour = 'Red' }
            default    { $accessValue = 'configured - expiry unknown'; $accessColour = 'Yellow' }
        }
    } elseif ($state['HasApp']) {
        $accessValue = 'app registered, certificate still needed - option 2'
    }
    Write-PubMenuStatusLine -Label 'Tenant access' -Value $accessValue -Colour $accessColour
    Write-PubMenuStatusLine -Label 'Site discovery' -Value (Get-PubDiscoveryRouteLabel)
    Write-PubMenuStatusLine -Label 'File list' -Value (Get-PubInventoryLabel)

    if ($state['HasInventory']) {
        $stats = $state['Stats']

        # Only the states that actually have rows, wrapped so a busy inventory
        # still fits an 80-column console.
        $parts = New-Object System.Collections.Generic.List[string]
        if ($stats.Pending -gt 0)    { $parts.Add(('{0} to download' -f $stats.Pending)) }
        if ($stats.Downloaded -gt 0) { $parts.Add(('{0} to convert'  -f $stats.Downloaded)) }
        if ($stats.Converted -gt 0)  { $parts.Add(('{0} to upload'   -f $stats.Converted)) }
        if ($stats.Uploaded -gt 0)   { $parts.Add(('{0} uploaded'    -f $stats.Uploaded)) }
        if ($stats.Failed -gt 0)     { $parts.Add(('{0} failed'      -f $stats.Failed)) }
        if ($stats.Skipped -gt 0)    { $parts.Add(('{0} skipped'     -f $stats.Skipped)) }
        if ($parts.Count -eq 0)      { $parts.Add('nothing to do') }

        Write-PubMenuWrappedValue -Label ' ' -Parts $parts.ToArray() -Continuation

        if ($script:Selection -and @($script:Selection).Count -gt 0) {
            Write-PubMenuStatusLine -Label 'Working on' -Value ('{0} of {1} row(s) you selected at option 6' -f @($script:Selection).Count, @($script:Inventory).Count) -Colour Yellow
        } else {
            Write-PubMenuStatusLine -Label 'Working on' -Value 'every row in the file list'
        }
    }

    $removeLabel = 'originals kept'
    if ([bool] $script:Config['RemoveSourceAfterUpload']) { $removeLabel = 'ORIGINALS DELETED after upload' }

    Write-PubMenuWrappedValue -Label 'Settings' -Parts @(
        ('existing PDF: {0}' -f $script:Config['ExistingPdfAction'])
        ('name clashes: {0}' -f $script:Config['UploadConflictAction'])
        $removeLabel
    )

    # ---- what to do next ----
    Write-PubMenuRule
    if ($next['Option']) {
        Write-Host '  NEXT: ' -ForegroundColor Cyan -NoNewline
        Write-Host ('option {0} - {1}' -f $next['Option'], $next['Text'])
    } else {
        Write-Host '  NEXT: ' -ForegroundColor Green -NoNewline
        Write-Host $next['Text']
    }

    if (-not $state['Publisher']) {
        Write-Host '  NOTE: ' -ForegroundColor Yellow -NoNewline
        Write-Host 'Publisher not installed here - options 8 and 10 must run elsewhere.'
    }
    Write-PubMenuRule

    # ---- the options ----
    $nextOption = [string] $next['Option']
    function Get-PubOptionState {
        param([string] $Number, [string] $Default = 'Normal')
        if ($Number -eq $nextOption) { return 'Next' }
        return $Default
    }

    Write-Host ''
    Write-Host '  SETUP - do these once, in order' -ForegroundColor White

    $appNote = 'not done yet'; $appState = 'Waiting'
    if ($state['HasApp']) { $appNote = 'done'; $appState = 'Done' }
    Write-PubMenuOption -Number '1' -Label 'Generate or connect Azure AD App Registration' -Note $appNote -State (Get-PubOptionState -Number '1' -Default $appState)

    $certNote = 'not done yet'; $certState = 'Waiting'
    if ($state['HasCert']) {
        switch ($state['CertState']) {
            'Valid'    { $certNote = 'done';          $certState = 'Done' }
            'Expiring' { $certNote = 'expiring soon'; $certState = 'Waiting' }
            'Expired'  { $certNote = 'EXPIRED';       $certState = 'Blocked' }
            default    { $certNote = 'done';          $certState = 'Done' }
        }
    }
    Write-PubMenuOption -Number '2' -Label 'Generate & upload authentication certificate' -Note $certNote -State (Get-PubOptionState -Number '2' -Default $certState)
    Write-PubMenuOption -Number '3' -Label 'Test connection to Microsoft Graph / SharePoint' -Note 'verify setup' -State (Get-PubOptionState -Number '3')

    Write-Host ''
    Write-Host '  DISCOVERY - find the Publisher files' -ForegroundColor White

    $scanNote = 'needs setup'; $scanState = 'Blocked'
    if ($state['HasApp'] -and $state['HasCert']) { $scanNote = ''; $scanState = 'Normal' }
    Write-PubMenuOption -Number '4' -Label 'Scan tenant for Publisher (.pub) files' -Note $scanNote -State (Get-PubOptionState -Number '4' -Default $scanState)

    $exportNote = 'no scan yet'; $exportState = 'Blocked'
    if ($state['HasInventory']) { $exportNote = ('{0} rows' -f $state['TotalRows']); $exportState = 'Normal' }
    Write-PubMenuOption -Number '5' -Label 'Export / re-export scan results to CSV' -Note $exportNote -State (Get-PubOptionState -Number '5' -Default $exportState)

    Write-Host ''
    Write-Host '  CONVERSION - fetch the files and make the PDFs' -ForegroundColor White

    $loadNote = 'none loaded'
    if ($state['HasInventory']) { $loadNote = 'change files' }
    Write-PubMenuOption -Number '6' -Label 'Load a CSV and select files to process' -Note $loadNote -State (Get-PubOptionState -Number '6')

    $downloadNote = 'no file list'; $downloadState = 'Blocked'
    if ($state['HasInventory']) {
        if ($state['Pending'] -gt 0) { $downloadNote = ('{0} to download' -f $state['Pending']); $downloadState = 'Normal' }
        else { $downloadNote = 'none waiting'; $downloadState = 'Normal' }
    }
    Write-PubMenuOption -Number '7' -Label 'Download selected files' -Note $downloadNote -State (Get-PubOptionState -Number '7' -Default $downloadState)

    $convertNote = 'no file list'; $convertState = 'Blocked'
    if (-not $state['Publisher']) {
        $convertNote = 'no Publisher'; $convertState = 'Blocked'
    } elseif ($state['HasInventory']) {
        if ($state['Downloaded'] -gt 0) { $convertNote = ('{0} ready' -f $state['Downloaded']); $convertState = 'Normal' }
        else { $convertNote = 'none ready'; $convertState = 'Normal' }
    }
    Write-PubMenuOption -Number '8' -Label 'Convert downloaded files to PDF' -Note $convertNote -State (Get-PubOptionState -Number '8' -Default $convertState)

    Write-Host ''
    Write-Host '  PUBLISH - put the PDFs back in SharePoint' -ForegroundColor White

    $uploadNote = 'none ready'; $uploadState = 'Blocked'
    if ($state['Converted'] -gt 0) { $uploadNote = ('{0} ready' -f $state['Converted']); $uploadState = 'Normal' }
    Write-PubMenuOption -Number '9' -Label 'Upload converted PDFs to original SharePoint location' -Note $uploadNote -State (Get-PubOptionState -Number '9' -Default $uploadState)

    Write-Host ''
    Write-Host '  UTILITIES' -ForegroundColor White

    $pipelineNote = 'all steps'
    $pipelineState = 'Normal'
    if (-not $state['Publisher']) { $pipelineNote = 'no Publisher'; $pipelineState = 'Blocked' }
    Write-PubMenuOption -Number '10' -Label 'Run full pipeline (4 -> 9) unattended' -Note $pipelineNote -State $pipelineState
    Write-PubMenuOption -Number '11' -Label 'View recent log' -Note 'run history'
    Write-PubMenuOption -Number '12' -Label 'Open working folder' -Note 'files on disk'
    Write-PubMenuOption -Number '13' -Label 'Change conversion & upload settings' -Note 'rules, folders'

    Write-Host ''
    Write-PubMenuOption -Number '0' -Label 'Exit'
    Write-PubMenuRule
}

# ---------------------------------------------------------------------------
# Option 1 - app registration
# ---------------------------------------------------------------------------
function Show-PubAppRegistrationMenu {
    [CmdletBinding()]
    param()

    while ($true) {
        $script:Config = Get-PubConfig

        Write-Host ''
        Write-Host '  AZURE AD APP REGISTRATION' -ForegroundColor Cyan
        Write-Host ('  Current: {0}' -f (Get-PubAppRegistrationStatus -Config $script:Config))
        if (-not [string]::IsNullOrWhiteSpace([string] $script:Config['AppId'])) {
            Write-Host ('  App ID : {0}   Auth scope: {1}' -f $script:Config['AppId'], $script:Config['AuthMethod'])
        }
        Write-Host ''
        Write-Host '   1) Create a new app registration in this tenant'
        Write-Host '   2) Connect to an existing app registration (enter App ID + Tenant ID)'
        Write-Host '   3) Re-check the configured app registration and its consent'
        Write-Host '   4) Grant this app access to one specific site (Sites.Selected mode)'
        Write-Host '   5) Show the permissions this app needs and what they expose'
        Write-Host '   0) Back to the main menu'
        Write-Host ''

        switch (Read-PubMenuChoice -Valid @('0', '1', '2', '3', '4', '5')) {
            '1' { Invoke-PubCreateAppRegistration; Wait-PubKeyPress }
            '2' { Invoke-PubConnectExistingApp;     Wait-PubKeyPress }
            '3' { Invoke-PubRecheckAppRegistration; Wait-PubKeyPress }
            '4' { Invoke-PubGrantSiteAccess;        Wait-PubKeyPress }
            '5' { Show-PubPermissionTable -AuthMethod ([string] $script:Config['AuthMethod']); Wait-PubKeyPress }
            '0' { return }
        }
    }
}

function Invoke-PubCreateAppRegistration {
    [CmdletBinding()]
    param()

    $tenant = Read-Host 'Tenant ID or domain (e.g. contoso.onmicrosoft.com)'
    if ([string]::IsNullOrWhiteSpace($tenant)) {
        Write-PubLog -Level Warn -Message 'No tenant supplied - cancelled.'
        return
    }

    Write-Host ''
    Write-Host '  HOW MUCH ACCESS SHOULD THIS APP HAVE?' -ForegroundColor Cyan
    Write-Host '   1) Tenant-wide + SharePoint admin - FULLY AUTOMATIC site discovery'
    Write-Host '      Adds SharePoint Sites.FullControl.All so discovery can read the tenant'
    Write-Host '      admin site list. Finds every site with no manual export. Largest grant:'
    Write-Host '      full administrative control of every site collection.'
    Write-Host '   2) Tenant-wide (Sites.Read.All + Sites.ReadWrite.All + Files.ReadWrite.All)'
    Write-Host '      Read AND WRITE to every site. Site discovery falls back to the search'
    Write-Host '      index, which can miss sites - see the README.'
    Write-Host '   3) Selected sites only (Sites.Selected) - tightest scope'
    Write-Host '      No access until an administrator grants each site (option 4 in this menu).'
    Write-Host '   0) Cancel'
    Write-Host ''

    $authMethod = 'TenantAdmin'
    switch (Read-PubMenuChoice -Valid @('0', '1', '2', '3')) {
        '1' { $authMethod = 'TenantAdmin' }
        '2' { $authMethod = 'AllSites' }
        '3' { $authMethod = 'SitesSelected' }
        '0' { Write-PubLog -Level Info -Message 'Cancelled.'; return }
    }

    Show-PubPermissionTable -AuthMethod $authMethod

    if (-not (Confirm-PubAction -Question 'Create the app registration with these permissions?')) {
        Write-PubLog -Level Warn -Message 'Cancelled - nothing was created.'
        return
    }

    # One-step path: PnP creates the app, the certificate and the consent flow
    # together, so options 1 and 2 collapse into a single action.
    if (Test-PubPnPAvailable -Quiet) {
        Write-Host ''
        Write-Host '  SETUP STYLE' -ForegroundColor Cyan
        Write-Host '  PnP PowerShell is available on this host.'
        Write-Host '   1) One-step setup with PnP (app + certificate + consent together)'
        Write-Host '   2) Step-by-step with Graph (app now, certificate at option 2)'
        Write-Host '   0) Cancel'
        Write-Host ''

        switch (Read-PubMenuChoice -Valid @('0', '1', '2')) {
            '1' { Invoke-PubPnPRegistration -Tenant $tenant -AuthMethod $authMethod; return }
            '2' { }
            '0' { Write-PubLog -Level Info -Message 'Cancelled.'; return }
        }
    } elseif ($authMethod -eq 'TenantAdmin') {
        Write-PubLog -Level Info -Message 'PnP PowerShell is not installed - the app will still be granted the SharePoint permission, but install PnP to use it: Install-Module PnP.PowerShell -Scope CurrentUser'
    }

    if (-not (Connect-PubGraphInteractive -TenantId $tenant)) { return }

    $context = Get-PubGraphContext
    $tenantId = $tenant
    if ($context -and $context.TenantId) { $tenantId = $context.TenantId }

    $shortName = ($tenant -split '\.')[0]
    $defaultName = 'SPO-Publisher-Converter-{0}' -f $shortName
    $displayName = Read-Host ('App registration name [{0}]' -f $defaultName)
    if ([string]::IsNullOrWhiteSpace($displayName)) { $displayName = $defaultName }

    $existing = $null
    try {
        $existingApps = Invoke-PubGraph -Uri ("applications?`$filter=displayName eq '{0}'" -f $displayName) -Method GET
        if ($existingApps -and $existingApps.PSObject.Properties['value'] -and $existingApps.value) { $existing = @($existingApps.value)[0] }
    } catch { }

    if ($existing) {
        Write-PubLog -Level Warn -Message ('An app registration named "{0}" already exists (AppId {1}).' -f $displayName, $existing.appId)
        if (-not (Confirm-PubAction -Question 'Reuse that existing registration instead of creating another?')) { return }

        $application = [pscustomobject] @{
            AppId              = $existing.appId
            ObjectId           = $existing.id
            DisplayName        = $existing.displayName
            ServicePrincipalId = $null
            AuthMethod         = $authMethod
        }
    } else {
        $application = New-PubAppRegistration -DisplayName $displayName -AuthMethod $authMethod
        if (-not $application) { return }
    }

    $script:Config['TenantId']       = $tenantId
    $script:Config['TenantDomain']   = $tenant
    $script:Config['AppId']          = $application.AppId
    $script:Config['AppObjectId']    = $application.ObjectId
    $script:Config['AppDisplayName'] = $application.DisplayName
    $script:Config['AuthMethod']     = $authMethod
    Save-PubConfig -Config $script:Config | Out-Null

    Write-PubLog -Level Info -Message 'Granting admin consent...'
    Grant-PubAdminConsent -AppId $application.AppId -ServicePrincipalId $application.ServicePrincipalId -AuthMethod $authMethod -TenantId $tenantId | Out-Null

    Write-Host ''
    Write-PubLog -Level Success -Message 'App registration step complete. Next: option 2 to generate the authentication certificate.'
}

function Invoke-PubPnPRegistration {
    <#
    .SYNOPSIS
        One-step app registration through PnP, then records it in config.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Tenant,

        [ValidateSet('AllSites', 'SitesSelected', 'TenantAdmin')]
        [string] $AuthMethod = 'TenantAdmin'
    )

    $shortName   = ($Tenant -split '\.')[0]
    $defaultName = 'SPO-Publisher-Converter-{0}' -f $shortName

    $displayName = Read-Host ('App registration name [{0}]' -f $defaultName)
    if ([string]::IsNullOrWhiteSpace($displayName)) { $displayName = $defaultName }

    $years  = 2
    $answer = Read-Host ('Certificate validity in years [{0}]' -f $years)
    if (-not [string]::IsNullOrWhiteSpace($answer)) {
        $parsed = 0
        if ([int]::TryParse($answer, [ref] $parsed) -and $parsed -ge 1 -and $parsed -le 5) { $years = $parsed }
        else { Write-PubLog -Level Warn -Message ('"{0}" is not 1-5 years - using {1}.' -f $answer, $years) }
    }

    $registration = Register-PubPnPApplication -DisplayName $displayName -Tenant $Tenant -AuthMethod $AuthMethod -ValidYears $years
    if (-not $registration) {
        Write-PubLog -Level Error -Message 'PnP registration did not complete. Try the step-by-step Graph path instead.'
        return
    }

    $script:Config['TenantId']              = $Tenant
    $script:Config['TenantDomain']          = $Tenant
    $script:Config['AppId']                 = $registration.AppId
    $script:Config['AppDisplayName']        = $displayName
    $script:Config['AuthMethod']            = $AuthMethod
    $script:Config['CertificateThumbprint'] = [string] $registration.Thumbprint
    $script:Config['CertificatePublicPath'] = [string] $registration.CerPath
    $script:Config['CertificatePfxPath']    = [string] $registration.PfxPath

    if ($registration.NotAfter) {
        $script:Config['CertificateExpiry'] = $registration.NotAfter.ToString('yyyy-MM-dd')
    }

    Save-PubConfig -Config $script:Config | Out-Null

    # The tenant id recorded above is the domain PnP was given; resolve the
    # real GUID now that we can sign in, so later runs connect cleanly.
    Write-Host ''
    Write-PubLog -Level Info -Message 'Consent and certificate publication can take a minute. Testing the connection...'
    Start-Sleep -Seconds 10

    if (Connect-PubGraphApp -Config $script:Config -Force) {
        $context = Get-PubGraphContext
        if ($context -and $context.TenantId) {
            $script:Config['TenantId'] = $context.TenantId
            Save-PubConfig -Config $script:Config | Out-Null
        }
        Write-PubLog -Level Success -Message 'Setup complete - options 1 and 2 are both done. Run option 3 to confirm SharePoint access.'
    } else {
        Write-PubLog -Level Warn -Message 'The app is registered but app-only sign-in did not work yet. Wait a minute and run option 3.'
    }

    if ($AuthMethod -eq 'TenantAdmin') {
        Write-Host ''
        Write-PubLog -Level Info -Message 'Checking the SharePoint tenant admin site list...'
        if (Connect-PubPnPAdmin -Config $script:Config -Force) {
            Write-PubLog -Level Success -Message 'Tenant admin enumeration is working - discovery will find every site automatically.'
        } else {
            Write-PubLog -Level Warn -Message 'Tenant admin enumeration is not working yet. SharePoint permission changes can take several minutes; re-test with option 3.'
        }
    }
}

function Invoke-PubConnectExistingApp {
    [CmdletBinding()]
    param()

    $appId    = (Read-Host 'App (client) ID').Trim()
    $tenantId = (Read-Host 'Tenant ID or domain').Trim()

    if ([string]::IsNullOrWhiteSpace($appId) -or [string]::IsNullOrWhiteSpace($tenantId)) {
        Write-PubLog -Level Warn -Message 'App ID and Tenant ID are both required - cancelled.'
        return
    }

    if (-not (Connect-PubGraphInteractive -TenantId $tenantId)) { return }

    $application = Get-PubApplication -AppId $appId
    if (-not $application) {
        Write-PubLog -Level Error -Message ('No app registration with App ID {0} was found in this tenant.' -f $appId)
        return
    }

    $context = Get-PubGraphContext
    if ($context -and $context.TenantId) { $tenantId = $context.TenantId }

    $script:Config['TenantId']       = $tenantId
    $script:Config['AppId']          = $application.appId
    $script:Config['AppObjectId']    = $application.id
    $script:Config['AppDisplayName'] = $application.displayName
    Save-PubConfig -Config $script:Config | Out-Null

    Write-PubLog -Level Success -Message ('Connected to "{0}".' -f $application.displayName)

    $thumbprint = (Read-Host 'Existing certificate thumbprint (leave blank to generate one at option 2)').Trim()
    if (-not [string]::IsNullOrWhiteSpace($thumbprint)) {
        $script:Config['CertificateThumbprint'] = $thumbprint.Replace(' ', '').ToUpper()

        $certificate = Get-PubCertificate -Config $script:Config
        if ($certificate) {
            $script:Config['CertificateExpiry'] = $certificate.NotAfter.ToString('yyyy-MM-dd')
            Write-PubLog -Level Success -Message ('Certificate found locally, expires {0}.' -f $certificate.NotAfter.ToString('yyyy-MM-dd'))
        } else {
            Write-PubLog -Level Warn -Message 'That thumbprint is not in this machine''s certificate store - app-only sign-in will fail until it is.'
        }
        Save-PubConfig -Config $script:Config | Out-Null
    }

    Test-PubAdminConsent -AppId $application.appId -AuthMethod ([string] $script:Config['AuthMethod']) | Out-Null
}

function Invoke-PubRecheckAppRegistration {
    [CmdletBinding()]
    param()

    $script:Config = Get-PubConfig

    if ([string]::IsNullOrWhiteSpace([string] $script:Config['AppId'])) {
        Write-PubLog -Level Warn -Message 'No app registration configured yet - use option 1 or 2 in this sub-menu.'
        return
    }

    if (-not (Connect-PubGraphInteractive -TenantId ([string] $script:Config['TenantId']))) { return }

    $application = Get-PubApplication -AppId ([string] $script:Config['AppId'])
    if (-not $application) {
        Write-PubLog -Level Error -Message 'The configured app registration no longer exists in this tenant. Create a new one with option 1.'
        return
    }

    Write-PubLog -Level Success -Message ('App registration "{0}" exists.' -f $application.displayName)

    $script:Config['AppObjectId']    = $application.id
    $script:Config['AppDisplayName'] = $application.displayName
    Save-PubConfig -Config $script:Config | Out-Null

    Test-PubAdminConsent -AppId $application.appId -AuthMethod ([string] $script:Config['AuthMethod']) | Out-Null

    $expiry = Test-PubCertificateExpiry -Config $script:Config
    Write-PubLog -Level Info -Message $expiry.Message
}

function Invoke-PubGrantSiteAccess {
    [CmdletBinding()]
    param()

    $script:Config = Get-PubConfig

    if ([string]::IsNullOrWhiteSpace([string] $script:Config['AppId'])) {
        Write-PubLog -Level Warn -Message 'Configure the app registration first.'
        return
    }

    Write-Host ''
    Write-Host '  Sites.Selected grants are per site. Discovery needs read; upload needs write.' -ForegroundColor Gray
    $siteUrl = (Read-Host 'Site URL (e.g. https://contoso.sharepoint.com/sites/Marketing)').Trim()
    if ([string]::IsNullOrWhiteSpace($siteUrl)) { return }

    Write-Host ''
    Write-Host '  ACCESS LEVEL FOR THIS SITE' -ForegroundColor Cyan
    Write-Host '   1) Read and write (needed for the full pipeline)'
    Write-Host '   2) Read only (discovery and download only)'
    Write-Host '   0) Cancel'
    Write-Host ''

    $role = 'write'
    switch (Read-PubMenuChoice -Valid @('0', '1', '2')) {
        '1' { $role = 'write' }
        '2' { $role = 'read' }
        '0' { return }
    }

    if (-not (Confirm-PubAction -Question ('Grant "{0}" on {1} to this app?' -f $role, $siteUrl))) { return }

    # The grant itself is a directory write, so it needs the interactive
    # administrator session rather than the app's own certificate.
    if (-not (Connect-PubGraphInteractive -TenantId ([string] $script:Config['TenantId']) -Scopes @('Sites.FullControl.All'))) { return }

    Grant-PubSiteSelectedPermission -SiteUrl $siteUrl -Role $role -Config $script:Config | Out-Null
}

# ---------------------------------------------------------------------------
# Option 2 - certificate
# ---------------------------------------------------------------------------
function Invoke-PubCertificateSetup {
    [CmdletBinding()]
    param()

    $script:Config = Get-PubConfig

    if ([string]::IsNullOrWhiteSpace([string] $script:Config['AppId'])) {
        Write-PubLog -Level Warn -Message 'Configure the app registration first (option 1).'
        return
    }

    $expiry = Test-PubCertificateExpiry -Config $script:Config
    if ($expiry.State -eq 'Valid') {
        Write-Host ''
        Write-PubLog -Level Info -Message $expiry.Message
        if (-not (Confirm-PubAction -Question 'A valid certificate is already configured. Generate a replacement anyway?')) { return }
    }

    $subject = 'SPO-PubConverter-{0}' -f ([string] $script:Config['AppId']).Substring(0, 8)
    $years   = 2

    $answer = Read-Host ('Certificate validity in years [{0}]' -f $years)
    if (-not [string]::IsNullOrWhiteSpace($answer)) {
        $parsed = 0
        if ([int]::TryParse($answer, [ref] $parsed) -and $parsed -ge 1 -and $parsed -le 5) { $years = $parsed }
        else { Write-PubLog -Level Warn -Message ('"{0}" is not 1-5 years - using {1}.' -f $answer, $years) }
    }

    $certificate = New-PubAuthCertificate -SubjectName $subject -ValidityYears $years
    if (-not $certificate) { return }

    if (-not (Connect-PubGraphInteractive -TenantId ([string] $script:Config['TenantId']))) {
        Write-PubLog -Level Warn -Message 'Certificate created locally but not uploaded - re-run this option when you can sign in.'
        return
    }

    $objectId = [string] $script:Config['AppObjectId']
    if ([string]::IsNullOrWhiteSpace($objectId)) {
        $application = Get-PubApplication -AppId ([string] $script:Config['AppId'])
        if (-not $application) {
            Write-PubLog -Level Error -Message 'Could not resolve the app registration to upload the certificate to.'
            return
        }
        $objectId = $application.id
        $script:Config['AppObjectId'] = $objectId
    }

    if (-not (Add-PubCertificateToApp -ApplicationObjectId $objectId -Base64Certificate $certificate.Base64)) { return }

    $script:Config['CertificateThumbprint'] = $certificate.Thumbprint
    $script:Config['CertificateExpiry']     = $certificate.NotAfter.ToString('yyyy-MM-dd')
    $script:Config['CertificatePublicPath'] = $certificate.CerPath
    $script:Config['CertificatePfxPath']    = $certificate.PfxPath
    Save-PubConfig -Config $script:Config | Out-Null

    Write-Host ''
    Write-PubLog -Level Success -Message 'Certificate ready. Azure AD can take a minute to publish it - then run option 3 to test the connection.'
}

# ---------------------------------------------------------------------------
# Option 4 - discovery
# ---------------------------------------------------------------------------
function Invoke-PubScanMenu {
    [CmdletBinding()]
    param(
        [switch] $Unattended
    )

    $script:Config = Get-PubConfig

    $scopePath    = ''
    $siteUrls     = @()
    $includeOneDrive = $false

    if (-not $Unattended) {
        Write-Host ''
        Write-Host '  SCAN SCOPE' -ForegroundColor Cyan
        Write-Host '   1) Every site the app registration can enumerate (full tenant crawl)'
        Write-Host '   2) Specific sites - type the URLs now (comma separated)'
        Write-Host '   3) Specific sites - read the URLs from a text or CSV file'
        Write-Host '      (including the SharePoint admin centre Active sites export, unedited -'
        Write-Host '       the one list guaranteed to contain every site collection)'
        Write-Host '   0) Cancel'
        Write-Host ''

        switch (Read-PubMenuChoice -Valid @('0', '1', '2', '3')) {
            '1' { }
            '2' {
                $entered = Read-Host 'Site URLs (comma separated)'
                if ([string]::IsNullOrWhiteSpace($entered)) { Write-PubLog -Level Warn -Message 'No URLs entered - cancelled.'; return }
                $siteUrls = @($entered -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            }
            '3' {
                $default = [string] $script:Config['ScopeSiteListPath']
                $prompt  = 'Path to the site list file'
                if (-not [string]::IsNullOrWhiteSpace($default)) { $prompt = ('{0} [{1}]' -f $prompt, $default) }

                $entered = (Read-Host $prompt).Trim()
                if ([string]::IsNullOrWhiteSpace($entered)) { $entered = $default }
                if ([string]::IsNullOrWhiteSpace($entered)) { Write-PubLog -Level Warn -Message 'No path entered - cancelled.'; return }

                $scopePath = $entered
                $script:Config['ScopeSiteListPath'] = $scopePath
                Save-PubConfig -Config $script:Config | Out-Null
            }
            '0' { return }
        }

        Write-Host ''
        Write-Host '  ONEDRIVE (personal) SITES' -ForegroundColor Cyan
        Write-Host '   1) SharePoint sites only (default)'
        Write-Host '   2) Also crawl OneDrive for Business sites'
        Write-Host '      Slower, and it reads every user''s personal files - only where the'
        Write-Host '      Publisher retirement has to cover OneDrive too.'
        Write-Host '   0) Cancel'
        Write-Host ''

        switch (Read-PubMenuChoice -Valid @('0', '1', '2')) {
            '1' { $includeOneDrive = $false }
            '2' { $includeOneDrive = $true }
            '0' { return }
        }
    }

    $rows = Invoke-PubDiscovery -ScopePath $scopePath -SiteUrl $siteUrls -IncludePersonalSites:$includeOneDrive -Config $script:Config
    if (-not $rows -or @($rows).Count -eq 0) {
        Write-PubLog -Level Warn -Message 'No Publisher files found in scope.'
        $script:Inventory = @()
        return
    }

    $script:Inventory = @($rows)
    $script:Selection = @()

    $path = Export-PubInventory -Rows $script:Inventory -Config $script:Config
    if ($path) { $script:InventoryPath = $path }
}

function Invoke-PubExportInventory {
    [CmdletBinding()]
    param()

    if (-not $script:Inventory -or @($script:Inventory).Count -eq 0) {
        Write-PubLog -Level Warn -Message 'Nothing to export - run a scan (option 4) or load a CSV (option 6) first.'
        return
    }

    $path = Export-PubInventory -Rows $script:Inventory -Config $script:Config
    if ($path) { $script:InventoryPath = $path }
}

# ---------------------------------------------------------------------------
# Option 6 - load and select
# ---------------------------------------------------------------------------
function Invoke-PubLoadInventory {
    [CmdletBinding()]
    param()

    $script:Config = Get-PubConfig

    $inventoryFolder = Get-PubWorkingFolder -SubFolder 'inventory' -Config $script:Config
    $recent = Get-ChildItem -LiteralPath $inventoryFolder -Filter 'PublisherFileInventory_*.csv' -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 5

    Write-Host ''
    Write-Host '  LOAD AN INVENTORY CSV' -ForegroundColor Cyan

    $options = @('0')
    $index   = 0

    foreach ($file in $recent) {
        $index++
        $options += "$index"
        Write-Host ('   {0}) {1}   ({2}, {3:yyyy-MM-dd HH:mm})' -f $index, $file.Name, ('{0:N0} KB' -f ($file.Length / 1KB)), $file.LastWriteTime)
    }

    $browseOption = $index + 1
    $options += "$browseOption"
    Write-Host ('   {0}) Enter the full path to another CSV' -f $browseOption)
    Write-Host '   0) Cancel'
    Write-Host ''

    $choice = Read-PubMenuChoice -Valid $options
    if ($choice -eq '0') { return }

    $path = ''
    if ([int] $choice -eq $browseOption) {
        $path = (Read-Host 'Full path to the CSV').Trim('"', ' ')
    } else {
        $path = $recent[[int] $choice - 1].FullName
    }

    if ([string]::IsNullOrWhiteSpace($path)) { return }

    $rows = Import-PubInventory -Path $path
    if (-not $rows -or @($rows).Count -eq 0) { return }

    $script:Inventory     = @($rows)
    $script:InventoryPath = $path
    $script:Selection     = @()

    $script:Config['LastInventoryCsv'] = $path
    Save-PubConfig -Config $script:Config | Out-Null

    Show-PubInventorySummary
    Invoke-PubSelectionMenu
}

function Show-PubInventorySummary {
    [CmdletBinding()]
    param()

    $statistics = Get-PubInventoryStatistic -Rows $script:Inventory

    Write-Host ''
    Write-Host ('  {0} row(s) loaded' -f $statistics.Total) -ForegroundColor Cyan
    foreach ($key in @('Pending', 'Downloaded', 'Converted', 'Uploaded', 'Skipped', 'Failed')) {
        Write-Host ('    {0,-11}: {1}' -f $key, $statistics[$key])
    }
}

function Invoke-PubSelectionMenu {
    [CmdletBinding()]
    param()

    while ($true) {
        $selectionCount = @($script:Selection).Count
        $label = ('every row in the CSV ({0})' -f @($script:Inventory).Count)
        if ($selectionCount -gt 0) { $label = ('{0} of {1} row(s)' -f $selectionCount, @($script:Inventory).Count) }

        Write-Host ''
        Write-Host '  SELECT FILES TO PROCESS' -ForegroundColor Cyan
        Write-Host ('  Options 7, 8 and 9 will act on: {0}' -f $label) -ForegroundColor Yellow
        Write-Host '  Filters stack, so you can narrow twice (site, then status).'
        Write-Host ''
        Write-Host '   1) Process every row in this CSV (clear any filters)'
        Write-Host '   2) Narrow down by site URL'
        Write-Host '   3) Narrow down by library or folder path'
        Write-Host '   4) Narrow down by file name'
        Write-Host '   5) Narrow down by status (e.g. retry only the failures)'
        Write-Host '   6) Show me the rows currently selected'
        Write-Host '   0) Done - use this selection'
        Write-Host ''

        switch (Read-PubMenuChoice -Valid @('0', '1', '2', '3', '4', '5', '6')) {
            '1' {
                $script:Selection = @()
                Write-PubLog -Level Info -Message 'Selection cleared - every row will be processed.'
            }
            '2' {
                $pattern = Read-Host 'Site URL contains (wildcards allowed, e.g. *marketing*)'
                if ($pattern) { Set-PubSelection -SiteFilter ('*{0}*' -f $pattern.Trim('*')) }
            }
            '3' {
                $pattern = Read-Host 'Library/folder contains (wildcards allowed)'
                if ($pattern) { Set-PubSelection -FolderFilter ('*{0}*' -f $pattern.Trim('*')) }
            }
            '4' {
                $pattern = Read-Host 'File name contains (wildcards allowed)'
                if ($pattern) { Set-PubSelection -NameFilter ('*{0}*' -f $pattern.Trim('*')) }
            }
            '5' {
                $source = $script:Selection
                if (-not $source -or @($source).Count -eq 0) { $source = $script:Inventory }
                $counts = Get-PubInventoryStatistic -Rows $source

                Write-Host ''
                Write-Host '  FILTER BY STATUS' -ForegroundColor Cyan
                Write-Host '  Keep only the rows at one stage of the pipeline.'
                Write-Host ''
                Write-Host ('   1) Pending     - found in SharePoint, not downloaded yet   ({0})' -f $counts.Pending)
                Write-Host ('   2) Downloaded  - on this machine, not converted yet        ({0})' -f $counts.Downloaded)
                Write-Host ('   3) Converted   - PDF made, not uploaded yet                ({0})' -f $counts.Converted)
                Write-Host ('   4) Uploaded    - finished, PDF is back in SharePoint       ({0})' -f $counts.Uploaded)
                Write-Host ('   5) Failed      - something went wrong, retry these         ({0})' -f $counts.Failed)
                Write-Host ('   6) Skipped     - deliberately left alone                   ({0})' -f $counts.Skipped)
                Write-Host '   0) Cancel'
                Write-Host ''

                $statusMap = @{ '1' = 'Pending'; '2' = 'Downloaded'; '3' = 'Converted'; '4' = 'Uploaded'; '5' = 'Failed'; '6' = 'Skipped' }
                $answer    = Read-PubMenuChoice -Valid @('0', '1', '2', '3', '4', '5', '6')
                if ($answer -ne '0') { Set-PubSelection -Status @($statusMap[$answer]) }
            }
            '6' { Show-PubSelectionPreview }
            '0' { return }
        }
    }
}

function Set-PubSelection {
    <#
    .SYNOPSIS
        Applies one filter to the current selection (filters stack).
    #>
    [CmdletBinding()]
    param(
        [string]   $SiteFilter,
        [string]   $FolderFilter,
        [string]   $NameFilter,
        [string[]] $Status
    )

    $source = $script:Selection
    if (-not $source -or @($source).Count -eq 0) { $source = $script:Inventory }

    $filtered = Select-PubInventoryRow -Rows $source -SiteFilter $SiteFilter -FolderFilter $FolderFilter -NameFilter $NameFilter -Status $Status

    if (@($filtered).Count -eq 0) {
        Write-PubLog -Level Warn -Message 'That filter matched no rows - the previous selection is unchanged.'
        return
    }

    $script:Selection = @($filtered)
    Write-PubLog -Level Success -Message ('{0} row(s) selected.' -f @($script:Selection).Count)
}

function Show-PubSelectionPreview {
    [CmdletBinding()]
    param(
        [int] $First = 25
    )

    $rows = $script:Selection
    if (-not $rows -or @($rows).Count -eq 0) { $rows = $script:Inventory }
    if (-not $rows -or @($rows).Count -eq 0) {
        Write-PubLog -Level Warn -Message 'No rows loaded.'
        return
    }

    Write-Host ''
    @($rows) | Select-Object -First $First FileName, Status, LibraryName, FolderPath, SiteUrl |
        Format-Table -AutoSize | Out-String | Write-Host

    if (@($rows).Count -gt $First) {
        Write-Host ('  ... and {0} more row(s).' -f (@($rows).Count - $First)) -ForegroundColor DarkGray
    }
}

function Get-PubRowsToProcess {
    <#
    .SYNOPSIS
        The rows the action phases operate on: the selection, or everything.
    #>
    [CmdletBinding()]
    param()

    if ($script:Selection -and @($script:Selection).Count -gt 0) { return @($script:Selection) }
    return @($script:Inventory)
}

function Save-PubInventoryInPlace {
    <#
    .SYNOPSIS
        Writes status updates back to the loaded CSV so a run can resume.
    #>
    [CmdletBinding()]
    param()

    if (-not $script:Inventory -or @($script:Inventory).Count -eq 0) { return }

    if ([string]::IsNullOrWhiteSpace($script:InventoryPath)) {
        $path = Export-PubInventory -Rows $script:Inventory -Config $script:Config
        if ($path) { $script:InventoryPath = $path }
        return
    }

    Export-PubInventory -Rows $script:Inventory -Path $script:InventoryPath -Config $script:Config | Out-Null
}

# ---------------------------------------------------------------------------
# Options 7, 8, 9 - download, convert, upload
# ---------------------------------------------------------------------------
function Invoke-PubDownloadStep {
    [CmdletBinding()]
    param()

    $rows = Get-PubRowsToProcess
    if (@($rows).Count -eq 0) {
        Write-PubLog -Level Warn -Message 'No inventory loaded - run option 4 or 6 first.'
        return
    }

    Invoke-PubDownload -Rows $rows -Config $script:Config | Out-Null
    Save-PubInventoryInPlace
}

function Invoke-PubConvertStep {
    [CmdletBinding()]
    param()

    $rows = Get-PubRowsToProcess
    if (@($rows).Count -eq 0) {
        Write-PubLog -Level Warn -Message 'No inventory loaded - run option 4 or 6 first.'
        return
    }

    $ready = @($rows | Where-Object { @('Downloaded', 'Converted', 'Failed') -contains [string] $_.Status })
    if (@($ready).Count -eq 0) {
        Write-PubLog -Level Warn -Message 'None of the selected rows have been downloaded yet - run option 7 first.'
        return
    }

    Invoke-PubConvert -Rows $ready -Config $script:Config | Out-Null
    Save-PubInventoryInPlace
}

function Invoke-PubUploadStep {
    [CmdletBinding()]
    param(
        [switch] $PreConfirmed
    )

    $rows = Get-PubRowsToProcess
    if (@($rows).Count -eq 0) {
        Write-PubLog -Level Warn -Message 'No inventory loaded - run option 4 or 6 first.'
        return
    }

    $removeSource = [bool] $script:Config['RemoveSourceAfterUpload']

    Invoke-PubUpload -Rows $rows -Config $script:Config -Confirmed:$PreConfirmed -RemoveSource:$removeSource | Out-Null
    Save-PubInventoryInPlace
}

# ---------------------------------------------------------------------------
# Option 10 - full pipeline
# ---------------------------------------------------------------------------
function Invoke-PubFullPipeline {
    [CmdletBinding()]
    param()

    Write-Host ''
    Write-Host '  RUN FULL PIPELINE (4 -> 9)' -ForegroundColor Cyan
    Write-Host '  Scan, export CSV, download, convert to PDF, then upload the PDFs to' -ForegroundColor Gray
    Write-Host '  their original SharePoint folders. The upload is a WRITE to SharePoint,' -ForegroundColor Gray
    Write-Host '  so it is confirmed once here rather than file by file.' -ForegroundColor Gray
    Write-Host ''
    Write-Host ('  Existing local PDF : {0}' -f $script:Config['ExistingPdfAction']) -ForegroundColor Gray
    Write-Host ('  Name collisions    : {0}' -f $script:Config['UploadConflictAction']) -ForegroundColor Gray

    if ([bool] $script:Config['RemoveSourceAfterUpload']) {
        Write-Host '  Source .pub files  : WILL BE DELETED after a successful upload' -ForegroundColor Red
    } else {
        Write-Host '  Source .pub files  : left in place' -ForegroundColor Gray
    }
    Write-Host ''

    if (-not (Test-PubPublisherAvailable -Quiet)) {
        Write-PubLog -Level Error -Message 'Microsoft Publisher is not available on this host - the pipeline would stop at the conversion step.'
        Write-PubLog -Level Info  -Message 'Run options 4-7 here, then run options 6-9 on a host with Publisher installed.'
        if (-not (Confirm-PubAction -Question 'Run the scan and download steps anyway?')) { return }
    }

    $typed = Read-Host 'Type RUN to start the unattended pipeline'
    if ($typed -cne 'RUN') {
        Write-PubLog -Level Warn -Message 'Cancelled - nothing has run.'
        return
    }

    $started = Get-Date
    Write-PubLog -Level Info -Message '=== Full pipeline started ==='

    Invoke-PubScanMenu -Unattended
    if (-not $script:Inventory -or @($script:Inventory).Count -eq 0) {
        Write-PubLog -Level Warn -Message 'Pipeline stopped: no Publisher files found.'
        return
    }

    $script:Selection = @()

    Invoke-PubDownloadStep
    Invoke-PubConvertStep
    Invoke-PubUploadStep -PreConfirmed

    $statistics = Get-PubInventoryStatistic -Rows $script:Inventory
    $elapsed    = (Get-Date) - $started

    Write-PubPhaseSummary -Phase 'Full pipeline' `
                          -Attempted $statistics.Total `
                          -Succeeded $statistics.Uploaded `
                          -Failed $statistics.Failed `
                          -Skipped $statistics.Skipped `
                          -ExtraLines @(
                              ('Downloaded : {0}' -f $statistics.Downloaded)
                              ('Converted  : {0}' -f $statistics.Converted)
                              ('CSV        : {0}' -f $script:InventoryPath)
                              ('Elapsed    : {0:hh\:mm\:ss}' -f $elapsed)
                          )
}

# ---------------------------------------------------------------------------
# Options 12, 13 - working folder and settings
# ---------------------------------------------------------------------------
function Open-PubWorkingFolder {
    [CmdletBinding()]
    param()

    $folder = Get-PubWorkingFolder -Config $script:Config
    Write-Host ''
    Write-Host ('  Working folder: {0}' -f $folder) -ForegroundColor Cyan
    Write-Host ('    originals : {0}' -f (Join-Path $folder 'originals')) -ForegroundColor Gray
    Write-Host ('    converted : {0}' -f (Join-Path $folder 'converted')) -ForegroundColor Gray
    Write-Host ('    inventory : {0}' -f (Join-Path $folder 'inventory')) -ForegroundColor Gray

    if (Test-PubIsWindows) {
        try { Start-Process -FilePath 'explorer.exe' -ArgumentList $folder | Out-Null } catch {
            Write-PubLog -Level Warn -Message ('Could not open Explorer: {0}' -f $_.Exception.Message)
        }
    }
}

function Invoke-PubSettingsMenu {
    [CmdletBinding()]
    param()

    while ($true) {
        $script:Config = Get-PubConfig

        $removeLabel = 'No - leave the .pub in place'
        if ([bool] $script:Config['RemoveSourceAfterUpload']) { $removeLabel = 'YES - delete the .pub after a successful upload' }

        Write-Host ''
        Write-Host '  CONVERSION & UPLOAD SETTINGS' -ForegroundColor Cyan
        Write-Host ''
        Write-Host ('   1) When a converted PDF already exists locally : {0}' -f $script:Config['ExistingPdfAction'])
        Write-Host ('   2) When a PDF of that name exists in SharePoint: {0}' -f $script:Config['UploadConflictAction'])
        Write-Host ('   3) Remove the original .pub after upload       : {0}' -f $removeLabel)
        Write-Host ('   4) Working folder                              : {0}' -f $script:Config['DefaultWorkingFolder'])
        Write-Host ('   5) Site enumeration method                     : {0}' -f $script:Config['EnumerationMethod'])
        Write-Host '   0) Back to the main menu'
        Write-Host ''

        switch (Read-PubMenuChoice -Valid @('0', '1', '2', '3', '4', '5')) {
            '1' { Set-PubThreeWaySetting -Name 'ExistingPdfAction'   -Title 'Existing local PDF' }
            '2' { Set-PubThreeWaySetting -Name 'UploadConflictAction' -Title 'SharePoint name collision' }
            '3' {
                Write-Host ''
                Write-Host '  Deleting the source .pub is irreversible from this tool (the file goes to the' -ForegroundColor Yellow
                Write-Host '  site recycle bin). It stays off unless you turn it on here, and each upload' -ForegroundColor Yellow
                Write-Host '  run still asks for a separate typed confirmation.' -ForegroundColor Yellow
                Write-Host ''
                Write-Host '  WHAT HAPPENS TO THE ORIGINAL .pub' -ForegroundColor Cyan
                Write-Host '   1) No - upload the PDF alongside the .pub (default)'
                Write-Host '   2) Yes - delete the .pub after its PDF uploads successfully'
                Write-Host '   0) Cancel'
                Write-Host ''

                switch (Read-PubMenuChoice -Valid @('0', '1', '2')) {
                    '1' { Set-PubConfigValue -Name 'RemoveSourceAfterUpload' -Value $false | Out-Null; Write-PubLog -Level Success -Message 'Originals will be kept.' }
                    '2' {
                        if (Confirm-PubAction -Question 'Really allow this tool to delete source .pub files?') {
                            Set-PubConfigValue -Name 'RemoveSourceAfterUpload' -Value $true | Out-Null
                            Write-PubLog -Level Warn -Message 'Source deletion is now ARMED - each upload run will still ask you to type DELETE.'
                        }
                    }
                    '0' { }
                }
            }
            '4' {
                $entered = (Read-Host 'New working folder path (blank to cancel)').Trim('"', ' ')
                if (-not [string]::IsNullOrWhiteSpace($entered)) {
                    Set-PubConfigValue -Name 'DefaultWorkingFolder' -Value $entered | Out-Null
                    $script:Config = Get-PubConfig
                    Get-PubWorkingFolder -Config $script:Config | Out-Null
                    Write-PubLog -Level Success -Message ('Working folder set to {0}' -f $entered)
                }
            }
            '5' {
                Write-Host ''
                Write-Host '  HOW SHOULD DISCOVERY FIND SITES?' -ForegroundColor Cyan
                Write-Host '   1) Auto      - SharePoint tenant admin list when available, else Graph (default)'
                Write-Host '   2) PnP       - insist on the tenant admin list; report loudly if it is unavailable'
                Write-Host '   3) Graph     - never use PnP, even when it is installed'
                Write-Host '   0) Cancel'
                Write-Host ''

                $map    = @{ '1' = 'Auto'; '2' = 'PnP'; '3' = 'Graph' }
                $answer = Read-PubMenuChoice -Valid @('0', '1', '2', '3')
                if ($answer -ne '0') {
                    Set-PubConfigValue -Name 'EnumerationMethod' -Value $map[$answer] | Out-Null
                    Write-PubLog -Level Success -Message ('Enumeration method set to {0}.' -f $map[$answer])
                }
            }
            '0' { return }
        }
    }
}

function Set-PubThreeWaySetting {
    <#
    .SYNOPSIS
        Skip / Overwrite / Version picker shared by both collision settings.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Name,
        [Parameter(Mandatory)] [string] $Title
    )

    Write-Host ''
    Write-Host ('  {0}' -f $Title.ToUpper()) -ForegroundColor Cyan
    Write-Host '   1) Skip      - leave the existing file alone and move on'
    Write-Host '   2) Overwrite - replace the existing file'
    Write-Host '   3) Version   - keep both, writing Name (2).pdf / Name 1.pdf'
    Write-Host '   0) Cancel'
    Write-Host ''

    $map = @{ '1' = 'Skip'; '2' = 'Overwrite'; '3' = 'Version' }
    $answer = Read-PubMenuChoice -Valid @('0', '1', '2', '3')
    if ($answer -eq '0') { return }

    Set-PubConfigValue -Name $Name -Value $map[$answer] | Out-Null
    Write-PubLog -Level Success -Message ('{0} set to {1}.' -f $Title, $map[$answer])
}

# ---------------------------------------------------------------------------
# Startup checks and main loop
# ---------------------------------------------------------------------------
function Invoke-PubConnectionTest {
    <#
    .SYNOPSIS
        Menu option 3 - tests Graph, SharePoint, and the enumeration route.
    #>
    [CmdletBinding()]
    param()

    Test-PubGraphAccess -Config $script:Config | Out-Null

    Write-Host ''
    Write-Host '  SITE ENUMERATION' -ForegroundColor Cyan

    if (Test-PubPnPEnumerationEnabled -Config $script:Config) {
        $script:VerifiedRoute = 'tenant admin list - finds every site (tested OK)'
        Write-PubLog -Level Success -Message 'Tenant admin route available - discovery will find every site collection, including ones the search index misses.'
    } else {
        $script:VerifiedRoute = 'Graph search - may miss some sites (tested)'
        Write-PubLog -Level Warn -Message 'Tenant admin route unavailable - discovery will fall back to Graph enumeration, which can miss unindexed sites, very new sites and Teams private-channel sites.'
        if (-not (Test-PubPnPAvailable -Quiet)) {
            Write-PubLog -Level Info -Message 'Install PnP PowerShell to enable it: Install-Module PnP.PowerShell -Scope CurrentUser'
        } elseif ([string] $script:Config['AuthMethod'] -ne 'TenantAdmin') {
            Write-PubLog -Level Info -Message ('This app is registered with the "{0}" scope. The tenant admin route needs the TenantAdmin scope (setup option 1).' -f $script:Config['AuthMethod'])
        }
        Write-PubLog -Level Info -Message 'Either way you can scope a complete scan manually: SharePoint admin centre -> Active sites -> Export to CSV, then menu option 4.'
    }
}

function Invoke-PubStartupCheck {
    [CmdletBinding()]
    param()

    Initialize-PubLogging -Name 'Session' | Out-Null

    Write-PubLog -Level Info -Message ('SharePoint Publisher File Converter {0} starting.' -f $script:Version)

    if ($PSVersionTable.PSVersion.Major -lt 5) {
        Write-PubLog -Level Error -Message 'PowerShell 5.1 or later is required.'
    }

    # Certificate expiry is checked at every startup, not just during setup.
    if (Test-PubConfigComplete -Config $script:Config) {
        $expiry = Test-PubCertificateExpiry -Config $script:Config
        if ($expiry.State -eq 'Expiring' -or $expiry.State -eq 'Expired') {
            Write-Host ''
            Write-Host ('  {0}' -f $expiry.Message) -ForegroundColor Yellow
        }
    }

    if (-not (Test-PubPublisherAvailable -Quiet)) {
        Write-PubLog -Level Warn -Message 'Microsoft Publisher was not detected on this host - options 8 and 10 will not be able to convert here.'
    }

    # Offer to pick up where the last run left off.
    $last = [string] $script:Config['LastInventoryCsv']
    if (-not [string]::IsNullOrWhiteSpace($last) -and (Test-Path -LiteralPath $last)) {
        Write-Host ''
        if (Confirm-PubAction -Question ('Resume from the last inventory, {0}?' -f (Split-Path -Leaf $last))) {
            $rows = Import-PubInventory -Path $last
            if ($rows -and @($rows).Count -gt 0) {
                $script:Inventory     = @($rows)
                $script:InventoryPath = $last
                Show-PubInventorySummary
            }
        }
    }
}

function Start-PubMenu {
    [CmdletBinding()]
    param()

    Invoke-PubStartupCheck
    Wait-PubKeyPress -Message 'Press any key to open the menu...'

    while ($true) {
        $script:Config = Get-PubConfig
        Show-PubMainMenu

        $choice = Read-PubMenuChoice -Valid @('0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13')

        try {
            switch ($choice) {
                '1'  { Show-PubAppRegistrationMenu }
                '2'  { Invoke-PubCertificateSetup;  Wait-PubKeyPress }
                '3'  { Invoke-PubConnectionTest; Wait-PubKeyPress }
                '4'  { Invoke-PubScanMenu;          Wait-PubKeyPress }
                '5'  { Invoke-PubExportInventory;   Wait-PubKeyPress }
                '6'  { Invoke-PubLoadInventory;     Wait-PubKeyPress }
                '7'  { Invoke-PubDownloadStep;      Wait-PubKeyPress }
                '8'  { Invoke-PubConvertStep;       Wait-PubKeyPress }
                '9'  { Invoke-PubUploadStep;        Wait-PubKeyPress }
                '10' { Invoke-PubFullPipeline;      Wait-PubKeyPress }
                '11' { Show-PubRecentLog;           Wait-PubKeyPress }
                '12' { Open-PubWorkingFolder;       Wait-PubKeyPress }
                '13' { Invoke-PubSettingsMenu }
                '0'  {
                    Write-PubLog -Level Info -Message 'Exiting.'
                    Disconnect-PubPnP
                    Disconnect-PubGraph
                    Stop-PubLogging
                    Write-Host ''
                    Write-Host ' Goodbye.' -ForegroundColor Cyan
                    return
                }
            }
        } catch {
            # A failure inside one option must never drop the technician out of
            # the tool - log it and redisplay the menu.
            Write-PubLog -Level Error -Message ('Unhandled error in option {0}: {1}' -f $choice, $_.Exception.Message)
            Write-PubLog -Level Debug -Message ($_.ScriptStackTrace)
            Wait-PubKeyPress
        }
    }
}

if (-not $NoRun) {
    Start-PubMenu
}
