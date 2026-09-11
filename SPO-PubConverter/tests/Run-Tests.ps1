<#
.SYNOPSIS
    Offline checks for the SharePoint Publisher File Converter.

.DESCRIPTION
    Exercises everything that does not need a tenant or Microsoft Publisher:
    the CSV schema, local path mapping, filtering, config handling,
    certificate-expiry logic, the skip/overwrite/version rule, and the parts of
    Tom's original conversion script the brief says to preserve.

    Runs on Windows PowerShell 5.1 or PowerShell 7, on any OS. Nothing here
    touches SharePoint, so it is safe to run at any time.

.EXAMPLE
    .\tests\Run-Tests.ps1
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $PSScriptRoot
$moduleRoot  = Join-Path $projectRoot 'modules'
$tempRoot    = Join-Path ([System.IO.Path]::GetTempPath()) ('PubConverterTests_{0}' -f (Get-Date -Format 'yyyyMMddHHmmss'))

New-Item -Path $tempRoot -ItemType Directory -Force | Out-Null

$script:Failures = 0
$script:Checks   = 0

function Assert-PubTest {
    param(
        [Parameter(Mandatory)] $Condition,
        [Parameter(Mandatory)] [string] $Name
    )

    $script:Checks++
    if ($Condition) {
        Write-Host ('  PASS  {0}' -f $Name) -ForegroundColor Green
    } else {
        Write-Host ('  FAIL  {0}' -f $Name) -ForegroundColor Red
        $script:Failures++
    }
}

function Write-PubTestSection {
    param([string] $Name)
    Write-Host ''
    Write-Host ('== {0}' -f $Name) -ForegroundColor Cyan
}

try {
    # -----------------------------------------------------------------------
    Write-PubTestSection 'Every file parses and every module imports'
    # -----------------------------------------------------------------------
    foreach ($file in (Get-ChildItem -Path $projectRoot -Recurse -Include '*.ps1', '*.psm1' -File)) {
        $tokens = $null
        $errors = $null
        [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref] $tokens, [ref] $errors) | Out-Null
        Assert-PubTest (-not $errors -or $errors.Count -eq 0) ('{0} parses' -f $file.Name)
    }

    foreach ($module in @('Logging', 'Config', 'Graph', 'AppRegistration', 'Discovery', 'Convert', 'Upload')) {
        Import-Module (Join-Path $moduleRoot ('{0}.psm1' -f $module)) -Force -DisableNameChecking
    }
    Assert-PubTest $true 'all seven modules import'

    Initialize-PubLogging -Name 'Tests' -LogRoot (Join-Path $tempRoot 'logs') -NoTranscript | Out-Null

    # -----------------------------------------------------------------------
    Write-PubTestSection 'Configuration and secrets'
    # -----------------------------------------------------------------------
    $configPath = Join-Path $tempRoot 'config.json'
    $config     = Get-PubConfig -Path $configPath

    Assert-PubTest ($config['ExistingPdfAction'] -eq 'Version')    'existing-PDF default is Version, not an error'
    Assert-PubTest ($config['UploadConflictAction'] -eq 'Version') 'upload collision default is Version'
    Assert-PubTest ($config['RemoveSourceAfterUpload'] -eq $false) 'source deletion is off by default'
    Assert-PubTest ($config['AuthMethod'] -eq 'AllSites')          'auth method defaults to AllSites'

    $config['TenantId']     = '00000000-1111-2222-3333-444444444444'
    $config['ClientSecret'] = 'this-must-never-be-written'
    Save-PubConfig -Config $config -Path $configPath | Out-Null

    $reloaded = Get-PubConfig -Path $configPath
    Assert-PubTest ($reloaded['TenantId'] -eq '00000000-1111-2222-3333-444444444444') 'config round-trips values'
    Assert-PubTest ([string]::IsNullOrEmpty([string] $reloaded['ClientSecret']))       'secret-looking keys are refused by Save-PubConfig'
    Assert-PubTest ((Get-Content -LiteralPath $configPath -Raw) -notmatch 'this-must-never-be-written') 'no secret material reaches config.json'

    # -----------------------------------------------------------------------
    Write-PubTestSection 'Certificate expiry is checked, with a 30-day warning'
    # -----------------------------------------------------------------------
    $certConfig = New-PubDefaultConfig
    Assert-PubTest ((Test-PubCertificateExpiry -Config $certConfig -Quiet).State -eq 'Unknown') 'no expiry recorded reports Unknown'

    $certConfig['CertificateExpiry'] = (Get-Date).AddDays(400).ToString('yyyy-MM-dd')
    Assert-PubTest ((Test-PubCertificateExpiry -Config $certConfig -Quiet).State -eq 'Valid') 'a distant expiry is Valid'

    $certConfig['CertificateExpiry'] = (Get-Date).AddDays(10).ToString('yyyy-MM-dd')
    Assert-PubTest ((Test-PubCertificateExpiry -Config $certConfig -Quiet).State -eq 'Expiring') 'expiry inside 30 days warns'

    $certConfig['CertificateExpiry'] = (Get-Date).AddDays(-1).ToString('yyyy-MM-dd')
    Assert-PubTest ((Test-PubCertificateExpiry -Config $certConfig -Quiet).State -eq 'Expired') 'a past expiry is Expired'

    # -----------------------------------------------------------------------
    Write-PubTestSection 'CSV inventory schema (brief section 4.3)'
    # -----------------------------------------------------------------------
    $columns   = Get-PubInventoryColumns
    $briefSchema = @('SiteUrl', 'LibraryName', 'FolderPath', 'FileName', 'FileSizeKB', 'LastModified', 'ModifiedBy', 'UniqueId', 'Status', 'Notes')
    Assert-PubTest ($null -eq (Compare-Object $columns[0..9] $briefSchema -SyncWindow 0)) 'the first ten columns are exactly the brief schema, in order'

    $row = New-PubInventoryRow -Values @{
        FileName = 'Newsletter.pub'; SiteUrl = 'https://contoso.sharepoint.com/sites/Marketing'
        LibraryName = 'Shared Documents'; FolderPath = '/2024/Q1 Drafts'; Status = 'Pending'
        DriveId = 'drive1'; ItemId = 'item1'
    }
    Assert-PubTest ($row.PSObject.Properties.Name.Count -eq $columns.Count) 'a new row carries every column'
    Assert-PubTest ($row.Notes -eq '') 'unset columns are empty, not missing'

    Set-PubInventoryStatus -Row $row -Status Downloaded -Notes "first line`nsecond line" | Out-Null
    Assert-PubTest ($row.Status -eq 'Downloaded') 'status updates in place'
    Assert-PubTest ($row.Notes -eq 'first line second line') 'notes are flattened so the CSV stays one row per file'
    Assert-PubTest (-not [string]::IsNullOrWhiteSpace($row.LastAction)) 'LastAction is stamped for resumability'

    # -----------------------------------------------------------------------
    Write-PubTestSection 'Graph folder paths and local working paths'
    # -----------------------------------------------------------------------
    Assert-PubTest ((Get-PubDriveFolderPath -ParentReference ([pscustomobject] @{ path = '/drives/b!x/root:/Marketing/Q1%20Drafts' })) -eq '/Marketing/Q1 Drafts') 'parentReference decodes to a library-relative folder path'
    Assert-PubTest ((Get-PubDriveFolderPath -ParentReference ([pscustomobject] @{ path = '/drives/b!x/root:' })) -eq '/') 'the library root is /'
    Assert-PubTest ((Get-PubDriveFolderPath -ParentReference $null) -eq '/') 'a missing parentReference does not throw'

    $workingConfig = New-PubDefaultConfig
    $workingConfig['DefaultWorkingFolder'] = Join-Path $tempRoot 'working'

    $originalPath = Get-PubLocalPath -Row $row -Kind Original -Config $workingConfig
    $pdfPath      = Get-PubLocalPath -Row $row -Kind Pdf      -Config $workingConfig

    Assert-PubTest ($originalPath -like '*originals*') 'downloads land under /working/originals'
    Assert-PubTest ($originalPath -like '*Shared Documents*') 'the library name is part of the local path'
    Assert-PubTest ($originalPath -like '*Q1 Drafts*Newsletter.pub') 'the folder tree is mirrored locally'
    Assert-PubTest ($pdfPath -like '*converted*Newsletter.pdf') 'PDFs land under /working/converted'

    $twinRow = New-PubInventoryRow -Values @{
        FileName = 'Newsletter.pub'; SiteUrl = 'https://contoso.sharepoint.com/sites/Sales'
        LibraryName = 'Shared Documents'; FolderPath = '/2024/Q1 Drafts'
    }
    Assert-PubTest ((Get-PubLocalPath -Row $twinRow -Config $workingConfig) -ne $originalPath) 'the same file name in two sites cannot collide'
    Assert-PubTest ((ConvertTo-PubSafeSegment -Segment 'a/b:c*d') -eq 'a_b_c_d') 'unsafe path characters are replaced'

    # -----------------------------------------------------------------------
    Write-PubTestSection 'Selection filters (menu option 6)'
    # -----------------------------------------------------------------------
    $rows = @(
        New-PubInventoryRow -Values @{ FileName = 'a.pub'; SiteUrl = 'https://c.sharepoint.com/sites/Marketing'; LibraryName = 'Docs'; FolderPath = '/2024'; Status = 'Pending'; DriveId = 'd'; ItemId = '1' }
        New-PubInventoryRow -Values @{ FileName = 'b.pub'; SiteUrl = 'https://c.sharepoint.com/sites/Sales';     LibraryName = 'Docs'; FolderPath = '/old';  Status = 'Failed';  DriveId = 'd'; ItemId = '2' }
        New-PubInventoryRow -Values @{ FileName = 'c.pub'; SiteUrl = 'https://c.sharepoint.com/sites/Sales';     LibraryName = 'Docs'; FolderPath = '/2024'; Status = 'Pending'; DriveId = 'd'; ItemId = '3' }
    )

    Assert-PubTest (@(Select-PubInventoryRow -Rows $rows -SiteFilter '*Sales*').Count -eq 2)   'filter by site'
    Assert-PubTest (@(Select-PubInventoryRow -Rows $rows -FolderFilter '*2024*').Count -eq 2)  'filter by folder'
    Assert-PubTest (@(Select-PubInventoryRow -Rows $rows -NameFilter 'a*').Count -eq 1)        'filter by file name'
    Assert-PubTest (@(Select-PubInventoryRow -Rows $rows -Status @('Failed')).Count -eq 1)     'filter by status (re-run only failures)'

    $statistics = Get-PubInventoryStatistic -Rows $rows
    Assert-PubTest ($statistics.Total -eq 3 -and $statistics.Pending -eq 2 -and $statistics.Failed -eq 1) 'status counts feed the menu header'

    # -----------------------------------------------------------------------
    Write-PubTestSection 'CSV round trip, including a hand-trimmed CSV'
    # -----------------------------------------------------------------------
    $csvPath = Join-Path $tempRoot 'inventory.csv'
    Export-PubInventory -Rows $rows -Path $csvPath -Config $workingConfig -NoConfigUpdate | Out-Null

    $loaded = Import-PubInventory -Path $csvPath
    Assert-PubTest (@($loaded).Count -eq 3) 'every row survives the round trip'
    Assert-PubTest ($loaded[0].FileName -eq 'a.pub') 'values survive the round trip'

    $trimmedPath = Join-Path $tempRoot 'inventory-trimmed.csv'
    Import-Csv -LiteralPath $csvPath | Select-Object SiteUrl, LibraryName, FolderPath, FileName, Status, DriveId, ItemId |
        Export-Csv -LiteralPath $trimmedPath -NoTypeInformation
    $trimmed = Import-PubInventory -Path $trimmedPath
    Assert-PubTest (@($trimmed).Count -eq 3) 'a CSV trimmed down in Excel still loads'
    Assert-PubTest ($trimmed[0].PSObject.Properties.Name.Count -eq $columns.Count) 'missing optional columns are repaired on load'

    $brokenPath = Join-Path $tempRoot 'inventory-broken.csv'
    Import-Csv -LiteralPath $csvPath | Select-Object SiteUrl, FileName | Export-Csv -LiteralPath $brokenPath -NoTypeInformation
    Assert-PubTest (@(Import-PubInventory -Path $brokenPath).Count -eq 0) 'a CSV with no DriveId/ItemId is rejected with a clear message'

    # -----------------------------------------------------------------------
    Write-PubTestSection 'Scan scope files, including the SharePoint admin centre export'
    # -----------------------------------------------------------------------
    $scopeText = Join-Path $tempRoot 'sites.txt'
    @(
        '# sites for the first test run'
        'https://contoso.sharepoint.com/sites/Marketing'
        ''
        'https://contoso.sharepoint.com/sites/Sales'
    ) | Set-Content -LiteralPath $scopeText

    $textUrls = Get-PubScopeUrl -Path $scopeText
    Assert-PubTest (@($textUrls).Count -eq 2) 'a text scope file reads one URL per line'
    Assert-PubTest ($textUrls -notcontains '# sites for the first test run') 'comment lines are ignored'

    # The admin centre's Active sites export: column is 'URL', not 'SiteUrl'.
    $adminExport = Join-Path $tempRoot 'admin-export.csv'
    @(
        [pscustomobject] @{ 'Site name' = 'Marketing'; 'URL' = 'https://contoso.sharepoint.com/sites/Marketing'; 'Storage used (GB)' = '12.4' }
        [pscustomobject] @{ 'Site name' = 'Sales';     'URL' = 'https://contoso.sharepoint.com/sites/Sales';     'Storage used (GB)' = '3.1' }
    ) | Export-Csv -LiteralPath $adminExport -NoTypeInformation

    $adminUrls = Get-PubScopeUrl -Path $adminExport
    Assert-PubTest (@($adminUrls).Count -eq 2) 'the admin centre export loads unedited'
    Assert-PubTest ($adminUrls[0] -eq 'https://contoso.sharepoint.com/sites/Marketing') 'URLs come from the admin export URL column'

    foreach ($columnName in @('SiteUrl', 'Site URL', 'Url', 'WebUrl')) {
        $aliasPath = Join-Path $tempRoot ('alias-{0}.csv' -f ($columnName -replace '\s', ''))
        (New-Object psobject -Property @{ $columnName = 'https://contoso.sharepoint.com/sites/Ops' }) |
            Select-Object $columnName | Export-Csv -LiteralPath $aliasPath -NoTypeInformation
        Assert-PubTest (@(Get-PubScopeUrl -Path $aliasPath).Count -eq 1) ('a CSV using the "{0}" column is accepted' -f $columnName)
    }

    $noColumnPath = Join-Path $tempRoot 'no-url-column.csv'
    [pscustomobject] @{ 'Site name' = 'Marketing'; 'Owner' = 'tom@contoso.com' } | Export-Csv -LiteralPath $noColumnPath -NoTypeInformation
    Assert-PubTest (@(Get-PubScopeUrl -Path $noColumnPath).Count -eq 0) 'a CSV with no URL column is rejected rather than silently empty'
    Assert-PubTest (@(Get-PubScopeUrl -Path (Join-Path $tempRoot 'does-not-exist.csv')).Count -eq 0) 'a missing scope file is reported, not thrown'

    # -----------------------------------------------------------------------
    Write-PubTestSection 'Upload collision behaviour'
    # -----------------------------------------------------------------------
    Assert-PubTest ((Get-PubConflictBehavior -Action 'Version')   -eq 'rename')  'Version maps to Graph rename (the safe default)'
    Assert-PubTest ((Get-PubConflictBehavior -Action 'Overwrite') -eq 'replace') 'Overwrite maps to Graph replace (opt-in)'
    Assert-PubTest ((Get-PubConflictBehavior -Action 'Skip')      -eq 'fail')    'Skip maps to Graph fail'

    # -----------------------------------------------------------------------
    Write-PubTestSection 'Permission scopes and blast radius'
    # -----------------------------------------------------------------------
    $allSites = Get-PubPermissionCatalog -AuthMethod 'AllSites'
    $selected = Get-PubPermissionCatalog -AuthMethod 'SitesSelected'

    Assert-PubTest ($null -ne ($allSites | Where-Object { $_.Name -eq 'Sites.ReadWrite.All' })) 'AllSites requests Sites.ReadWrite.All'
    Assert-PubTest ($null -ne ($allSites | Where-Object { $_.Name -eq 'Files.ReadWrite.All' })) 'AllSites requests Files.ReadWrite.All'
    Assert-PubTest ($null -eq ($allSites | Where-Object { $_.Name -eq 'Sites.Selected' }))      'AllSites does not mix in Sites.Selected'
    Assert-PubTest ($null -ne ($selected | Where-Object { $_.Name -eq 'Sites.Selected' }))      'SitesSelected requests Sites.Selected'
    Assert-PubTest ($null -eq ($selected | Where-Object { $_.Name -eq 'Files.ReadWrite.All' })) 'SitesSelected never requests tenant-wide write'

    # -----------------------------------------------------------------------
    Write-PubTestSection 'Conversion prerequisites fail closed'
    # -----------------------------------------------------------------------
    if (Test-PubIsWindows) {
        Assert-PubTest ($null -ne (Get-PubWindowsPowerShellPath)) 'Windows PowerShell 5.1 is located on Windows'
    } else {
        Assert-PubTest ((Test-PubPublisherAvailable -Quiet) -eq $false) 'the Publisher check fails closed on a non-Windows host'
        Assert-PubTest ($null -eq (Get-PubWindowsPowerShellPath))       'no Windows PowerShell is reported off-Windows'
    }

    # -----------------------------------------------------------------------
    Write-PubTestSection "Tom's conversion script: preserved mechanics and the new skip/overwrite/version rule"
    # -----------------------------------------------------------------------
    $converterPath = Join-Path $moduleRoot 'Convert.Publisher.ps1'
    $converterText = Get-Content -LiteralPath $converterPath -Raw

    Assert-PubTest ($converterText -match [regex]::Escape('[Microsoft.Office.Interop.Publisher.PbFixedFormatType]::pbFixedFormatTypePDF')) 'the exact ExportAsFixedFormat enum is unchanged'
    Assert-PubTest ($converterText -match 'Add-Type -AssemblyName Microsoft\.Office\.Interop\.Publisher') 'the Add-Type Interop pattern is unchanged'
    Assert-PubTest ($converterText -match 'New-Object -ComObject Publisher\.Application')                 'the Publisher.Application COM pattern is unchanged'
    Assert-PubTest ($converterText -match '(?s)finally \{.*\$app\.Quit\(\)')                              'app.Quit() still runs in the finally block'
    Assert-PubTest ($converterText -match 'ReleaseComObject')                                             'documents are released so MSPUB cannot accumulate'

    $tokens = $null
    $errors = $null
    $converterAst = [System.Management.Automation.Language.Parser]::ParseFile($converterPath, [ref] $tokens, [ref] $errors)
    $parameterNames = $converterAst.ParamBlock.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath }
    foreach ($parameter in @('Filter', 'Recurse', 'JobFile', 'ResultFile', 'ExistingPdfAction')) {
        Assert-PubTest ($parameterNames -contains $parameter) ('the converter still accepts -{0}' -f $parameter)
    }

    $resolveFunction = $converterAst.FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Resolve-PdfDestination'
    }, $true)
    Assert-PubTest ($resolveFunction.Count -eq 1) 'Resolve-PdfDestination is defined'

    # Load just that function so the rule can be tested without Publisher.
    Invoke-Expression $resolveFunction[0].Extent.Text

    $pdfFolder = Join-Path $tempRoot 'pdf'
    New-Item -Path $pdfFolder -ItemType Directory -Force | Out-Null
    $targetPdf = Join-Path $pdfFolder 'Newsletter.pdf'

    $result = Resolve-PdfDestination -PdfPath $targetPdf -Action 'Version'
    Assert-PubTest ($result.Action -eq 'Convert' -and $result.Path -eq $targetPdf) 'no existing PDF: convert straight to the target'

    Set-Content -LiteralPath $targetPdf -Value 'existing'

    $result = Resolve-PdfDestination -PdfPath $targetPdf -Action 'Skip'
    Assert-PubTest ($result.Action -eq 'Skip' -and $result.Message -match 'skipped') 'existing PDF + Skip is a skip with a reason, not an error'

    $result = Resolve-PdfDestination -PdfPath $targetPdf -Action 'Version'
    Assert-PubTest ($result.Path -eq (Join-Path $pdfFolder 'Newsletter (2).pdf')) 'existing PDF + Version writes (2)'

    Set-Content -LiteralPath (Join-Path $pdfFolder 'Newsletter (2).pdf') -Value 'existing'
    $result = Resolve-PdfDestination -PdfPath $targetPdf -Action 'Version'
    Assert-PubTest ($result.Path -eq (Join-Path $pdfFolder 'Newsletter (3).pdf')) 'versioning walks past existing versions'

    $result = Resolve-PdfDestination -PdfPath $targetPdf -Action 'Overwrite'
    Assert-PubTest ($result.Action -eq 'Convert' -and -not (Test-Path -LiteralPath $targetPdf)) 'existing PDF + Overwrite replaces it'

    # -----------------------------------------------------------------------
    Write-PubTestSection 'Logging'
    # -----------------------------------------------------------------------
    Write-PubFileResult -Outcome Converted -FileName 'Newsletter.pub' -Detail 'Newsletter.pdf'
    Write-PubPhaseSummary -Phase 'Test' -Attempted 3 -Succeeded 2 -Failed 1

    $logContent = Get-Content -LiteralPath (Get-PubLogFile) -Raw
    Assert-PubTest ($logContent -match 'CONVERTED \| Newsletter\.pub') 'one summary line per file is logged'
    Assert-PubTest ($logContent -match 'Test summary')                 'the phase summary is logged'
} finally {
    Write-Host ''
    Write-Host ('{0} check(s) run, {1} failure(s).' -f $script:Checks, $script:Failures)

    if (Test-Path -LiteralPath $tempRoot) {
        Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

if ($script:Failures -gt 0) {
    Write-Host 'TESTS FAILED' -ForegroundColor Red
    exit 1
}

Write-Host 'ALL TESTS PASSED' -ForegroundColor Green
exit 0
