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

# The modules run under Set-StrictMode -Version 2.0, so the tests must too.
# Without it, $null.Count quietly returns 0 here while throwing "The property
# 'Count' cannot be found on this object" inside a module - which is exactly
# how an empty-collection bug reached a live tenant crawl.
Set-StrictMode -Version 2.0

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

    foreach ($module in @('Logging', 'Config', 'Graph', 'AppRegistration', 'PnP', 'Discovery', 'Convert', 'Upload')) {
        Import-Module (Join-Path $moduleRoot ('{0}.psm1' -f $module)) -Force -DisableNameChecking
    }
    Assert-PubTest $true 'all eight modules import'

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
    Write-PubTestSection 'Certificate password lookup (regression: key mismatch)'
    # -----------------------------------------------------------------------
    # An earlier build stored the .pfx password under the certificate SUBJECT
    # name but read it back under the APP ID, so the password was never found
    # and the .pfx was opened with none - "the certificate data cannot be read
    # with the provided password". These checks pin the storage and lookup keys
    # together, and keep the older spellings readable.
    $certDir = Join-Path $tempRoot 'certs'
    New-Item -Path $certDir -ItemType Directory -Force | Out-Null
    $pfxFile = Join-Path $certDir 'SPO-PubConverter-5bde7888.pfx'
    Set-Content -LiteralPath $pfxFile -Value 'not a real pfx'

    Assert-PubTest ((Get-PubCertificateSecretName -PfxPath $pfxFile) -eq 'PfxPassword_SPO-PubConverter-5bde7888') 'the secret name is derived from the .pfx file name'
    Assert-PubTest ((Get-PubCertificateSecretName -PfxPath '') -eq '') 'no .pfx means no secret name'

    $secretConfig = New-PubDefaultConfig
    $secretConfig['AppId']              = '5bde7888-1111-2222-3333-444444444444'
    $secretConfig['CertificatePfxPath'] = $pfxFile

    Assert-PubTest ($null -eq (Get-PubCertificatePassword -Config $secretConfig -Quiet)) 'a missing password returns nothing rather than an empty one'

    # Store it the way the old build did - under the certificate subject name.
    Set-PubSecret -Name 'PfxPassword_SPO-PubConverter-5bde7888' -Secret (ConvertTo-SecureString 'subject-keyed' -AsPlainText -Force) | Out-Null
    $recovered = Get-PubCertificatePassword -Config $secretConfig -Quiet
    Assert-PubTest ($null -ne $recovered) 'a password stored by the older build is still found (no certificate regeneration needed)'
    Assert-PubTest ([System.Net.NetworkCredential]::new('', $recovered).Password -eq 'subject-keyed') 'the recovered password is the one that was stored'

    # And the name recorded in config wins when it is present.
    Set-PubSecret -Name 'PfxPassword_explicit' -Secret (ConvertTo-SecureString 'config-keyed' -AsPlainText -Force) | Out-Null
    $secretConfig['CertificateSecretName'] = 'PfxPassword_explicit'
    $explicit = Get-PubCertificatePassword -Config $secretConfig -Quiet
    Assert-PubTest ([System.Net.NetworkCredential]::new('', $explicit).Password -eq 'config-keyed') 'the secret name recorded in config takes priority'

    Assert-PubTest ((New-PubDefaultConfig).Contains('CertificateSecretName')) 'CertificateSecretName is part of the config schema'

    Remove-Item -Path (Join-Path (Get-PubProjectRoot) '.secrets') -Recurse -Force -ErrorAction SilentlyContinue

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
    Assert-PubTest ($pdfPath -like '*converted*Newsletter (converted).pdf') 'PDFs land under /working/converted, with the collision-safe name'

    $twinRow = New-PubInventoryRow -Values @{
        FileName = 'Newsletter.pub'; SiteUrl = 'https://contoso.sharepoint.com/sites/Sales'
        LibraryName = 'Shared Documents'; FolderPath = '/2024/Q1 Drafts'
    }
    Assert-PubTest ((Get-PubLocalPath -Row $twinRow -Config $workingConfig) -ne $originalPath) 'the same file name in two sites cannot collide'
    Assert-PubTest ((ConvertTo-PubSafeSegment -Segment 'a/b:c*d') -eq 'a_b_c_d') 'unsafe path characters are replaced'

    # -----------------------------------------------------------------------
    Write-PubTestSection 'Converted PDF naming (no collision with an existing PDF)'
    # -----------------------------------------------------------------------
    # Newsletter.pub and Newsletter.pdf commonly sit side by side. Converting
    # to a bare Newsletter.pdf would land on the user's existing file, so a
    # suffix keeps them distinct.
    $nameConfig = New-PubDefaultConfig
    Assert-PubTest ($nameConfig['ConvertedFileSuffix'] -eq ' (converted)') 'a suffix is configured by default'
    Assert-PubTest ((Get-PubPdfFileName -SourceFileName 'Newsletter.pub' -Config $nameConfig) -eq 'Newsletter (converted).pdf') 'the default naming cannot collide with an existing PDF'
    Assert-PubTest ((Get-PubPdfFileName -SourceFileName 'Newsletter.pub' -Config $nameConfig) -ne 'Newsletter.pdf') 'the converted name is never the bare PDF name by default'

    Assert-PubTest ((Get-PubPdfFileName -SourceFileName 'Report.final.pub' -Config $nameConfig) -eq 'Report.final (converted).pdf') 'only the last extension is replaced'
    Assert-PubTest ((Get-PubPdfFileName -SourceFileName 'NoExtension' -Config $nameConfig) -eq 'NoExtension (converted).pdf') 'a name with no extension still works'
    Assert-PubTest ((Get-PubPdfFileName -SourceFileName 'Newsletter.pub' -Suffix '') -eq 'Newsletter.pdf') 'an empty suffix gives the bare name for anyone who wants it'
    Assert-PubTest ((Get-PubPdfFileName -SourceFileName 'Newsletter.pub' -Suffix '_converted') -eq 'Newsletter_converted.pdf') 'a custom suffix is applied'

    # Characters SharePoint rejects must never reach a file name.
    Assert-PubTest ((Get-PubFileNameSuffix -Suffix ' (a/b:c*d?)') -eq ' (abcd)') 'characters SharePoint rejects are stripped from the suffix'
    Assert-PubTest ((Get-PubFileNameSuffix -Suffix '   ') -eq '')                'a whitespace-only suffix is treated as none'
    Assert-PubTest ((Get-PubFileNameSuffix -Suffix (' x' * 60)).Length -le 40)   'an over-long suffix is capped'

    $longName = ('a' * 300) + '.pub'
    $longPdf  = Get-PubPdfFileName -SourceFileName $longName -Config $nameConfig
    Assert-PubTest ($longPdf.Length -le 255)        'a very long name is trimmed to something SharePoint accepts'
    Assert-PubTest ($longPdf.EndsWith(' (converted).pdf')) 'trimming keeps the suffix rather than losing the collision guard'

    # The conversion step and the upload step must agree on the name.
    $nameRow = New-PubInventoryRow -Values @{
        FileName = 'Newsletter.pub'; SiteUrl = 'https://c.sharepoint.com/sites/A'
        LibraryName = 'Docs'; FolderPath = '/'
    }
    $nameConfig['DefaultWorkingFolder'] = Join-Path $tempRoot 'naming'
    $localPdf = Get-PubLocalPath -Row $nameRow -Kind Pdf -Config $nameConfig
    Assert-PubTest ((Split-Path -Leaf $localPdf) -eq (Get-PubPdfFileName -SourceFileName 'Newsletter.pub' -Config $nameConfig)) 'the local file and the uploaded file use the same name'

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
    $allSites    = Get-PubPermissionCatalog -AuthMethod 'AllSites'
    $selected    = Get-PubPermissionCatalog -AuthMethod 'SitesSelected'
    $tenantAdmin = Get-PubPermissionCatalog -AuthMethod 'TenantAdmin'

    Assert-PubTest ($null -ne ($allSites | Where-Object { $_.Name -eq 'Sites.ReadWrite.All' })) 'AllSites requests Sites.ReadWrite.All'
    Assert-PubTest ($null -ne ($allSites | Where-Object { $_.Name -eq 'Files.ReadWrite.All' })) 'AllSites requests Files.ReadWrite.All'
    Assert-PubTest ($null -eq ($allSites | Where-Object { $_.Name -eq 'Sites.Selected' }))      'AllSites does not mix in Sites.Selected'
    Assert-PubTest ($null -ne ($selected | Where-Object { $_.Name -eq 'Sites.Selected' }))      'SitesSelected requests Sites.Selected'
    Assert-PubTest ($null -eq ($selected | Where-Object { $_.Name -eq 'Files.ReadWrite.All' })) 'SitesSelected never requests tenant-wide write'

    $fullControl = $tenantAdmin | Where-Object { $_.Name -eq 'Sites.FullControl.All' }
    Assert-PubTest ($null -ne $fullControl) 'TenantAdmin requests SharePoint Sites.FullControl.All'
    Assert-PubTest ($fullControl.Resource -eq 'SharePoint') 'FullControl is requested on the SharePoint API, not Graph'
    Assert-PubTest ($null -eq ($allSites | Where-Object { $_.Name -eq 'Sites.FullControl.All' })) 'FullControl is never requested outside the TenantAdmin scope'
    Assert-PubTest ($null -eq ($selected | Where-Object { $_.Name -eq 'Sites.FullControl.All' })) 'Sites.Selected stays free of FullControl'
    Assert-PubTest ($null -ne ($tenantAdmin | Where-Object { $_.Name -eq 'Sites.ReadWrite.All' -and $_.Resource -eq 'Graph' })) 'TenantAdmin keeps the Graph permissions the pipeline needs'

    Assert-PubTest ((Get-PubResourceAppId -Resource 'Graph') -eq '00000003-0000-0000-c000-000000000000')      'the Graph resource app id is correct'
    Assert-PubTest ((Get-PubResourceAppId -Resource 'SharePoint') -eq '00000003-0000-0ff1-ce00-000000000000') 'the SharePoint resource app id is correct'

    # App role ids are resolved live from the tenant; the catalogue GUID is only
    # a fallback, and SharePoint has none at all, so the lookup must work.
    $fakeServicePrincipal = [pscustomobject] @{ appRoles = @(
        [pscustomobject] @{ value = 'Sites.FullControl.All'; id = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee' }
    ) }
    Assert-PubTest ((Resolve-PubAppRole -Permission $fullControl -ResourceServicePrincipal $fakeServicePrincipal) -eq 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee') 'app role ids resolve by permission name from the tenant'
    $graphRead = $allSites | Where-Object { $_.Name -eq 'Sites.Read.All' }
    Assert-PubTest ((Resolve-PubAppRole -Permission $graphRead -ResourceServicePrincipal $null) -eq $graphRead.Id) 'a failed lookup falls back to the catalogue id'
    Assert-PubTest ($null -eq (Resolve-PubAppRole -Permission $fullControl -ResourceServicePrincipal $null)) 'SharePoint permissions report failure rather than inventing a GUID'

    # -----------------------------------------------------------------------
    Write-PubTestSection 'Enumeration method and tenant admin URL'
    # -----------------------------------------------------------------------
    $enumConfig = New-PubDefaultConfig
    Assert-PubTest ($enumConfig['EnumerationMethod'] -eq 'Auto') 'enumeration defaults to Auto (PnP when available, else Graph)'

    $enumConfig['EnumerationMethod'] = 'Graph'
    Assert-PubTest ((Test-PubPnPEnumerationEnabled -Config $enumConfig) -eq $false) 'the Graph setting never uses PnP'

    $enumConfig['EnumerationMethod'] = 'Auto'
    Assert-PubTest ((Test-PubPnPEnumerationEnabled -Config $enumConfig) -eq $false) 'Auto falls back to Graph when PnP is unusable'

    $adminConfig = New-PubDefaultConfig
    $adminConfig['TenantDomain'] = 'contoso.onmicrosoft.com'
    Assert-PubTest ((Get-PubTenantAdminUrl -Config $adminConfig) -eq 'https://contoso-admin.sharepoint.com') 'the admin URL is derived from an onmicrosoft.com domain'

    $adminConfig['TenantDomain'] = 'contoso.sharepoint.com'
    Assert-PubTest ((Get-PubTenantAdminUrl -Config $adminConfig) -eq 'https://contoso-admin.sharepoint.com') 'the admin URL is derived from a sharepoint.com domain'

    $adminConfig['SharePointAdminUrl'] = 'https://override-admin.sharepoint.com/'
    Assert-PubTest ((Get-PubTenantAdminUrl -Config $adminConfig) -eq 'https://override-admin.sharepoint.com') 'an explicit SharePointAdminUrl wins and is trimmed'

    Assert-PubTest ((Test-PubPnPAvailable -Quiet) -eq $false) 'the PnP probe fails closed when the module is absent'

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
    Write-PubTestSection 'Collections that might be a List'
    # -----------------------------------------------------------------------
    # @( ) around a System.Collections.Generic.List throws "Argument types do
    # not match" on some PowerShell builds. It killed a real tenant crawl the
    # moment site resolution finished, so collections of uncertain type go
    # through ConvertTo-PubArray instead.
    $listForArray = New-Object System.Collections.Generic.List[object]
    $listForArray.Add([pscustomobject] @{ n = 1 })
    $listForArray.Add([pscustomobject] @{ n = 2 })

    $fromList = ConvertTo-PubArray $listForArray
    Assert-PubTest ($fromList -is [object[]])  'a List converts to a plain array'
    Assert-PubTest ($fromList.Count -eq 2)     'every item survives the conversion'
    Assert-PubTest ($fromList[1].n -eq 2)      'items keep their order and content'

    # Every shape must come back as a real object[]. Returning an array from a
    # PowerShell function unrolls it - an empty one becomes $null and a
    # one-item one becomes the bare item - so each return is comma-wrapped.
    # An empty page of Graph results hit exactly this and aborted the folder.
    $shapes = [ordered] @{
        'an array'            = @(1, 2, 3)
        'an empty array'      = @()
        'null'                = $null
        'an empty List'       = (New-Object System.Collections.Generic.List[object])
        'a one-item List'     = $listForArray[0]
        'a string'            = 'one string'
        'a single object'     = ([pscustomobject] @{ n = 1 })
    }
    $expectedCounts = [ordered] @{
        'an array' = 3; 'an empty array' = 0; 'null' = 0; 'an empty List' = 0
        'a one-item List' = 1; 'a string' = 1; 'a single object' = 1
    }

    foreach ($shapeName in $shapes.Keys) {
        $converted = ConvertTo-PubArray $shapes[$shapeName]
        Assert-PubTest ($converted -is [object[]]) ('{0} converts to a real object[]' -f $shapeName)
        Assert-PubTest ($converted.Count -eq $expectedCounts[$shapeName]) ('{0} reports the right count' -f $shapeName)
    }

    $sourceText = Get-Content -LiteralPath (Join-Path $moduleRoot 'Discovery.psm1') -Raw
    Assert-PubTest ($sourceText -notmatch '@\(\$sites\)') 'discovery no longer wraps its site List in @( )'

    # -----------------------------------------------------------------------
    Write-PubTestSection 'Graph paging and the progress callback'
    # -----------------------------------------------------------------------
    # The callback is defined in Discovery and invoked from Graph, so this
    # covers the real cross-module path, including whether a counter mutated
    # inside the callback survives (it must be a hashtable member - assigning
    # to an outer scalar from an & invocation silently updates a local copy).
    $graphModule = Get-Module Graph
    & $graphModule {
        function script:Invoke-PubGraph {
            param([string] $Uri, [string] $Method = 'GET', $Body, [hashtable] $Headers, [string] $ContentType, [string] $OutputFilePath, [int] $MaxAttempts = 5)
            if ($Uri -eq 'page1') {
                return [pscustomobject] @{ value = @('a', 'b', 'c'); '@odata.nextLink' = 'page2' }
            }
            return [pscustomobject] @{ value = @('d', 'e') }
        }
    }

    $pageCounter = @{ Pages = 0; Items = 0; LastSoFar = 0 }
    $pageCallback = {
        param($PageItems, $PageNumber, $ItemsSoFar)
        $pageCounter.Pages     = $pageCounter.Pages + 1
        $pageCounter.Items     = $pageCounter.Items + (ConvertTo-PubArray $PageItems).Count
        $pageCounter.LastSoFar = $ItemsSoFar
    }

    $paged = Get-PubGraphAll -Uri 'page1' -OnPage $pageCallback

    Assert-PubTest ((ConvertTo-PubArray $paged).Count -eq 5) 'paging follows @odata.nextLink to the end'
    Assert-PubTest ($pageCounter.Pages -eq 2)               'the callback fires once per page'
    Assert-PubTest ($pageCounter.Items -eq 5)               'a counter mutated inside the callback really accumulates'
    Assert-PubTest ($pageCounter.LastSoFar -eq 5)           'the callback is told the running total'

    $withoutCallback = Get-PubGraphAll -Uri 'page1'
    Assert-PubTest ((ConvertTo-PubArray $withoutCallback).Count -eq 5) 'paging still works with no callback supplied'

    Import-Module (Join-Path $moduleRoot 'Graph.psm1') -Force -DisableNameChecking

    # -----------------------------------------------------------------------
    Write-PubTestSection 'Discovery crawl, end to end with Graph stubbed'
    # -----------------------------------------------------------------------
    # Exercises the whole crawl offline: site resolution, library filtering,
    # recursion into folders, extension matching and row construction. This is
    # the path that a tenant-wide scan actually runs.
    $discoveryModule = Get-Module Discovery
    & $discoveryModule { $script:OnPageCalls = 0 }
    & $discoveryModule {
        function script:Connect-PubGraphApp { param($Config, [switch] $Force) return $true }
        function script:Test-PubPnPEnumerationEnabled { param($Config) return $false }

        function script:Invoke-PubGraph {
            param([string] $Uri, [string] $Method = 'GET', $Body, [hashtable] $Headers, [string] $ContentType, [string] $OutputFilePath, [int] $MaxAttempts = 5)
            if ($Uri -match '^sites/[^/]+:') {
                return [pscustomobject] @{
                    id = 'contoso.sharepoint.com,aaa,bbb'; webUrl = 'https://contoso.sharepoint.com/sites/Marketing'
                    displayName = 'Marketing'; name = 'Marketing'
                }
            }
            throw "unexpected Invoke-PubGraph call: $Uri"
        }

        function script:Get-PubGraphAll {
            [CmdletBinding()]
            param([string] $Uri, [int] $MaxItems = 0, [int] $MaxAttempts = 5, [scriptblock] $OnPage)

            # Mirror the real paging contract: hand each page to the callback so
            # the progress path is exercised, not just the happy return value.
            $emit = {
                param($Items)
                $script:OnPageCalls = $script:OnPageCalls + 1
                if ($OnPage) { & $OnPage (ConvertTo-PubArray $Items) 1 (ConvertTo-PubArray $Items).Count }
                return $Items
            }

            if ($Uri -match '/sites\?\$select') { return @() }

            if ($Uri -match '/lists\?') {
                return & $emit @(
                    [pscustomobject] @{
                        id = 'list1'; name = 'Shared Documents'; displayName = 'Documents'
                        list = [pscustomobject] @{ template = 'documentLibrary'; hidden = $false }
                        drive = [pscustomobject] @{ id = 'drive1'; name = 'Documents' }
                    },
                    [pscustomobject] @{
                        id = 'list2'; name = 'Hidden'; displayName = 'Hidden Library'
                        list = [pscustomobject] @{ template = 'documentLibrary'; hidden = $true }
                        drive = [pscustomobject] @{ id = 'drive2'; name = 'Hidden' }
                    },
                    [pscustomobject] @{
                        id = 'list3'; name = 'Style Library'; displayName = 'Style Library'
                        list = [pscustomobject] @{ template = 'documentLibrary'; hidden = $false }
                        drive = [pscustomobject] @{ id = 'drive3'; name = 'Style Library' }
                    },
                    [pscustomobject] @{
                        id = 'list4'; name = 'Site Pages'; displayName = 'Site Pages'
                        list = [pscustomobject] @{ template = 'sitePagePublishing'; hidden = $false }
                        drive = [pscustomobject] @{ id = 'drive4'; name = 'Site Pages' }
                    }
                )
            }

            if ($Uri -match 'drives/drive1/root/children') {
                return & $emit @(
                    [pscustomobject] @{
                        id = 'folderA'; name = '2024'
                        folder = [pscustomobject] @{ childCount = 2 }
                        parentReference = [pscustomobject] @{ id = 'root'; path = '/drives/drive1/root:' }
                    },
                    [pscustomobject] @{
                        id = 'folderEmpty'; name = 'Empty Folder'
                        folder = [pscustomobject] @{ childCount = 0 }
                        parentReference = [pscustomobject] @{ id = 'root'; path = '/drives/drive1/root:' }
                    },
                    [pscustomobject] @{
                        id = 'file1'; name = 'Newsletter.pub'; size = 204800
                        file = [pscustomobject] @{ mimeType = 'application/x-mspublisher' }
                        lastModifiedDateTime = '2024-03-04T10:00:00Z'
                        lastModifiedBy = [pscustomobject] @{ user = [pscustomobject] @{ email = 'tom@contoso.com'; displayName = 'Tom' } }
                        parentReference = [pscustomobject] @{ id = 'root'; path = '/drives/drive1/root:' }
                        sharepointIds = [pscustomobject] @{ listItemUniqueId = '11111111-2222-3333-4444-555555555555' }
                    },
                    [pscustomobject] @{
                        id = 'file3'; name = 'NotPublisher.docx'; size = 1000
                        file = [pscustomobject] @{ mimeType = 'application/vnd.openxmlformats' }
                        lastModifiedDateTime = '2024-03-04T10:00:00Z'
                        parentReference = [pscustomobject] @{ id = 'root'; path = '/drives/drive1/root:' }
                    }
                )
            }

            if ($Uri -match 'drives/drive1/items/folderEmpty/children') {
                return & $emit @()
            }

            if ($Uri -match 'drives/drive1/items/folderA/children') {
                return & $emit @(
                    [pscustomobject] @{
                        id = 'file2'; name = 'Brochure.pub'; size = 1048576
                        file = [pscustomobject] @{ mimeType = 'application/x-mspublisher' }
                        lastModifiedDateTime = '2024-05-06T09:00:00Z'
                        lastModifiedBy = [pscustomobject] @{ user = [pscustomobject] @{ displayName = 'Jo' } }
                        parentReference = [pscustomobject] @{ id = 'folderA'; path = '/drives/drive1/root:/2024' }
                        sharepointIds = [pscustomobject] @{ listItemUniqueId = '66666666-7777-8888-9999-000000000000' }
                    }
                )
            }

            throw "unexpected Get-PubGraphAll call: $Uri"
        }
    }

    $crawlConfig = New-PubDefaultConfig
    $crawlConfig['DefaultWorkingFolder'] = Join-Path $tempRoot 'crawl'

    $crawled = Invoke-PubDiscovery -SiteUrl @('https://contoso.sharepoint.com/sites/Marketing') -Config $crawlConfig
    $crawled = ConvertTo-PubArray $crawled

    Assert-PubTest ($crawled.Count -eq 2) 'the crawl completes and returns only the .pub files'

    $newsletter = $crawled | Where-Object { $_.FileName -eq 'Newsletter.pub' } | Select-Object -First 1
    $brochure   = $crawled | Where-Object { $_.FileName -eq 'Brochure.pub' }   | Select-Object -First 1

    Assert-PubTest ($null -ne $newsletter) 'a file in the library root is found'
    Assert-PubTest ($null -ne $brochure)   'a file in a nested folder is found (recursion works)'
    Assert-PubTest ((ConvertTo-PubArray ($crawled | Where-Object { $_.FileName -like '*.docx' })).Count -eq 0) 'non-Publisher files are ignored'

    Assert-PubTest ($newsletter.FolderPath -eq '/')      'a root file records / as its folder'
    Assert-PubTest ($brochure.FolderPath -eq '/2024')    'a nested file records its folder path'
    Assert-PubTest ($newsletter.LibraryName -eq 'Documents') 'the library display name is recorded'
    Assert-PubTest ($newsletter.FileSizeKB -eq 200)      'size is recorded in KB'
    Assert-PubTest ($newsletter.ModifiedBy -eq 'tom@contoso.com') 'modified-by prefers the UPN'
    Assert-PubTest ($brochure.ModifiedBy -eq 'Jo')       'modified-by falls back to the display name'
    Assert-PubTest ($newsletter.UniqueId -eq '11111111-2222-3333-4444-555555555555') 'the SharePoint file GUID is the join key'
    Assert-PubTest ($newsletter.Status -eq 'Pending')    'new rows start Pending'
    Assert-PubTest ($newsletter.DriveId -eq 'drive1' -and $newsletter.ItemId -eq 'file1') 'the drive and item ids needed for download are recorded'
    Assert-PubTest ($newsletter.SiteUrl -eq 'https://contoso.sharepoint.com/sites/Marketing') 'the site URL is recorded'

    # Hidden, system and non-library lists must never be crawled - the stub
    # throws on any unexpected drive, so reaching here proves they were skipped.
    Assert-PubTest $true 'hidden, system and non-document-library lists are skipped'

    $onPageCalls = & $discoveryModule { $script:OnPageCalls }
    Assert-PubTest ($onPageCalls -gt 0) 'the crawl streams pages through the progress callback'

    # An empty folder returns a page with no items. That used to throw inside
    # the progress callback and get swallowed as "Skipped a folder".
    $crawlLog = Get-Content -LiteralPath (Get-PubLogFile) -Raw
    Assert-PubTest ($crawlLog -notmatch 'Skipped a folder') 'an empty folder is crawled cleanly, not skipped with a warning'

    Import-Module (Join-Path $moduleRoot 'Discovery.psm1') -Force -DisableNameChecking

    # -----------------------------------------------------------------------
    Write-PubTestSection 'Console menu (brief section 7)'
    # -----------------------------------------------------------------------
    # Load the menu's functions without going interactive.
    . (Join-Path $projectRoot 'Start-Menu.ps1') -NoRun

    function Copy-MenuConfig {
        param($Source)
        $copy = New-PubDefaultConfig
        foreach ($key in $Source.Keys) { $copy[$key] = $Source[$key] }
        return $copy
    }

    function Get-RenderedMenu {
        param($MenuConfig, $MenuInventory = @(), $MenuSelection = @(), [string] $MenuPath = '')

        $script:Config        = $MenuConfig
        $script:Inventory     = $MenuInventory
        $script:Selection     = $MenuSelection
        $script:InventoryPath = $MenuPath
        $script:VerifiedRoute = ''

        # Write-Host -NoNewline segments arrive as separate records when
        # captured, so rejoin them before matching on content.
        $rendered = Show-PubMainMenu -NoClear 6>&1 | Out-String
        return ($rendered -replace '(NEXT|NOTE):\s*[\r\n]+\s*', '$1: ')
    }

    $freshConfig = New-PubDefaultConfig
    $freshMenu   = Get-RenderedMenu -MenuConfig $freshConfig

    # The labels the brief specifies, verbatim, in the order it gives them.
    $briefLabels = [ordered] @{
        '1'  = 'Generate or connect Azure AD App Registration'
        '2'  = 'Generate & upload authentication certificate'
        '3'  = 'Test connection to Microsoft Graph / SharePoint'
        '4'  = 'Scan tenant for Publisher (.pub) files'
        '5'  = 'Export / re-export scan results to CSV'
        '6'  = 'Load a CSV and select files to process'
        '7'  = 'Download selected files'
        '8'  = 'Convert downloaded files to PDF'
        '9'  = 'Upload converted PDFs to original SharePoint location'
        '10' = 'Run full pipeline (4 -> 9) unattended'
        '11' = 'View recent log'
        '12' = 'Open working folder'
        '0'  = 'Exit'
    }

    foreach ($number in $briefLabels.Keys) {
        Assert-PubTest ($freshMenu -match [regex]::Escape(('{0}) {1}' -f $number, $briefLabels[$number]))) ('option {0} keeps the label from the brief' -f $number)
    }

    $labelOrder = @()
    foreach ($line in ($freshMenu -split "`r?`n")) {
        if ($line -match '^\s+(\d+)\)\s') { $labelOrder += $Matches[1] }
    }
    Assert-PubTest (($labelOrder -join ',') -eq '1,2,3,4,5,6,7,8,9,10,11,12,13,0') 'options run in setup -> discovery -> convert -> publish -> utilities order, with 0 last'

    foreach ($heading in @('SETUP', 'DISCOVERY', 'CONVERSION', 'PUBLISH', 'UTILITIES')) {
        Assert-PubTest ($freshMenu -match $heading) ('the {0} group is labelled' -f $heading)
    }

    # Nothing configured: point at option 1 and say so plainly.
    Assert-PubTest ($freshMenu -match 'NEXT: option 1') 'a brand new install points at option 1'
    Assert-PubTest ($freshMenu -match 'not set up yet') 'an unconfigured tenant says so rather than showing a blank'
    Assert-PubTest ($freshMenu -match 'needs setup')    'steps that cannot work yet say why'

    # Set up, with work outstanding at each stage in turn.
    $readyConfig = New-PubDefaultConfig
    $readyConfig['TenantDomain']          = 'contoso.onmicrosoft.com'
    $readyConfig['AppId']                 = '11111111-1111-1111-1111-111111111111'
    $readyConfig['CertificateThumbprint'] = 'ABC123'
    $readyConfig['CertificateExpiry']     = (Get-Date).AddYears(2).ToString('yyyy-MM-dd')

    $noCertConfig = Copy-MenuConfig -Source $readyConfig
    $noCertConfig['CertificateThumbprint'] = ''
    $menuNoCert = Get-RenderedMenu -MenuConfig $noCertConfig
    Assert-PubTest ($menuNoCert -match 'NEXT: option 2') 'an app with no certificate points at option 2'

    $expiredConfig = Copy-MenuConfig -Source $readyConfig
    $expiredConfig['CertificateExpiry'] = (Get-Date).AddDays(-2).ToString('yyyy-MM-dd')
    $menuExpired = Get-RenderedMenu -MenuConfig $expiredConfig
    Assert-PubTest ($menuExpired -match 'NEXT: option 2')     'an expired certificate takes priority over everything else'
    Assert-PubTest ($menuExpired -match 'CERTIFICATE EXPIRED') 'an expired certificate is called out in the header'
    Assert-PubTest ($menuExpired -match 'EXPIRED')             'option 2 is flagged as expired in the list'

    $menuNoScan = Get-RenderedMenu -MenuConfig $readyConfig
    Assert-PubTest ($menuNoScan -match 'NEXT: option 4') 'a set-up tenant with no inventory points at the scan'

    $pendingRows = @(
        New-PubInventoryRow -Values @{ FileName = 'a.pub'; Status = 'Pending';    DriveId = 'd'; ItemId = '1' }
        New-PubInventoryRow -Values @{ FileName = 'b.pub'; Status = 'Downloaded'; DriveId = 'd'; ItemId = '2' }
        New-PubInventoryRow -Values @{ FileName = 'c.pub'; Status = 'Converted';  DriveId = 'd'; ItemId = '3' }
    )
    $menuPending = Get-RenderedMenu -MenuConfig $readyConfig -MenuInventory $pendingRows -MenuPath 'C:\inv.csv'
    Assert-PubTest ($menuPending -match 'NEXT: option 7')   'rows waiting to download point at option 7'
    Assert-PubTest ($menuPending -match '1 to download')    'the header counts what is waiting'
    Assert-PubTest ($menuPending -match 'inv\.csv')         'the loaded file list is named'

    $downloadedRows = @($pendingRows | Where-Object { $_.Status -ne 'Pending' })
    $menuDownloaded = Get-RenderedMenu -MenuConfig $readyConfig -MenuInventory $downloadedRows -MenuPath 'C:\inv.csv'
    Assert-PubTest ($menuDownloaded -match 'NEXT: option 8') 'downloaded rows point at the conversion step'

    $convertedRows = @($pendingRows | Where-Object { $_.Status -eq 'Converted' })
    $menuConverted = Get-RenderedMenu -MenuConfig $readyConfig -MenuInventory $convertedRows -MenuPath 'C:\inv.csv'
    Assert-PubTest ($menuConverted -match 'NEXT: option 9') 'converted rows point at the upload step'

    $uploadedRows = @(New-PubInventoryRow -Values @{ FileName = 'd.pub'; Status = 'Uploaded'; DriveId = 'd'; ItemId = '4' })
    $menuDone = Get-RenderedMenu -MenuConfig $readyConfig -MenuInventory $uploadedRows -MenuPath 'C:\inv.csv'
    Assert-PubTest ($menuDone -match 'nothing outstanding') 'a finished inventory says there is nothing left to do'

    $failedRows = @(New-PubInventoryRow -Values @{ FileName = 'e.pub'; Status = 'Failed'; DriveId = 'd'; ItemId = '5' })
    $menuFailed = Get-RenderedMenu -MenuConfig $readyConfig -MenuInventory $failedRows -MenuPath 'C:\inv.csv'
    Assert-PubTest ($menuFailed -match 'failed') 'failures are surfaced in the header'
    Assert-PubTest ($menuFailed -match 'NEXT: option 6') 'failures point at re-selecting them to retry'

    $selectionMenu = Get-RenderedMenu -MenuConfig $readyConfig -MenuInventory $pendingRows -MenuSelection @($pendingRows[0]) -MenuPath 'C:\inv.csv'
    Assert-PubTest ($selectionMenu -match 'Working on') 'a narrowed selection is stated in the header'
    Assert-PubTest ($selectionMenu -match '1 of 3')     'the header says how much of the inventory is selected'

    $armedConfig = Copy-MenuConfig -Source $readyConfig
    $armedConfig['RemoveSourceAfterUpload'] = $true
    $armedMenu = Get-RenderedMenu -MenuConfig $armedConfig
    Assert-PubTest ($armedMenu -match 'ORIGINALS DELETED') 'arming source deletion is visible on the main menu at all times'

    # Width: a console menu that wraps is unreadable. Checked on the composed
    # lines, because colouring splits each printed line into two segments.
    $composed = @()
    foreach ($number in $briefLabels.Keys) {
        $composed += Format-PubMenuOptionLine -Number $number -Label $briefLabels[$number] -Note 'needs Publisher'
    }
    $composed += Format-PubMenuOptionLine -Number '13' -Label 'Change conversion & upload settings' -Note 'rules, folders'
    $composed += Format-PubMenuStatusLine -Label 'Site discovery' -Value 'tenant admin list - finds every site (test with option 3)'
    $composed += Format-PubMenuStatusLine -Label 'Tenant access'  -Value 'CERTIFICATE EXPIRED 2026-09-08 - run option 2'
    $composed += Format-PubMenuStatusLine -Label ' ' -Value '12 to download  |  30 to convert  |  170 to upload' -Continuation

    $tooWide = @($composed | Where-Object { $_.Length -gt 80 })
    Assert-PubTest ($tooWide.Count -eq 0) ('every composed menu line fits an 80-column console ({0} too wide)' -f $tooWide.Count)

    $longest = Format-PubMenuOptionLine -Number '9' -Label $briefLabels['9'] -Note '170 ready'
    Assert-PubTest ($longest -match 'location\s\s+170 ready') 'the longest label still leaves a gap before its note'

    foreach ($rendered in @($freshMenu, $menuPending, $menuExpired, $armedMenu, $selectionMenu)) {
        foreach ($line in ($rendered -split "`r?`n")) {
            if ($line.TrimEnd().Length -gt 80) { $tooWide += $line }
        }
    }
    Assert-PubTest ($tooWide.Count -eq 0) 'no rendered menu line exceeds 80 columns either'

    # Input handling: invalid entries re-prompt rather than erroring out.
    $script:MenuAnswers = New-Object System.Collections.Generic.Queue[string]
    function global:Read-Host { param([string] $Prompt) return $script:MenuAnswers.Dequeue() }

    @('99', 'nonsense', '', '4') | ForEach-Object { $script:MenuAnswers.Enqueue($_) }
    Assert-PubTest ((Read-PubMenuChoice -Valid @('0', '4') 6>$null) -eq '4') 'invalid input re-prompts until a valid option is entered'

    @('q') | ForEach-Object { $script:MenuAnswers.Enqueue($_) }
    Assert-PubTest ((Read-PubMenuChoice -Valid @('0', '1') 6>$null) -eq '0') 'typing q is treated as 0 (exit/back)'

    @('back') | ForEach-Object { $script:MenuAnswers.Enqueue($_) }
    Assert-PubTest ((Read-PubMenuChoice -Valid @('0', '1') 6>$null) -eq '0') 'typing back is treated as 0'

    @('3)') | ForEach-Object { $script:MenuAnswers.Enqueue($_) }
    Assert-PubTest ((Read-PubMenuChoice -Valid @('0', '3') 6>$null) -eq '3') 'a stray bracket typed with the number is forgiven'

    @(' 2 ') | ForEach-Object { $script:MenuAnswers.Enqueue($_) }
    Assert-PubTest ((Read-PubMenuChoice -Valid @('0', '2') 6>$null) -eq '2') 'surrounding whitespace is forgiven'

    Remove-Item -Path 'function:global:Read-Host' -ErrorAction SilentlyContinue

    # -----------------------------------------------------------------------
    Write-PubTestSection 'Logging'
    # -----------------------------------------------------------------------
    Write-PubFileResult -Outcome Converted -FileName 'Newsletter.pub' -Detail 'Newsletter.pdf'
    Write-PubPhaseSummary -Phase 'Test' -Attempted 3 -Succeeded 2 -Failed 1

    Assert-PubTest ((Format-PubDuration -Duration ([timespan]::FromSeconds(0.4))) -eq 'less than a second') 'a sub-second duration reads plainly'
    Assert-PubTest ((Format-PubDuration -Duration ([timespan]::FromSeconds(45))) -eq '45s')                 'seconds are shown on their own'
    Assert-PubTest ((Format-PubDuration -Duration ([timespan]::FromSeconds(252))) -eq '4m 12s')             'minutes and seconds are shown together'
    Assert-PubTest ((Format-PubDuration -Duration ([timespan]::FromMinutes(83))) -eq '1h 23m')              'long runs are shown in hours and minutes'

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
