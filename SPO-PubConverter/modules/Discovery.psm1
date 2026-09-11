<#
.SYNOPSIS
    Phase 1 - crawl SharePoint Online for Microsoft Publisher (.pub) files and
    export the CSV inventory that every later phase reads.

.DESCRIPTION
    Implements brief section 4. Also owns the inventory CSV schema (section
    4.3) and the read/write helpers for it, which Convert.psm1 and Upload.psm1
    import so there is exactly one definition of the file format.

    Crawl behaviour:
      * Every site the app can see, or a scoped subset supplied by the operator.
      * Every non-hidden document library, recursed to the bottom of the folder
        tree, including subsites.
      * Match on the .pub extension only - files are never opened at this stage.
      * Items in the Recycle Bin are not returned by driveItem enumeration, so
        they are excluded automatically.
      * Live progress, because a tenant-wide crawl can run for a long time.
#>

Set-StrictMode -Version 2.0

foreach ($dependency in @('Logging', 'Config', 'Graph')) {
    if (-not (Get-Module -Name $dependency)) {
        Import-Module (Join-Path $PSScriptRoot ("{0}.psm1" -f $dependency)) -Force -DisableNameChecking
    }
}

# Libraries that are plumbing rather than user content.
$script:SystemLibraries = @(
    'Form Templates'
    'Style Library'
    'Preservation Hold Library'
    'Site Collection Documents'
    'Site Collection Images'
    'Customized Reports'
    'Teams Wiki Data'
)

function Get-PubInventoryColumns {
    <#
    .SYNOPSIS
        The CSV column order (brief section 4.3, plus operational columns).

    .DESCRIPTION
        The first ten columns are exactly the schema in the brief. The trailing
        columns are what the later phases need to find their way back to the
        file without re-crawling the tenant, and to resume after an
        interruption. They are additive - a CSV trimmed by the operator in
        Excel still loads, as long as the columns are left intact.
    #>
    [CmdletBinding()]
    param()

    return @(
        'SiteUrl'
        'LibraryName'
        'FolderPath'
        'FileName'
        'FileSizeKB'
        'LastModified'
        'ModifiedBy'
        'UniqueId'
        'Status'
        'Notes'
        # --- operational columns ---
        'SiteId'
        'DriveId'
        'ItemId'
        'ParentItemId'
        'LocalPath'
        'PdfPath'
        'LastAction'
    )
}

function New-PubInventoryRow {
    <#
    .SYNOPSIS
        Builds one inventory row with every column present.
    #>
    [CmdletBinding()]
    param(
        [hashtable] $Values = @{}
    )

    $row = [ordered] @{}
    foreach ($column in (Get-PubInventoryColumns)) {
        if ($Values.ContainsKey($column)) { $row[$column] = $Values[$column] }
        else { $row[$column] = '' }
    }

    return [pscustomobject] $row
}

function Get-PubSiteList {
    <#
    .SYNOPSIS
        Returns the sites to crawl.

    .PARAMETER ScopePath
        Optional path to a text file of site URLs (one per line) or a CSV with a
        SiteUrl column. Lines starting with # are ignored.

    .PARAMETER SiteUrl
        Optional explicit list of site URLs (comma separated at the menu).

    .PARAMETER IncludePersonalSites
        Include OneDrive for Business (-my.sharepoint.com) sites. Off by
        default - the brief scopes this to SharePoint sites.
    #>
    [CmdletBinding()]
    param(
        [string]   $ScopePath,
        [string[]] $SiteUrl,
        [switch]   $IncludePersonalSites
    )

    $sites = New-Object System.Collections.Generic.List[object]
    $seen  = New-Object 'System.Collections.Generic.HashSet[string]'

    $explicitUrls = New-Object System.Collections.Generic.List[string]

    if ($SiteUrl) {
        foreach ($url in $SiteUrl) {
            if (-not [string]::IsNullOrWhiteSpace($url)) { $explicitUrls.Add($url.Trim()) }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($ScopePath)) {
        if (-not (Test-Path -LiteralPath $ScopePath)) {
            Write-PubLog -Level Error -Message ('Scope file not found: {0}' -f $ScopePath)
            return @()
        }

        if ($ScopePath -like '*.csv') {
            foreach ($row in (Import-Csv -LiteralPath $ScopePath)) {
                if ($row.PSObject.Properties['SiteUrl'] -and $row.SiteUrl) { $explicitUrls.Add(([string] $row.SiteUrl).Trim()) }
            }
        } else {
            foreach ($line in (Get-Content -LiteralPath $ScopePath)) {
                $trimmed = $line.Trim()
                if ($trimmed -and -not $trimmed.StartsWith('#')) { $explicitUrls.Add($trimmed) }
            }
        }
    }

    if ($explicitUrls.Count -gt 0) {
        Write-PubLog -Level Info -Message ('Resolving {0} scoped site URL(s)...' -f $explicitUrls.Count)

        foreach ($url in ($explicitUrls | Select-Object -Unique)) {
            try {
                $parsed   = [uri] $url.TrimEnd('/')
                $sitePath = $parsed.AbsolutePath.TrimEnd('/')
                $graphUri = "sites/{0}:{1}" -f $parsed.Host, $sitePath
                if ([string]::IsNullOrWhiteSpace($sitePath)) { $graphUri = "sites/{0}" -f $parsed.Host }

                $site = Invoke-PubGraph -Uri ("{0}?`$select=id,webUrl,displayName,name" -f $graphUri) -Method GET
                if ($site -and $seen.Add($site.id)) { $sites.Add($site) }
            } catch {
                Write-PubLog -Level Error -Message ('Could not resolve site {0}: {1}' -f $url, (Get-PubGraphErrorMessage -ErrorRecord $_))
            }
        }
    } else {
        Write-PubLog -Level Info -Message 'Enumerating every site the app registration can see...'

        $allSites = @()
        try {
            $allSites = Get-PubGraphAll -Uri 'sites/getAllSites'
        } catch {
            Write-PubLog -Level Debug -Message ('getAllSites unavailable ({0}); falling back to site search.' -f (Get-PubGraphErrorMessage -ErrorRecord $_))
        }

        if (-not $allSites -or $allSites.Count -eq 0) {
            try {
                $allSites = Get-PubGraphAll -Uri 'sites?search=*&$select=id,webUrl,displayName,name'
            } catch {
                Write-PubLog -Level Error -Message ('Could not enumerate sites: {0}' -f (Get-PubGraphErrorMessage -ErrorRecord $_))
                return @()
            }
        }

        foreach ($site in $allSites) {
            if (-not $site.PSObject.Properties['webUrl'] -or -not $site.webUrl) { continue }
            if (-not $IncludePersonalSites -and $site.webUrl -match '-my\.sharepoint\.com') { continue }
            if ($seen.Add($site.id)) { $sites.Add($site) }
        }
    }

    # Pull in subsites, which site search does not always return.
    $topLevel = @($sites)
    foreach ($site in $topLevel) {
        try {
            $subSites = Get-PubGraphAll -Uri ("sites/{0}/sites?`$select=id,webUrl,displayName,name" -f $site.id)
            foreach ($subSite in $subSites) {
                if (-not $subSite.PSObject.Properties['webUrl'] -or -not $subSite.webUrl) { continue }
                if ($seen.Add($subSite.id)) { $sites.Add($subSite) }
            }
        } catch {
            Write-PubLog -Level Debug -Message ('No subsites read for {0}: {1}' -f $site.webUrl, (Get-PubGraphErrorMessage -ErrorRecord $_))
        }
    }

    Write-PubLog -Level Success -Message ('{0} site(s) in scope.' -f $sites.Count)
    return $sites.ToArray()
}

function Get-PubDocumentLibrary {
    <#
    .SYNOPSIS
        Returns the crawlable document libraries of one site.

    .DESCRIPTION
        Uses /lists rather than /drives because only the list gives the hidden
        flag and the template, which is how system libraries are filtered out.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Site
    )

    $libraries = New-Object System.Collections.Generic.List[object]

    try {
        $lists = Get-PubGraphAll -Uri ("sites/{0}/lists?`$expand=drive&`$select=id,name,displayName,list" -f $Site.id)
    } catch {
        Write-PubLog -Level Warn -Message ('Could not list libraries for {0}: {1}' -f $Site.webUrl, (Get-PubGraphErrorMessage -ErrorRecord $_))
        return @()
    }

    foreach ($list in $lists) {
        if (-not $list.PSObject.Properties['list'] -or -not $list.list) { continue }
        if ($list.list.template -ne 'documentLibrary') { continue }
        if ($list.list.PSObject.Properties['hidden'] -and $list.list.hidden) { continue }
        if (-not $list.PSObject.Properties['drive'] -or -not $list.drive) { continue }

        $displayName = $list.displayName
        if ([string]::IsNullOrWhiteSpace($displayName)) { $displayName = $list.name }
        if ($script:SystemLibraries -contains $displayName) { continue }

        $libraries.Add([pscustomobject] @{
            DriveId     = $list.drive.id
            DisplayName = $displayName
            ListId      = $list.id
        })
    }

    return $libraries.ToArray()
}

function Get-PubDriveFolderPath {
    <#
    .SYNOPSIS
        Converts a driveItem parentReference into a library-relative folder path.

    .DESCRIPTION
        Graph reports parentReference.path as '/drives/{id}/root:/Folder/Sub'.
        The stored FolderPath is the part after 'root:' - '/' for the library
        root - which is what the upload phase needs to put the PDF back in the
        right folder.
    #>
    [CmdletBinding()]
    param($ParentReference)

    if (-not $ParentReference -or -not $ParentReference.PSObject.Properties['path'] -or -not $ParentReference.path) {
        return '/'
    }

    $path = [string] $ParentReference.path
    $index = $path.IndexOf('root:')
    if ($index -ge 0) { $path = $path.Substring($index + 5) }

    if ([string]::IsNullOrWhiteSpace($path)) { return '/' }

    try { $path = [uri]::UnescapeDataString($path) } catch { }
    if (-not $path.StartsWith('/')) { $path = '/' + $path }

    return $path
}

function Get-PubFileInLibrary {
    <#
    .SYNOPSIS
        Walks one document library and returns every .pub file it holds.

    .DESCRIPTION
        Iterative (stack-based) rather than recursive so a deep folder tree
        cannot blow the call stack. Folder-level failures are logged and
        skipped, never fatal.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Site,
        [Parameter(Mandatory)] $Library,
        [scriptblock] $OnFileFound
    )

    $found  = New-Object System.Collections.Generic.List[object]
    $select = 'id,name,size,webUrl,lastModifiedDateTime,lastModifiedBy,parentReference,file,folder,sharepointIds'

    $folderStack = New-Object System.Collections.Stack
    $folderStack.Push('root')

    while ($folderStack.Count -gt 0) {
        $folderId = $folderStack.Pop()

        $uri = "drives/{0}/items/{1}/children?`$select={2}&`$top=200" -f $Library.DriveId, $folderId, $select
        if ($folderId -eq 'root') {
            $uri = "drives/{0}/root/children?`$select={1}&`$top=200" -f $Library.DriveId, $select
        }

        $children = @()
        try {
            $children = Get-PubGraphAll -Uri $uri
        } catch {
            Write-PubLog -Level Warn -Message ('Skipped a folder in {0} / {1}: {2}' -f $Site.webUrl, $Library.DisplayName, (Get-PubGraphErrorMessage -ErrorRecord $_))
            continue
        }

        foreach ($item in $children) {
            if ($item.PSObject.Properties['folder'] -and $item.folder) {
                $folderStack.Push($item.id)
                continue
            }

            if (-not $item.PSObject.Properties['file'] -or -not $item.file) { continue }
            if ([System.IO.Path]::GetExtension($item.name) -ne '.pub') { continue }

            $modifiedBy = ''
            if ($item.PSObject.Properties['lastModifiedBy'] -and $item.lastModifiedBy -and $item.lastModifiedBy.PSObject.Properties['user'] -and $item.lastModifiedBy.user) {
                $user = $item.lastModifiedBy.user
                if ($user.PSObject.Properties['email'] -and $user.email) { $modifiedBy = $user.email }
                elseif ($user.PSObject.Properties['displayName'] -and $user.displayName) { $modifiedBy = $user.displayName }
            }

            $uniqueId = $item.id
            if ($item.PSObject.Properties['sharepointIds'] -and $item.sharepointIds -and $item.sharepointIds.PSObject.Properties['listItemUniqueId'] -and $item.sharepointIds.listItemUniqueId) {
                $uniqueId = $item.sharepointIds.listItemUniqueId
            }

            $sizeKb = 0
            if ($item.PSObject.Properties['size'] -and $item.size) {
                $sizeKb = [math]::Round(([double] $item.size) / 1KB, 1)
            }

            $parentItemId = ''
            if ($item.PSObject.Properties['parentReference'] -and $item.parentReference -and $item.parentReference.PSObject.Properties['id']) {
                $parentItemId = $item.parentReference.id
            }

            $row = New-PubInventoryRow -Values @{
                SiteUrl      = $Site.webUrl
                LibraryName  = $Library.DisplayName
                FolderPath   = (Get-PubDriveFolderPath -ParentReference $item.parentReference)
                FileName     = $item.name
                FileSizeKB   = $sizeKb
                LastModified = $item.lastModifiedDateTime
                ModifiedBy   = $modifiedBy
                UniqueId     = $uniqueId
                Status       = 'Pending'
                Notes        = ''
                SiteId       = $Site.id
                DriveId      = $Library.DriveId
                ItemId       = $item.id
                ParentItemId = $parentItemId
                LocalPath    = ''
                PdfPath      = ''
                LastAction   = (Get-Date).ToString('s')
            }

            $found.Add($row)
            if ($OnFileFound) { & $OnFileFound $row }
        }
    }

    return $found.ToArray()
}

function Invoke-PubDiscovery {
    <#
    .SYNOPSIS
        Menu option 4 - crawls the tenant (or the scoped sites) for .pub files.

    .PARAMETER ScopePath
        Optional file of site URLs to restrict the scan.

    .PARAMETER SiteUrl
        Optional explicit site URLs to restrict the scan.

    .OUTPUTS
        The inventory rows found, ready for Export-PubInventory.
    #>
    [CmdletBinding()]
    param(
        [string]   $ScopePath,
        [string[]] $SiteUrl,
        [switch]   $IncludePersonalSites,
        $Config
    )

    if (-not $Config) { $Config = Get-PubConfig }

    if (-not (Connect-PubGraphApp -Config $Config)) { return @() }

    $started = Get-Date
    Write-PubLog -Level Info -Message '=== Phase 1: Discovery ==='

    $sites = Get-PubSiteList -ScopePath $ScopePath -SiteUrl $SiteUrl -IncludePersonalSites:$IncludePersonalSites
    if (-not $sites -or $sites.Count -eq 0) {
        Write-PubLog -Level Warn -Message 'No sites in scope - nothing to scan.'
        return @()
    }

    $rows            = New-Object System.Collections.Generic.List[object]
    $siteIndex       = 0
    $sitesFailed     = 0
    $librariesCrawled = 0

    foreach ($site in $sites) {
        $siteIndex++

        $percent = [int] (($siteIndex / $sites.Count) * 100)
        Write-Progress -Activity 'Scanning SharePoint for .pub files' `
                       -Status ('Site {0} of {1} - {2} found so far' -f $siteIndex, $sites.Count, $rows.Count) `
                       -CurrentOperation $site.webUrl `
                       -PercentComplete $percent

        try {
            $libraries = Get-PubDocumentLibrary -Site $site
        } catch {
            $sitesFailed++
            Write-PubLog -Level Error -Message ('Site failed: {0} - {1}' -f $site.webUrl, (Get-PubGraphErrorMessage -ErrorRecord $_))
            continue
        }

        foreach ($library in $libraries) {
            $librariesCrawled++
            try {
                $files = Get-PubFileInLibrary -Site $site -Library $library
                foreach ($file in $files) {
                    $rows.Add($file)
                    Write-PubFileResult -Outcome Found -FileName $file.FileName -Detail ('{0} / {1}{2}' -f $site.webUrl, $library.DisplayName, $file.FolderPath)
                }
            } catch {
                Write-PubLog -Level Error -Message ('Library failed: {0} / {1} - {2}' -f $site.webUrl, $library.DisplayName, (Get-PubGraphErrorMessage -ErrorRecord $_))
            }
        }
    }

    Write-Progress -Activity 'Scanning SharePoint for .pub files' -Completed

    $elapsed = (Get-Date) - $started
    Write-PubPhaseSummary -Phase 'Discovery' `
                          -Attempted $sites.Count `
                          -Succeeded ($sites.Count - $sitesFailed) `
                          -Failed $sitesFailed `
                          -ExtraLines @(
                              ('Libraries crawled : {0}' -f $librariesCrawled)
                              ('Publisher files   : {0}' -f $rows.Count)
                              ('Elapsed           : {0:hh\:mm\:ss}' -f $elapsed)
                          )

    return $rows.ToArray()
}

function Export-PubInventory {
    <#
    .SYNOPSIS
        Writes the inventory to a timestamped CSV and records it in config.

    .PARAMETER Path
        Explicit output path. Omit for
        <working>\inventory\PublisherFileInventory_yyyy-MM-dd_HHmm.csv.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        $Rows,

        [string] $Path,
        $Config,
        [switch] $NoConfigUpdate
    )

    if (-not $Config) { $Config = Get-PubConfig }

    if ([string]::IsNullOrWhiteSpace($Path)) {
        $folder = Get-PubWorkingFolder -SubFolder 'inventory' -Config $Config
        $Path   = Join-Path $folder ('PublisherFileInventory_{0}.csv' -f (Get-Date -Format 'yyyy-MM-dd_HHmm'))
    }

    $columns = Get-PubInventoryColumns

    try {
        @($Rows) | Select-Object $columns | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
    } catch {
        Write-PubLog -Level Error -Message ('Could not write the CSV to {0}: {1}' -f $Path, $_.Exception.Message)
        return $null
    }

    Write-PubLog -Level Success -Message ('Inventory written: {0} ({1} rows)' -f $Path, @($Rows).Count)

    if (-not $NoConfigUpdate) {
        $Config['LastInventoryCsv'] = $Path
        Save-PubConfig -Config $Config | Out-Null
    }

    return $Path
}

function Import-PubInventory {
    <#
    .SYNOPSIS
        Loads an inventory CSV, filling in any columns the operator removed.

    .DESCRIPTION
        Trimming the CSV down in Excel is an expected workflow (brief section
        5.1), so a missing optional column is repaired rather than rejected.
        A CSV missing the columns needed to locate the file in SharePoint is
        rejected with a clear message.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-PubLog -Level Error -Message ('Inventory CSV not found: {0}' -f $Path)
        return @()
    }

    try {
        $raw = @(Import-Csv -LiteralPath $Path)
    } catch {
        Write-PubLog -Level Error -Message ('Could not read {0}: {1}' -f $Path, $_.Exception.Message)
        return @()
    }

    if ($raw.Count -eq 0) {
        Write-PubLog -Level Warn -Message ('{0} contains no rows.' -f $Path)
        return @()
    }

    $present  = $raw[0].PSObject.Properties.Name
    $required = @('FileName', 'DriveId', 'ItemId')
    $missing  = @()
    foreach ($column in $required) {
        if ($present -notcontains $column) { $missing += $column }
    }

    if ($missing.Count -gt 0) {
        Write-PubLog -Level Error -Message ('{0} is missing required column(s): {1}. Re-run discovery to produce a usable inventory.' -f $Path, ($missing -join ', '))
        return @()
    }

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($item in $raw) {
        $values = @{}
        foreach ($property in $item.PSObject.Properties) {
            $values[$property.Name] = $property.Value
        }
        if (-not $values.ContainsKey('Status') -or [string]::IsNullOrWhiteSpace([string] $values['Status'])) {
            $values['Status'] = 'Pending'
        }
        $rows.Add((New-PubInventoryRow -Values $values))
    }

    Write-PubLog -Level Success -Message ('Loaded {0} row(s) from {1}' -f $rows.Count, $Path)
    return $rows.ToArray()
}

function Set-PubInventoryStatus {
    <#
    .SYNOPSIS
        Updates one row's Status/Notes and stamps LastAction.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Row,

        [Parameter(Mandatory)]
        [ValidateSet('Pending', 'Downloaded', 'Converted', 'Uploaded', 'Failed', 'Skipped')]
        [string] $Status,

        [string] $Notes
    )

    $Row.Status     = $Status
    $Row.LastAction = (Get-Date).ToString('s')

    if ($PSBoundParameters.ContainsKey('Notes')) {
        $clean = ''
        if ($Notes) { $clean = ($Notes -replace '[\r\n]+', ' ').Trim() }
        $Row.Notes = $clean
    }

    return $Row
}

function Get-PubInventoryStatistic {
    <#
    .SYNOPSIS
        Counts rows by Status, for the menu header and phase summaries.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        $Rows
    )

    $statistics = [ordered] @{
        Total      = @($Rows).Count
        Pending    = 0
        Downloaded = 0
        Converted  = 0
        Uploaded   = 0
        Failed     = 0
        Skipped    = 0
    }

    foreach ($row in @($Rows)) {
        $status = [string] $row.Status
        if ([string]::IsNullOrWhiteSpace($status)) { $status = 'Pending' }
        if ($statistics.Contains($status)) { $statistics[$status] = $statistics[$status] + 1 }
    }

    return $statistics
}

Export-ModuleMember -Function @(
    'Get-PubInventoryColumns'
    'New-PubInventoryRow'
    'Get-PubSiteList'
    'Get-PubDocumentLibrary'
    'Get-PubDriveFolderPath'
    'Get-PubFileInLibrary'
    'Invoke-PubDiscovery'
    'Export-PubInventory'
    'Import-PubInventory'
    'Set-PubInventoryStatus'
    'Get-PubInventoryStatistic'
)
