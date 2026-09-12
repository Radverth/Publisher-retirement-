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

foreach ($dependency in @('Logging', 'Config', 'Graph', 'PnP')) {
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

function Get-PubScopeUrl {
    <#
    .SYNOPSIS
        Reads a list of site URLs from a text or CSV file.

    .DESCRIPTION
        Accepts the SharePoint admin centre's own export as-is: Active sites ->
        Export to CSV names its column 'URL', not 'SiteUrl'. That export comes
        from the SharePoint tenant store rather than the search index, so it is
        the one list guaranteed to contain every site collection - which makes
        it the reliable way to scope a complete tenant crawl.

        A plain text file is read one URL per line; lines starting with # are
        ignored.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Path
    )

    $urls = New-Object System.Collections.Generic.List[string]

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-PubLog -Level Error -Message ('Scope file not found: {0}' -f $Path)
        return @()
    }

    if ($Path -notlike '*.csv') {
        foreach ($line in (Get-Content -LiteralPath $Path)) {
            $trimmed = $line.Trim()
            if ($trimmed -and -not $trimmed.StartsWith('#')) { $urls.Add($trimmed) }
        }
        return $urls.ToArray()
    }

    $rows = @(Import-Csv -LiteralPath $Path)
    if ($rows.Count -eq 0) {
        Write-PubLog -Level Warn -Message ('{0} contains no rows.' -f $Path)
        return @()
    }

    # In preference order: this tool's own export, then the SharePoint admin
    # centre export, then the obvious hand-rolled variants.
    $candidates = @('SiteUrl', 'Site Url', 'Site URL', 'URL', 'Url', 'WebUrl', 'Web Url')
    $columnName = $null

    foreach ($candidate in $candidates) {
        $match = $rows[0].PSObject.Properties | Where-Object { $_.Name -eq $candidate } | Select-Object -First 1
        if ($match) { $columnName = $match.Name; break }
    }

    if (-not $columnName) {
        $available = ($rows[0].PSObject.Properties.Name -join ', ')
        Write-PubLog -Level Error -Message ('{0} has no site URL column. Looked for: {1}. Columns present: {2}' -f $Path, ($candidates -join ', '), $available)
        return @()
    }

    Write-PubLog -Level Info -Message ('Reading site URLs from the "{0}" column of {1}.' -f $columnName, (Split-Path -Leaf $Path))

    foreach ($row in $rows) {
        $value = [string] $row.$columnName
        if (-not [string]::IsNullOrWhiteSpace($value)) { $urls.Add($value.Trim()) }
    }

    return $urls.ToArray()
}

function Test-PubPnPEnumerationEnabled {
    <#
    .SYNOPSIS
        True when discovery should use the SharePoint tenant admin site list.

    .DESCRIPTION
        Driven by the EnumerationMethod setting:
          Auto  - use PnP when it is installed and the app can sign in to the
                  admin site, otherwise fall back to Graph. (default)
          PnP   - insist on PnP; a failure is reported rather than hidden.
          Graph - never use PnP.
    #>
    [CmdletBinding()]
    param($Config)

    if (-not $Config) { $Config = Get-PubConfig }

    $method = [string] $Config['EnumerationMethod']
    if ([string]::IsNullOrWhiteSpace($method)) { $method = 'Auto' }

    if ($method -eq 'Graph') { return $false }

    if (-not (Test-PubPnPAvailable -Quiet)) {
        if ($method -eq 'PnP') {
            Write-PubLog -Level Warn -Message 'EnumerationMethod is PnP but PnP PowerShell is not usable on this host - falling back to Graph.'
        }
        return $false
    }

    if (-not (Connect-PubPnPAdmin -Config $Config)) {
        if ($method -eq 'PnP') {
            Write-PubLog -Level Warn -Message 'EnumerationMethod is PnP but the tenant admin connection failed - falling back to Graph.'
        }
        return $false
    }

    return $true
}

function Get-PubAllTenantSite {
    <#
    .SYNOPSIS
        Enumerates every site in the tenant, trying the reliable routes first.

    .DESCRIPTION
        Access is not the constraint here - an app-only token with
        Sites.Read.All can read any site regardless of who owns it or belongs
        to it. The constraint is ENUMERATION: knowing a site exists in order to
        crawl it. The three routes, best first:

          1. v1.0 /sites/getAllSites  - tenant store, complete. Not yet
             available in every tenant.
          2. beta /sites/getAllSites  - same data, beta endpoint.
          3. v1.0 /sites?search=*     - backed by the SEARCH INDEX, so it can
             miss sites excluded from indexing, sites indexed too recently, and
             Teams private-channel sites.

        The route actually used is logged and returned, because a short site
        count from route 3 is a coverage problem the operator needs to see, not
        a quiet default. For a guaranteed-complete list, export Active sites
        from the SharePoint admin centre and scope the scan to that CSV.
    #>
    [CmdletBinding()]
    param()

    $routes = @(
        [pscustomobject] @{ Name = 'v1.0 getAllSites'; Uri = 'sites/getAllSites';                                  Complete = $true  }
        [pscustomobject] @{ Name = 'beta getAllSites'; Uri = 'https://graph.microsoft.com/beta/sites/getAllSites'; Complete = $true  }
        [pscustomobject] @{ Name = 'site search';      Uri = 'sites?search=*&$select=id,webUrl,displayName,name';  Complete = $false }
    )

    foreach ($route in $routes) {
        try {
            $result = Get-PubGraphAll -Uri $route.Uri
            if ($result -and @($result).Count -gt 0) {
                return [pscustomobject] @{
                    Sites    = @($result)
                    Route    = $route.Name
                    Complete = $route.Complete
                }
            }
            Write-PubLog -Level Debug -Message ('{0} returned nothing - trying the next route.' -f $route.Name)
        } catch {
            Write-PubLog -Level Debug -Message ('{0} unavailable: {1}' -f $route.Name, (Get-PubGraphErrorMessage -ErrorRecord $_))
        }
    }

    return [pscustomobject] @{ Sites = @(); Route = 'none'; Complete = $false }
}

function Get-PubSiteList {
    <#
    .SYNOPSIS
        Returns the sites to crawl.

    .DESCRIPTION
        With app-only authentication there is no user context, so site
        ownership and membership are irrelevant - every site the app has been
        consented to is readable. What varies is whether a site can be
        enumerated in the first place; see Get-PubAllTenantSite.

    .PARAMETER ScopePath
        Optional path to a text file of site URLs (one per line) or a CSV -
        including the SharePoint admin centre's Active sites export, unedited.

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
        [switch]   $IncludePersonalSites,
        $Config
    )

    if (-not $Config) { $Config = Get-PubConfig }

    $sites = New-Object System.Collections.Generic.List[object]
    $seen  = New-Object 'System.Collections.Generic.HashSet[string]'

    $explicitUrls = New-Object System.Collections.Generic.List[string]

    if ($SiteUrl) {
        foreach ($url in $SiteUrl) {
            if (-not [string]::IsNullOrWhiteSpace($url)) { $explicitUrls.Add($url.Trim()) }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($ScopePath)) {
        foreach ($url in (Get-PubScopeUrl -Path $ScopePath)) { $explicitUrls.Add($url) }
    }

    # The PnP tenant-admin route returns URLs rather than Graph site objects, so
    # it feeds the same resolution path as an operator-supplied site list.
    $enumerationRoute = ''
    $routeComplete    = $false

    if ($explicitUrls.Count -eq 0 -and (Test-PubPnPEnumerationEnabled -Config $Config)) {
        # ConvertTo-PubArray, not @( ): a function returning an empty array
        # yields $null, and @($null) counts as one item - which would queue a
        # null site URL and fail to resolve it.
        $pnpUrls = ConvertTo-PubArray (Get-PubPnPSiteUrl -IncludePersonalSites:$IncludePersonalSites -Config $Config)
        if ($pnpUrls.Count -gt 0) {
            foreach ($url in $pnpUrls) { $explicitUrls.Add($url) }
            $enumerationRoute = 'SharePoint tenant admin (PnP)'
            $routeComplete    = $true
        } else {
            Write-PubLog -Level Warn -Message 'The PnP tenant-admin route returned nothing - falling back to Graph enumeration.'
        }
    }

    if ($explicitUrls.Count -gt 0) {
        $unique = @($explicitUrls | Select-Object -Unique)
        Write-PubLog -Level Info -Message ('Resolving {0} site URL(s) to Graph sites...' -f $unique.Count)

        $resolved = 0
        $failed   = 0

        foreach ($url in $unique) {
            $resolved++
            if ($unique.Count -gt 25) {
                Write-Progress -Activity 'Resolving sites' `
                               -Status ('{0} of {1}' -f $resolved, $unique.Count) `
                               -CurrentOperation $url `
                               -PercentComplete ([int] (($resolved / $unique.Count) * 100))
            }

            if (-not $IncludePersonalSites -and $url -match '-my\.sharepoint\.com') { continue }

            try {
                $parsed   = [uri] $url.TrimEnd('/')
                $sitePath = $parsed.AbsolutePath.TrimEnd('/')
                $graphUri = "sites/{0}:{1}" -f $parsed.Host, $sitePath
                if ([string]::IsNullOrWhiteSpace($sitePath)) { $graphUri = "sites/{0}" -f $parsed.Host }

                $resolvedSite = Invoke-PubGraph -Uri ("{0}?`$select=id,webUrl,displayName,name" -f $graphUri) -Method GET
                if ($resolvedSite -and $seen.Add($resolvedSite.id)) { $sites.Add($resolvedSite) }
            } catch {
                $failed++
                Write-PubLog -Level Error -Message ('Could not resolve site {0}: {1}' -f $url, (Get-PubGraphErrorMessage -ErrorRecord $_))
            }
        }

        if ($unique.Count -gt 25) { Write-Progress -Activity 'Resolving sites' -Completed }

        if ($enumerationRoute) {
            Write-PubLog -Level Info -Message ('Enumeration route: {0} - {1} site(s) resolved, {2} unresolvable.' -f $enumerationRoute, $sites.Count, $failed)
        }
        if ($failed -gt 0) {
            Write-PubLog -Level Warn -Message 'Unresolvable sites are usually locked, deleted-but-not-purged, or blocked by a Restricted Access Control policy. They are listed above and skipped.'
        }
    } else {
        Write-PubLog -Level Info -Message 'Enumerating every site the app registration can see...'

        $enumeration = Get-PubAllTenantSite

        if (@($enumeration.Sites).Count -eq 0) {
            Write-PubLog -Level Error -Message 'Could not enumerate any sites. Check that admin consent has been granted for Sites.Read.All, then re-test with menu option 3.'
            return @()
        }

        Write-PubLog -Level Info -Message ('Enumeration route: {0} - {1} site(s) returned.' -f $enumeration.Route, @($enumeration.Sites).Count)

        if (-not $enumeration.Complete) {
            Write-PubLog -Level Warn -Message 'This route reads the SEARCH INDEX, so it can miss sites excluded from indexing, very new sites, and Teams private-channel sites.'
            Write-PubLog -Level Warn -Message 'For a guaranteed-complete crawl: SharePoint admin centre -> Active sites -> Export to CSV, then re-run the scan scoped to that CSV.'
        }

        $personalSkipped = 0
        foreach ($enumeratedSite in $enumeration.Sites) {
            if (-not $enumeratedSite.PSObject.Properties['webUrl'] -or -not $enumeratedSite.webUrl) { continue }

            if ($enumeratedSite.webUrl -match '-my\.sharepoint\.com') {
                if (-not $IncludePersonalSites) { $personalSkipped++; continue }
            }

            if ($seen.Add($enumeratedSite.id)) { $sites.Add($enumeratedSite) }
        }

        if ($personalSkipped -gt 0) {
            Write-PubLog -Level Info -Message ('{0} OneDrive personal site(s) skipped. Choose "include OneDrive" at the scan scope prompt to crawl them too.' -f $personalSkipped)
        }
    }

    # Pull in subsites, which site search does not always return.
    #
    # NOTE: the loop variable here is deliberately NOT $site. PowerShell binds a
    # variable's type per scope when it compiles a function, and $site is
    # already used for a hand-assigned Graph result and for another foreach
    # above. Reusing it a third time makes the compiler throw "Argument types do
    # not match" before the loop body ever runs.
    $topLevel = $sites.ToArray()
    foreach ($parentSite in $topLevel) {
        try {
            $subSites = Get-PubGraphAll -Uri ("sites/{0}/sites?`$select=id,webUrl,displayName,name" -f $parentSite.id)
            foreach ($childSite in $subSites) {
                if (-not $childSite.PSObject.Properties['webUrl'] -or -not $childSite.webUrl) { continue }
                if ($seen.Add($childSite.id)) { $sites.Add($childSite) }
            }
        } catch {
            Write-PubLog -Level Debug -Message ('No subsites read for {0}: {1}' -f $parentSite.webUrl, (Get-PubGraphErrorMessage -ErrorRecord $_))
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
        [scriptblock] $OnFileFound,

        [int]    $ProgressId = 0,
        [int]    $ParentProgressId = 1,
        [string] $ProgressActivity
    )

    $found  = New-Object System.Collections.Generic.List[object]
    $select = 'id,name,size,webUrl,lastModifiedDateTime,lastModifiedBy,parentReference,file,folder,sharepointIds'

    $folderStack = New-Object System.Collections.Stack
    $folderStack.Push('root')

    # Counters live in a hashtable, not in plain variables: the page callback
    # below is invoked with & from Get-PubGraphAll, which gives it its own
    # scope, and assigning to an outer scalar from there would silently update
    # a local copy while the real counter stayed at zero. Mutating a member of
    # a shared hashtable is the same object in both scopes.
    $counters = @{ Folders = 0; Items = 0; Path = '/' }

    # A big library is hundreds of sequential requests. Without an update per
    # page the whole site looks frozen, which is exactly how this reads to an
    # operator watching the console.
    $reportProgress = {
        param([string] $Operation)

        if ($ProgressId -le 0) { return }

        Write-Progress -Id $ProgressId -ParentId $ParentProgressId `
                       -Activity $ProgressActivity `
                       -Status ('{0} folder(s) read | {1} item(s) seen | {2} .pub found | {3} folder(s) queued' -f $counters.Folders, $counters.Items, $found.Count, $folderStack.Count) `
                       -CurrentOperation $Operation
    }

    & $reportProgress 'starting'

    while ($folderStack.Count -gt 0) {
        $folderId = $folderStack.Pop()

        $uri = "drives/{0}/items/{1}/children?`$select={2}&`$top=200" -f $Library.DriveId, $folderId, $select
        if ($folderId -eq 'root') {
            $uri = "drives/{0}/root/children?`$select={1}&`$top=200" -f $Library.DriveId, $select
        }

        $onPage = {
            param($PageItems, $PageNumber, $ItemsSoFar)

            $counters.Items = $counters.Items + (ConvertTo-PubArray $PageItems).Count

            $suffix = ''
            if ($PageNumber -gt 1) { $suffix = (' (page {0})' -f $PageNumber) }
            & $reportProgress ('{0}{1}' -f $counters.Path, $suffix)
        }

        $children = @()
        try {
            $children = Get-PubGraphAll -Uri $uri -OnPage $onPage
        } catch {
            Write-PubLog -Level Warn -Message ('Skipped a folder in {0} / {1}: {2}' -f $Site.webUrl, $Library.DisplayName, (Get-PubGraphErrorMessage -ErrorRecord $_))
            continue
        }

        $counters.Folders = $counters.Folders + 1

        foreach ($item in $children) {
            if ($item.PSObject.Properties['folder'] -and $item.folder) {
                $folderStack.Push($item.id)
                continue
            }

            if ($item.PSObject.Properties['parentReference'] -and $item.parentReference) {
                $counters.Path = Get-PubDriveFolderPath -ParentReference $item.parentReference
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

        & $reportProgress $counters.Path
    }

    if ($ProgressId -gt 0) {
        Write-Progress -Id $ProgressId -ParentId $ParentProgressId -Activity $ProgressActivity -Completed
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

    .PARAMETER AutoSaveEvery
        Write a partial CSV every N sites so a long crawl that is interrupted
        keeps what it has already found. 0 disables it.

    .OUTPUTS
        The inventory rows found, ready for Export-PubInventory.
    #>
    [CmdletBinding()]
    param(
        [string]   $ScopePath,
        [string[]] $SiteUrl,
        [switch]   $IncludePersonalSites,
        $Config,

        [int] $AutoSaveEvery = 10
    )

    if (-not $Config) { $Config = Get-PubConfig }

    if (-not (Connect-PubGraphApp -Config $Config)) { return @() }

    $started = Get-Date
    Write-PubLog -Level Info -Message '=== Phase 1: Discovery ==='

    $sites = Get-PubSiteList -ScopePath $ScopePath -SiteUrl $SiteUrl -IncludePersonalSites:$IncludePersonalSites -Config $Config
    if (-not $sites -or $sites.Count -eq 0) {
        Write-PubLog -Level Warn -Message 'No sites in scope - nothing to scan.'
        return @()
    }

    $rows            = New-Object System.Collections.Generic.List[object]
    $siteIndex       = 0
    $sitesFailed     = 0
    $librariesCrawled = 0
    $partialCsvPath  = ''

    foreach ($site in $sites) {
        $siteIndex++
        $siteStarted = Get-Date

        $elapsed = (Get-Date) - $started
        $percent = [int] (($siteIndex / $sites.Count) * 100)

        # An estimate is only meaningful once a couple of sites have finished,
        # and sites vary wildly in size, so it is explicitly a rough figure.
        $remainingText = ''
        if ($siteIndex -gt 2) {
            $perSite   = $elapsed.TotalSeconds / ($siteIndex - 1)
            $remaining = [timespan]::FromSeconds($perSite * ($sites.Count - $siteIndex + 1))
            $remainingText = ' | about {0} left' -f (Format-PubDuration -Duration $remaining)
        }

        Write-Progress -Id 1 -Activity 'Scanning SharePoint for .pub files' `
                       -Status ('Site {0} of {1} | {2} .pub found | {3} elapsed{4}' -f $siteIndex, $sites.Count, $rows.Count, (Format-PubDuration -Duration $elapsed), $remainingText) `
                       -CurrentOperation $site.webUrl `
                       -PercentComplete $percent

        try {
            $libraries = Get-PubDocumentLibrary -Site $site
        } catch {
            $sitesFailed++
            Write-PubLog -Level Error -Message ('Site failed: {0} - {1}' -f $site.webUrl, (Get-PubGraphErrorMessage -ErrorRecord $_))
            continue
        }

        $libraries    = ConvertTo-PubArray $libraries
        $libraryIndex = 0

        foreach ($library in $libraries) {
            $librariesCrawled++
            $libraryIndex++

            $activity = 'Library {0} of {1}: {2}' -f $libraryIndex, $libraries.Count, $library.DisplayName

            try {
                $files = Get-PubFileInLibrary -Site $site -Library $library `
                                              -ProgressId 2 -ParentProgressId 1 -ProgressActivity $activity
                foreach ($file in $files) {
                    $rows.Add($file)
                    Write-PubFileResult -Outcome Found -FileName $file.FileName -Detail ('{0} / {1}{2}' -f $site.webUrl, $library.DisplayName, $file.FolderPath)
                }
            } catch {
                Write-PubLog -Level Error -Message ('Library failed: {0} / {1} - {2}' -f $site.webUrl, $library.DisplayName, (Get-PubGraphErrorMessage -ErrorRecord $_))
            }
        }

        # A site big enough to look like a hang gets a line in the log saying
        # how long it actually took, so slow sites can be identified afterwards.
        $siteDuration = (Get-Date) - $siteStarted
        if ($siteDuration.TotalSeconds -ge 60) {
            Write-PubLog -Level Info -Message ('Site {0} of {1} took {2}: {3} ({4} librar(y/ies), {5} .pub found so far)' -f `
                $siteIndex, $sites.Count, (Format-PubDuration -Duration $siteDuration), $site.webUrl, $libraries.Count, $rows.Count)
        }

        # Save what has been found so far, so a long crawl that is interrupted
        # is not lost - the partial CSV can be loaded with menu option 6.
        if ($AutoSaveEvery -gt 0 -and $rows.Count -gt 0 -and ($siteIndex % $AutoSaveEvery) -eq 0 -and $siteIndex -lt $sites.Count) {
            if ([string]::IsNullOrWhiteSpace($partialCsvPath)) {
                $inventoryFolder = Get-PubWorkingFolder -SubFolder 'inventory' -Config $Config
                $partialCsvPath  = Join-Path $inventoryFolder ('PublisherFileInventory_{0}_partial.csv' -f (Get-Date -Format 'yyyy-MM-dd_HHmm'))
            }

            Export-PubInventory -Rows $rows.ToArray() -Path $partialCsvPath -Config $Config -NoConfigUpdate | Out-Null
            Write-PubLog -Level Info -Message ('Progress saved after {0} site(s): {1}' -f $siteIndex, $partialCsvPath)
        }
    }

    Write-Progress -Id 2 -Activity 'Library' -Completed
    Write-Progress -Id 1 -Activity 'Scanning SharePoint for .pub files' -Completed

    $elapsed = (Get-Date) - $started
    $scopeLabel = 'every site the app can enumerate'
    if ($ScopePath)               { $scopeLabel = 'site list from {0}' -f (Split-Path -Leaf $ScopePath) }
    elseif ($SiteUrl -and @($SiteUrl).Count -gt 0) { $scopeLabel = '{0} site URL(s) supplied' -f @($SiteUrl).Count }

    $oneDriveLabel = 'excluded'
    if ($IncludePersonalSites) { $oneDriveLabel = 'included' }

    Write-PubPhaseSummary -Phase 'Discovery' `
                          -Attempted $sites.Count `
                          -Succeeded ($sites.Count - $sitesFailed) `
                          -Failed $sitesFailed `
                          -ExtraLines @(
                              ('Scope             : {0}' -f $scopeLabel)
                              ('OneDrive sites    : {0}' -f $oneDriveLabel)
                              ('Libraries crawled : {0}' -f $librariesCrawled)
                              ('Publisher files   : {0}' -f $rows.Count)
                              $(if ($partialCsvPath) { 'Partial saves     : {0}' -f $partialCsvPath } else { '' })
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
        (ConvertTo-PubArray $Rows) | Select-Object $columns | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
    } catch {
        Write-PubLog -Level Error -Message ('Could not write the CSV to {0}: {1}' -f $Path, $_.Exception.Message)
        return $null
    }

    Write-PubLog -Level Success -Message ('Inventory written: {0} ({1} rows)' -f $Path, (ConvertTo-PubArray $Rows).Count)

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

    $rows = ConvertTo-PubArray $Rows

    $statistics = [ordered] @{
        Total      = $rows.Count
        Pending    = 0
        Downloaded = 0
        Converted  = 0
        Uploaded   = 0
        Failed     = 0
        Skipped    = 0
    }

    foreach ($row in $rows) {
        $status = [string] $row.Status
        if ([string]::IsNullOrWhiteSpace($status)) { $status = 'Pending' }
        if ($statistics.Contains($status)) { $statistics[$status] = $statistics[$status] + 1 }
    }

    return $statistics
}

Export-ModuleMember -Function @(
    'Get-PubInventoryColumns'
    'New-PubInventoryRow'
    'Get-PubScopeUrl'
    'Test-PubPnPEnumerationEnabled'
    'Get-PubAllTenantSite'
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
