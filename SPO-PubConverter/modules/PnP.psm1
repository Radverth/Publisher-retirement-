<#
.SYNOPSIS
    Optional PnP PowerShell integration: fully automatic app registration and
    complete tenant site enumeration.

.DESCRIPTION
    Everything in this module is optional. The tool runs end to end on the
    Microsoft Graph SDK alone; PnP is used where it removes a manual step that
    Graph cannot:

      * Get-PnPTenantSite reads the SharePoint tenant admin site list, which
        comes from the tenant store rather than the search index. That is the
        complete list of site collections - including sites excluded from
        search indexing, brand new sites, and Teams private-channel sites -
        so no manual "export Active sites to CSV" step is needed.
      * Register-PnPEntraIDApp creates the app registration, generates and
        uploads the certificate, and runs the consent flow in one call.

    THE COST, STATED PLAINLY
    ------------------------
    Reading the tenant admin site list app-only requires the SharePoint
    application permission Sites.FullControl.All (Sites.Manage.All is the
    documented floor). That is full administrative control of every site
    collection in the tenant - a larger grant than the Graph Sites.Read.All
    used elsewhere, and it cannot be combined with Sites.Selected. It is
    requested only when the operator chooses the TenantAdmin scope at setup.

    REQUIREMENTS
    ------------
    Current PnP.PowerShell needs PowerShell 7.4.6 or later; Windows PowerShell
    5.1 is only supported by PnP 1.12.0, which is long unmaintained. When PnP
    is absent or the host is too old, every function here degrades quietly and
    the Graph enumeration routes take over.
#>

Set-StrictMode -Version 2.0

foreach ($dependency in @('Logging', 'Config', 'Graph', 'AppRegistration')) {
    if (-not (Get-Module -Name $dependency)) {
        Import-Module (Join-Path $PSScriptRoot ("{0}.psm1" -f $dependency)) -Force -DisableNameChecking
    }
}

$script:PnPProbeResult = $null
$script:PnPConnected   = $false

function Test-PubPnPAvailable {
    <#
    .SYNOPSIS
        True when PnP PowerShell can be used on this host.

    .PARAMETER Refresh
        Re-probe instead of using the cached result.
    #>
    [CmdletBinding()]
    param(
        [switch] $Refresh,
        [switch] $Quiet
    )

    if ($null -ne $script:PnPProbeResult -and -not $Refresh) { return $script:PnPProbeResult }

    $module = Get-Module -ListAvailable -Name 'PnP.PowerShell' | Sort-Object Version -Descending | Select-Object -First 1

    if (-not $module) {
        if (-not $Quiet) {
            Write-PubLog -Level Info -Message 'PnP PowerShell is not installed - Graph enumeration will be used instead.'
            Write-PubLog -Level Info -Message 'To enable fully automatic site discovery: Install-Module PnP.PowerShell -Scope CurrentUser'
        }
        $script:PnPProbeResult = $false
        return $false
    }

    # PnP 2.x/3.x are PowerShell 7 only. 1.x runs on 5.1 but is unmaintained.
    if ($module.Version.Major -ge 2 -and $PSVersionTable.PSVersion.Major -lt 7) {
        if (-not $Quiet) {
            Write-PubLog -Level Warn -Message ('PnP PowerShell {0} needs PowerShell 7.4.6 or later; this host is {1}. Graph enumeration will be used instead.' -f $module.Version, $PSVersionTable.PSVersion)
        }
        $script:PnPProbeResult = $false
        return $false
    }

    if (-not $Quiet) {
        Write-PubLog -Level Debug -Message ('PnP PowerShell {0} available.' -f $module.Version)
    }

    $script:PnPProbeResult = $true
    return $true
}

function Get-PubTenantAdminUrl {
    <#
    .SYNOPSIS
        Works out the SharePoint tenant admin URL for this tenant.

    .DESCRIPTION
        Uses the explicit SharePointAdminUrl from config when set; otherwise
        derives it from the tenant domain (contoso.onmicrosoft.com ->
        https://contoso-admin.sharepoint.com), falling back to the root site's
        host from Graph.
    #>
    [CmdletBinding()]
    param($Config)

    if (-not $Config) { $Config = Get-PubConfig }

    $configured = [string] $Config['SharePointAdminUrl']
    if (-not [string]::IsNullOrWhiteSpace($configured)) { return $configured.TrimEnd('/') }

    $domain = [string] $Config['TenantDomain']
    if ($domain -match '^([A-Za-z0-9\-]+)\.onmicrosoft\.com$') {
        return ('https://{0}-admin.sharepoint.com' -f $Matches[1])
    }
    if ($domain -match '^([A-Za-z0-9\-]+)\.sharepoint\.com$') {
        return ('https://{0}-admin.sharepoint.com' -f $Matches[1])
    }

    # Last resort: ask Graph for the root site and derive the host from it.
    try {
        if (Connect-PubGraphApp -Config $Config) {
            $rootSite = Invoke-PubGraph -Uri 'sites/root?$select=webUrl' -Method GET
            if ($rootSite -and $rootSite.webUrl -and ([uri] $rootSite.webUrl).Host -match '^([A-Za-z0-9\-]+)\.sharepoint\.com$') {
                return ('https://{0}-admin.sharepoint.com' -f $Matches[1])
            }
        }
    } catch {
        Write-PubLog -Level Debug -Message ('Could not derive the admin URL from Graph: {0}' -f $_.Exception.Message)
    }

    return ''
}

function Connect-PubPnPAdmin {
    <#
    .SYNOPSIS
        Connects PnP app-only to the tenant admin site.

    .DESCRIPTION
        Uses the same bespoke app registration and certificate as the Graph
        connection - PnP does not need, and this tool does not use, the PnP
        Management Shell multi-tenant app.

        Returns $true on success. Never throws: a failure here just means the
        Graph enumeration routes are used instead.
    #>
    [CmdletBinding()]
    param(
        $Config,
        [switch] $Force
    )

    if (-not $Config) { $Config = Get-PubConfig }
    if (-not (Test-PubPnPAvailable -Quiet)) { return $false }

    if ($script:PnPConnected -and -not $Force) { return $true }

    if (-not (Test-PubConfigComplete -Config $Config)) {
        Write-PubLog -Level Warn -Message 'No app registration or certificate configured - cannot connect PnP.'
        return $false
    }

    $adminUrl = Get-PubTenantAdminUrl -Config $Config
    if ([string]::IsNullOrWhiteSpace($adminUrl)) {
        Write-PubLog -Level Warn -Message 'Could not determine the SharePoint admin URL. Set SharePointAdminUrl in config.json.'
        return $false
    }

    $tenant = [string] $Config['TenantDomain']
    if ([string]::IsNullOrWhiteSpace($tenant)) { $tenant = [string] $Config['TenantId'] }

    try {
        Import-Module -Name 'PnP.PowerShell' -ErrorAction Stop -DisableNameChecking

        $parameters = @{
            Url         = $adminUrl
            ClientId    = [string] $Config['AppId']
            Tenant      = $tenant
            ErrorAction = 'Stop'
        }

        if (Test-PubIsWindows) {
            $parameters['Thumbprint'] = [string] $Config['CertificateThumbprint']
        } else {
            $pfxPath = [string] $Config['CertificatePfxPath']
            if ([string]::IsNullOrWhiteSpace($pfxPath) -or -not (Test-Path -LiteralPath $pfxPath)) {
                Write-PubLog -Level Warn -Message 'No .pfx available for PnP app-only sign-in on this host.'
                return $false
            }
            $parameters['CertificatePath'] = $pfxPath

            $password = Get-PubCertificatePassword -Config $Config
            if (-not $password) {
                Write-PubLog -Level Warn -Message 'The .pfx password could not be found, so PnP cannot open the certificate.'
                return $false
            }
            $parameters['CertificatePassword'] = $password
        }

        Connect-PnPOnline @parameters
        $script:PnPConnected = $true

        Write-PubLog -Level Success -Message ('PnP connected to {0}.' -f $adminUrl)
        return $true
    } catch {
        $script:PnPConnected = $false
        $message = $_.Exception.Message

        Write-PubLog -Level Warn -Message ('PnP admin sign-in failed: {0}' -f $message)
        if ($message -match 'Unauthorized|denied|403') {
            Write-PubLog -Level Info -Message 'The app most likely lacks the SharePoint Sites.FullControl.All application permission. Re-run setup option 1 and choose the tenant-admin scope, or let discovery fall back to Graph enumeration.'
        }
        return $false
    }
}

function Disconnect-PubPnP {
    [CmdletBinding()]
    param()

    if (-not $script:PnPConnected) { return }
    try { Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch { }
    $script:PnPConnected = $false
}

function Get-PubPnPSiteUrl {
    <#
    .SYNOPSIS
        Returns every site collection URL from the SharePoint tenant admin API.

    .DESCRIPTION
        This is the complete list - the whole point of the PnP route. Redirect
        stubs left behind by site renames and a few system templates are
        dropped, because they hold no user content and only waste a Graph
        lookup each. Teams channel sites are deliberately kept: they are real
        document stores and are exactly what the search-index route misses.

    .PARAMETER IncludePersonalSites
        Also return OneDrive for Business sites.
    #>
    [CmdletBinding()]
    param(
        [switch] $IncludePersonalSites,
        $Config
    )

    if (-not $Config) { $Config = Get-PubConfig }
    if (-not (Connect-PubPnPAdmin -Config $Config)) { return @() }

    $skipTemplates = @(
        'REDIRECTSITE#0'          # stub left behind when a site is renamed
        'SRCHCEN#0'               # search centre
        'SPSMSITEHOST#0'          # OneDrive host
        'POINTPUBLISHINGHUB#0'    # Viva Engage / video hub
        'POINTPUBLISHINGTOPIC#0'
        'APPCATALOG#0'
    )

    try {
        $parameters = @{ ErrorAction = 'Stop' }
        if ($IncludePersonalSites) { $parameters['IncludeOneDriveSites'] = $true }

        $sites = Get-PnPTenantSite @parameters
    } catch {
        Write-PubLog -Level Warn -Message ('Get-PnPTenantSite failed: {0}' -f $_.Exception.Message)
        return @()
    }

    $urls    = New-Object System.Collections.Generic.List[string]
    $skipped = 0

    foreach ($site in $sites) {
        if (-not $site.PSObject.Properties['Url'] -or [string]::IsNullOrWhiteSpace($site.Url)) { continue }

        $template = ''
        if ($site.PSObject.Properties['Template'] -and $site.Template) { $template = [string] $site.Template }

        if ($skipTemplates -contains $template) { $skipped++; continue }

        $urls.Add(([string] $site.Url).TrimEnd('/'))
    }

    Write-PubLog -Level Success -Message ('SharePoint admin site list: {0} site collection(s){1}.' -f $urls.Count, $(if ($skipped -gt 0) { ", $skipped system/redirect site(s) skipped" } else { '' }))

    return $urls.ToArray()
}

function Register-PubPnPApplication {
    <#
    .SYNOPSIS
        One-step app registration: app, certificate, upload and consent.

    .DESCRIPTION
        Wraps Register-PnPEntraIDApp (Register-PnPAzureADApp on older PnP
        versions), which creates the Entra ID application, generates a
        self-signed certificate, uploads the public key and opens the consent
        flow. That replaces four separate Graph calls and the manual consent
        URL step.

        The cmdlet's return shape has varied between versions, so the app id
        and thumbprint are read defensively and, failing that, looked up from
        the certificate files it wrote.

    .OUTPUTS
        An object with AppId, Thumbprint, NotAfter, CerPath and PfxPath, or
        $null if registration failed.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $DisplayName,
        [Parameter(Mandatory)] [string] $Tenant,

        [ValidateSet('AllSites', 'SitesSelected', 'TenantAdmin')]
        [string] $AuthMethod = 'TenantAdmin',

        [int] $ValidYears = 2,
        [switch] $DeviceLogin
    )

    if (-not (Test-PubPnPAvailable)) { return $null }

    $command = Get-Command -Name 'Register-PnPEntraIDApp' -ErrorAction SilentlyContinue
    if (-not $command) { $command = Get-Command -Name 'Register-PnPAzureADApp' -ErrorAction SilentlyContinue }
    if (-not $command) {
        Write-PubLog -Level Error -Message 'This PnP version has neither Register-PnPEntraIDApp nor Register-PnPAzureADApp.'
        return $null
    }

    $permissions      = Get-PubPermissionCatalog -AuthMethod $AuthMethod
    $graphPermissions = @($permissions | Where-Object { $_.Resource -eq 'Graph' }      | ForEach-Object { $_.Name })
    $spoPermissions   = @($permissions | Where-Object { $_.Resource -eq 'SharePoint' } | ForEach-Object { $_.Name })

    $outPath = Join-Path (Get-PubProjectRoot) '.certs'
    if (-not (Test-Path -LiteralPath $outPath)) { New-Item -Path $outPath -ItemType Directory -Force | Out-Null }

    $parameters = @{
        ApplicationName = $DisplayName
        Tenant          = $Tenant
        OutPath         = $outPath
        ValidYears      = $ValidYears
        ErrorAction     = 'Stop'
    }

    if ($graphPermissions.Count -gt 0) { $parameters['GraphApplicationPermissions']      = $graphPermissions }
    if ($spoPermissions.Count -gt 0)   { $parameters['SharePointApplicationPermissions'] = $spoPermissions }
    if ($DeviceLogin)                  { $parameters['DeviceLogin']                      = $true }
    if (Test-PubIsWindows)             { $parameters['Store']                            = 'CurrentUser' }

    Write-Host ''
    Write-PubLog -Level Info -Message ('Registering "{0}" via {1}...' -f $DisplayName, $command.Name)
    Write-PubLog -Level Info -Message 'A browser window will open twice: once to sign in, once to grant admin consent. Complete both.'

    try {
        Import-Module -Name 'PnP.PowerShell' -ErrorAction Stop -DisableNameChecking
        $result = & $command.Name @parameters
    } catch {
        Write-PubLog -Level Error -Message ('Registration failed: {0}' -f $_.Exception.Message)
        return $null
    }

    if (-not $result) {
        Write-PubLog -Level Error -Message 'The registration cmdlet returned nothing.'
        return $null
    }

    # --- read the app id and certificate back, whatever shape came out ---
    $appId = $null
    foreach ($name in @('AppId', 'ApplicationId', 'ClientId')) {
        if ($result.PSObject.Properties[$name] -and $result.$name) { $appId = [string] $result.$name; break }
    }

    $certificate = $null
    if ($result.PSObject.Properties['Certificate'] -and $result.Certificate) { $certificate = $result.Certificate }

    $thumbprint = $null
    foreach ($name in @('Thumbprint', 'CertificateThumbprint')) {
        if ($result.PSObject.Properties[$name] -and $result.$name) { $thumbprint = [string] $result.$name; break }
    }
    if (-not $thumbprint -and $certificate -and $certificate.PSObject.Properties['Thumbprint']) {
        $thumbprint = [string] $certificate.Thumbprint
    }

    $cerPath = ''
    $pfxPath = ''
    foreach ($file in (Get-ChildItem -LiteralPath $outPath -Filter ('{0}*' -f $DisplayName) -File -ErrorAction SilentlyContinue)) {
        if ($file.Extension -eq '.cer') { $cerPath = $file.FullName }
        if ($file.Extension -eq '.pfx') { $pfxPath = $file.FullName }
    }

    $notAfter = (Get-Date).AddYears($ValidYears)
    if (-not [string]::IsNullOrWhiteSpace($cerPath)) {
        try {
            $x509       = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 ($cerPath)
            $notAfter   = $x509.NotAfter
            if (-not $thumbprint) { $thumbprint = $x509.Thumbprint }
        } catch {
            Write-PubLog -Level Debug -Message ('Could not read {0}: {1}' -f $cerPath, $_.Exception.Message)
        }
    }

    if (-not $appId) {
        Write-PubLog -Level Error -Message 'Registration ran but no App ID could be read from the result. Check the Entra ID portal, then use "Connect to an existing app registration".'
        return $null
    }

    if (-not $thumbprint) {
        Write-PubLog -Level Warn -Message 'Registration ran but no certificate thumbprint could be read. Run setup option 2 to generate and upload one.'
    }

    if (-not [string]::IsNullOrWhiteSpace($pfxPath)) {
        Write-PubLog -Level Warn -Message ('{0} holds the private key - protect it, and do not commit or e-mail it.' -f $pfxPath)
    }

    Write-PubLog -Level Success -Message ('Registered. AppId {0}' -f $appId)

    return [pscustomobject] @{
        AppId      = $appId
        Thumbprint = $thumbprint
        NotAfter   = $notAfter
        CerPath    = $cerPath
        PfxPath    = $pfxPath
    }
}

Export-ModuleMember -Function @(
    'Test-PubPnPAvailable'
    'Get-PubTenantAdminUrl'
    'Connect-PubPnPAdmin'
    'Disconnect-PubPnP'
    'Get-PubPnPSiteUrl'
    'Register-PubPnPApplication'
)
