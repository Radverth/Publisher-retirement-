<#
.SYNOPSIS
    Phase 0 - detect, create and validate the Azure AD app registration and the
    certificate used for unattended Microsoft Graph / SharePoint access.

.DESCRIPTION
    Idempotent by design (brief section 3): every entry point checks what
    already exists before creating anything, so re-running setup is safe.

    APPROACH CHOSEN - Microsoft Graph SDK, not PnP PowerShell
    ---------------------------------------------------------
    The brief asks for one approach, documented. This build uses a bespoke
    single-tenant app registration driven by the Microsoft Graph PowerShell SDK
    for everything: app registration management, site crawling, download and
    upload.

    Trade-off vs registering the PnP Management Shell multi-tenant app:
      + One principal the customer owns, named for this tool, revocable on its
        own without affecting any other PnP-based tooling in the tenant.
      + Permissions are explicit and auditable in the app's own blade, and can
        be narrowed to Sites.Selected per-site grants.
      + No dependency on a Microsoft-operated multi-tenant app that other
        Affinity IT tooling may also rely on - revoking it here would break
        those too.
      - Slightly more setup code than 'Register-PnPManagementShellAccess', and
        CSOM-only SharePoint features are unavailable. Nothing this tool needs
        (enumerate sites, read/write drive items) requires CSOM.

    BLAST RADIUS - READ THIS BEFORE CONSENTING
    ------------------------------------------
    In the default AllSites mode this app is granted Sites.ReadWrite.All and
    Files.ReadWrite.All as APPLICATION permissions. That is tenant-wide read
    AND WRITE access to every SharePoint site and every OneDrive in the tenant,
    with no user context and no per-site restriction. Anyone who can run code
    as this app, or who holds its certificate, can read or modify any file in
    the tenant.

    Mitigations, in order of preference:
      1. Choose the Sites.Selected auth method at setup. The app then has no
         access at all until an administrator explicitly grants it to named
         sites (Grant-PubSiteSelectedPermission below, or menu option 1 ->
         'Grant this app access to a specific site'). This is the tighter,
         recommended posture for a scoped retirement project.
      2. Keep the private key in the certificate store, not as a loose .pfx.
      3. Delete the app registration when the Publisher retirement project is
         finished - it is a project tool, not permanent infrastructure.
#>

Set-StrictMode -Version 2.0

foreach ($dependency in @('Logging', 'Config', 'Graph')) {
    if (-not (Get-Module -Name $dependency)) {
        Import-Module (Join-Path $PSScriptRoot ("{0}.psm1" -f $dependency)) -Force -DisableNameChecking
    }
}

$script:GraphAppId      = '00000003-0000-0000-c000-000000000000'
$script:SharePointAppId = '00000003-0000-0ff1-ce00-000000000000'

function Get-PubPermissionCatalog {
    <#
    .SYNOPSIS
        The Microsoft Graph application permissions this tool can request.

    .PARAMETER AuthMethod
        AllSites      - tenant-wide read/write (default, biggest blast radius).
        SitesSelected - no access until granted per-site by an administrator.
    #>
    [CmdletBinding()]
    param(
        [ValidateSet('AllSites', 'SitesSelected', 'TenantAdmin')]
        [string] $AuthMethod = 'AllSites'
    )

    $all = @(
        [pscustomobject] @{ Name = 'Sites.Read.All';      Resource = 'Graph';      Id = '332a536c-c7ef-4017-ab91-336970924f0d'; Purpose = 'Enumerate site collections and crawl document libraries during discovery'; Methods = @('AllSites', 'TenantAdmin') }
        [pscustomobject] @{ Name = 'Sites.ReadWrite.All'; Resource = 'Graph';      Id = '9492366f-7969-46a4-8d15-ed1a20078fff'; Purpose = 'Upload converted PDFs back to the source library';                        Methods = @('AllSites', 'TenantAdmin') }
        [pscustomobject] @{ Name = 'Files.ReadWrite.All'; Resource = 'Graph';      Id = '75359482-378d-4052-8f01-80520e7db3cd'; Purpose = 'Download source .pub files and write PDFs via Graph drive items';         Methods = @('AllSites', 'TenantAdmin') }
        [pscustomobject] @{ Name = 'Sites.Selected';      Resource = 'Graph';      Id = '883ea226-0bf2-4a8f-9f9d-92c9162a727d'; Purpose = 'Access only the sites an administrator explicitly grants (tight scope)';  Methods = @('SitesSelected') }
        [pscustomobject] @{ Name = 'Directory.Read.All';  Resource = 'Graph';      Id = '7ab1d382-f21e-4acd-a863-ba3e13f7da61'; Purpose = 'Resolve site/user metadata where needed (optional)';                      Methods = @('AllSites', 'SitesSelected', 'TenantAdmin') }
        [pscustomobject] @{ Name = 'Sites.FullControl.All'; Resource = 'SharePoint'; Id = ''; Purpose = 'Read the SharePoint tenant admin site list, so discovery finds every site without a manual export'; Methods = @('TenantAdmin') }
    )

    return $all | Where-Object { $_.Methods -contains $AuthMethod }
}

function Get-PubResourceAppId {
    <#
    .SYNOPSIS
        Maps a catalogue Resource name to its well-known application id.
    #>
    [CmdletBinding()]
    param(
        [ValidateSet('Graph', 'SharePoint')]
        [string] $Resource = 'Graph'
    )

    if ($Resource -eq 'SharePoint') { return $script:SharePointAppId }
    return $script:GraphAppId
}

function Resolve-PubAppRole {
    <#
    .SYNOPSIS
        Finds a permission's app role id on the resource service principal.

    .DESCRIPTION
        Resolved live from the tenant by permission name rather than trusting a
        hardcoded GUID. The catalogue's Id is only a fallback for when the
        lookup cannot run - and the SharePoint permissions have no hardcoded id
        at all, so they must come from here.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Permission,
        $ResourceServicePrincipal
    )

    if ($ResourceServicePrincipal -and $ResourceServicePrincipal.PSObject.Properties['appRoles']) {
        foreach ($role in $ResourceServicePrincipal.appRoles) {
            if ($role.value -eq $Permission.Name) { return $role.id }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($Permission.Id)) { return $Permission.Id }

    Write-PubLog -Level Warn -Message ('Could not resolve the app role id for {0} on the {1} API.' -f $Permission.Name, $Permission.Resource)
    return $null
}

function Show-PubPermissionTable {
    [CmdletBinding()]
    param(
        [ValidateSet('AllSites', 'SitesSelected', 'TenantAdmin')]
        [string] $AuthMethod = 'AllSites'
    )

    Write-Host ''
    Write-Host 'Microsoft Graph application permissions to be requested:' -ForegroundColor Cyan
    Get-PubPermissionCatalog -AuthMethod $AuthMethod |
        Format-Table @{ N = 'Permission'; E = { $_.Name } }, @{ N = 'API'; E = { $_.Resource } }, @{ N = 'Type'; E = { 'Application' } }, @{ N = 'Purpose'; E = { $_.Purpose } } -AutoSize |
        Out-String | Write-Host

    if ($AuthMethod -eq 'TenantAdmin') {
        Write-Host 'WARNING: this scope includes SharePoint Sites.FullControl.All - full administrative' -ForegroundColor Red
        Write-Host '         control of every site collection in the tenant, above and beyond the' -ForegroundColor Red
        Write-Host '         tenant-wide read/write below. It is what lets discovery read the tenant' -ForegroundColor Red
        Write-Host '         admin site list, so no manual site export is needed. Delete the app' -ForegroundColor Red
        Write-Host '         registration when the Publisher retirement is finished.' -ForegroundColor Red
        Write-Host ''
    }

    if ($AuthMethod -eq 'AllSites' -or $AuthMethod -eq 'TenantAdmin') {
        Write-Host 'WARNING: Sites.ReadWrite.All and Files.ReadWrite.All are TENANT-WIDE write permissions.' -ForegroundColor Yellow
        Write-Host '         Every SharePoint site and OneDrive in the tenant becomes readable and writable' -ForegroundColor Yellow
        Write-Host '         by anything holding this app''s certificate. Choose Sites.Selected instead if you' -ForegroundColor Yellow
        Write-Host '         only need a handful of sites - it requires one extra per-site grant each.' -ForegroundColor Yellow
        Write-Host ''
    }
}

function Get-PubApplication {
    <#
    .SYNOPSIS
        Looks up an app registration by appId. Returns $null when absent.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $AppId
    )

    try {
        $uri      = "applications?`$filter=appId eq '$AppId'"
        $response = Invoke-PubGraph -Uri $uri -Method GET
        if ($response -and $response.PSObject.Properties['value'] -and $response.value) {
            return @($response.value)[0]
        }
    } catch {
        Write-PubLog -Level Debug -Message ('Application lookup failed: {0}' -f (Get-PubGraphErrorMessage -ErrorRecord $_))
    }

    return $null
}

function Get-PubServicePrincipal {
    <#
    .SYNOPSIS
        Looks up a service principal by appId. Returns $null when absent.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $AppId
    )

    try {
        $uri      = "servicePrincipals?`$filter=appId eq '$AppId'"
        $response = Invoke-PubGraph -Uri $uri -Method GET
        if ($response -and $response.PSObject.Properties['value'] -and $response.value) {
            return @($response.value)[0]
        }
    } catch {
        Write-PubLog -Level Debug -Message ('Service principal lookup failed: {0}' -f (Get-PubGraphErrorMessage -ErrorRecord $_))
    }

    return $null
}

function New-PubAppRegistration {
    <#
    .SYNOPSIS
        Creates the single-tenant app registration and its service principal.

    .DESCRIPTION
        Requires an existing interactive Graph session with
        Application.ReadWrite.All. Returns an object with AppId, ObjectId and
        ServicePrincipalId, or $null on failure.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $DisplayName,

        [ValidateSet('AllSites', 'SitesSelected', 'TenantAdmin')]
        [string] $AuthMethod = 'AllSites'
    )

    $permissions            = Get-PubPermissionCatalog -AuthMethod $AuthMethod
    $requiredResourceAccess = @()

    foreach ($resourceName in @('Graph', 'SharePoint')) {
        $forResource = @($permissions | Where-Object { $_.Resource -eq $resourceName })
        if ($forResource.Count -eq 0) { continue }

        $resourceAppId           = Get-PubResourceAppId -Resource $resourceName
        $resourceServicePrincipal = Get-PubServicePrincipal -AppId $resourceAppId
        $resourceAccess          = @()

        foreach ($permission in $forResource) {
            $roleId = Resolve-PubAppRole -Permission $permission -ResourceServicePrincipal $resourceServicePrincipal
            if ($roleId) { $resourceAccess += @{ id = $roleId; type = 'Role' } }
        }

        if ($resourceAccess.Count -gt 0) {
            $requiredResourceAccess += @{ resourceAppId = $resourceAppId; resourceAccess = $resourceAccess }
        }
    }

    $body = @{
        displayName            = $DisplayName
        signInAudience         = 'AzureADMyOrg'
        description            = 'Created by the SharePoint Publisher File Converter to find and convert .pub files.'
        requiredResourceAccess = $requiredResourceAccess
    }

    try {
        Write-PubLog -Level Info -Message ('Creating app registration "{0}"...' -f $DisplayName)
        $application = Invoke-PubGraph -Uri 'applications' -Method POST -Body $body -ContentType 'application/json'
        Write-PubLog -Level Success -Message ('App registration created. AppId {0}' -f $application.appId)
    } catch {
        Write-PubLog -Level Error -Message ('Could not create the app registration: {0}' -f (Get-PubGraphErrorMessage -ErrorRecord $_))
        return $null
    }

    # Directory replication can lag; retry the service principal creation briefly.
    $servicePrincipal = $null
    for ($attempt = 1; $attempt -le 5; $attempt++) {
        try {
            $servicePrincipal = Invoke-PubGraph -Uri 'servicePrincipals' -Method POST -Body @{ appId = $application.appId } -ContentType 'application/json'
            break
        } catch {
            $existing = Get-PubServicePrincipal -AppId $application.appId
            if ($existing) { $servicePrincipal = $existing; break }

            if ($attempt -eq 5) {
                Write-PubLog -Level Error -Message ('Could not create the service principal: {0}' -f (Get-PubGraphErrorMessage -ErrorRecord $_))
            } else {
                Write-PubLog -Level Warn -Message ('Service principal not ready yet - retrying in {0}s...' -f (2 * $attempt))
                Start-Sleep -Seconds (2 * $attempt)
            }
        }
    }

    if ($servicePrincipal) {
        Write-PubLog -Level Success -Message ('Service principal created ({0}).' -f $servicePrincipal.id)
    }

    $servicePrincipalId = $null
    if ($servicePrincipal) { $servicePrincipalId = $servicePrincipal.id }

    return [pscustomobject] @{
        AppId              = $application.appId
        ObjectId           = $application.id
        DisplayName        = $application.displayName
        ServicePrincipalId = $servicePrincipalId
        AuthMethod         = $AuthMethod
    }
}

function Grant-PubAdminConsent {
    <#
    .SYNOPSIS
        Grants the requested application permissions (tenant admin consent).

    .DESCRIPTION
        Assigns each app role directly via appRoleAssignedTo, which is the
        programmatic equivalent of clicking 'Grant admin consent'. If the
        signed-in account cannot do that, the tenant-specific consent URL is
        printed and the operator completes it in a browser.

        Returns $true when every permission is confirmed granted.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $AppId,
        [string] $ServicePrincipalId,

        [ValidateSet('AllSites', 'SitesSelected', 'TenantAdmin')]
        [string] $AuthMethod = 'AllSites',

        [Parameter(Mandatory)] [string] $TenantId
    )

    if (-not $ServicePrincipalId) {
        $servicePrincipal = Get-PubServicePrincipal -AppId $AppId
        if ($servicePrincipal) { $ServicePrincipalId = $servicePrincipal.id }
    }

    if (-not $ServicePrincipalId) {
        Write-PubLog -Level Warn -Message 'Could not resolve the service principal needed to grant consent automatically.'
        Show-PubConsentUrl -AppId $AppId -TenantId $TenantId
        return $false
    }

    $permissions = Get-PubPermissionCatalog -AuthMethod $AuthMethod
    $granted     = 0
    $failed      = 0

    foreach ($resourceName in @('Graph', 'SharePoint')) {
        $forResource = @($permissions | Where-Object { $_.Resource -eq $resourceName })
        if ($forResource.Count -eq 0) { continue }

        $resourceServicePrincipal = Get-PubServicePrincipal -AppId (Get-PubResourceAppId -Resource $resourceName)
        if (-not $resourceServicePrincipal) {
            Write-PubLog -Level Warn -Message ('The {0} service principal is not present in this tenant - its permissions cannot be granted from here.' -f $resourceName)
            $failed += $forResource.Count
            continue
        }

        foreach ($permission in $forResource) {
            $roleId = Resolve-PubAppRole -Permission $permission -ResourceServicePrincipal $resourceServicePrincipal
            if (-not $roleId) { $failed++; continue }

            $body = @{
                principalId = $ServicePrincipalId
                resourceId  = $resourceServicePrincipal.id
                appRoleId   = $roleId
            }

            try {
                Invoke-PubGraph -Uri ("servicePrincipals/{0}/appRoleAssignedTo" -f $resourceServicePrincipal.id) -Method POST -Body $body -ContentType 'application/json' | Out-Null
                Write-PubLog -Level Success -Message ('Granted {0} ({1}).' -f $permission.Name, $resourceName)
                $granted++
            } catch {
                $message = Get-PubGraphErrorMessage -ErrorRecord $_
                if ($message -match 'Permission being assigned already exists') {
                    Write-PubLog -Level Info -Message ('{0} ({1}) was already granted.' -f $permission.Name, $resourceName)
                    $granted++
                } else {
                    Write-PubLog -Level Warn -Message ('Could not grant {0} ({1}): {2}' -f $permission.Name, $resourceName, $message)
                    $failed++
                }
            }
        }
    }

    if ($failed -gt 0) {
        Write-PubLog -Level Warn -Message 'One or more permissions could not be granted from here - the signed-in account probably lacks consent rights.'
        Show-PubConsentUrl -AppId $AppId -TenantId $TenantId

        $answer = Read-Host 'Press Enter once admin consent has been completed in the browser (or type S to skip)'
        if ($answer -notmatch '^[Ss]') {
            return (Test-PubAdminConsent -AppId $AppId -AuthMethod $AuthMethod)
        }
        return $false
    }

    return ($granted -gt 0)
}

function Show-PubConsentUrl {
    <#
    .SYNOPSIS
        Prints the tenant-specific admin consent URL for manual completion.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $AppId,
        [Parameter(Mandatory)] [string] $TenantId
    )

    $url = 'https://login.microsoftonline.com/{0}/adminconsent?client_id={1}' -f $TenantId, $AppId

    Write-Host ''
    Write-Host 'ADMIN CONSENT REQUIRED' -ForegroundColor Yellow
    Write-Host 'Open this URL as a Global Administrator and approve the permissions:' -ForegroundColor Yellow
    Write-Host ''
    Write-Host ('  {0}' -f $url) -ForegroundColor Cyan
    Write-Host ''

    Write-PubLog -Level Info -Message ('Admin consent URL: {0}' -f $url) -NoConsole
}

function Test-PubAdminConsent {
    <#
    .SYNOPSIS
        Confirms every required app role is actually assigned to the app.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $AppId,

        [ValidateSet('AllSites', 'SitesSelected', 'TenantAdmin')]
        [string] $AuthMethod = 'AllSites'
    )

    $servicePrincipal = Get-PubServicePrincipal -AppId $AppId
    if (-not $servicePrincipal) {
        Write-PubLog -Level Warn -Message 'No service principal found for this app - consent cannot be confirmed.'
        return $false
    }

    try {
        $assignments = Get-PubGraphAll -Uri ("servicePrincipals/{0}/appRoleAssignments" -f $servicePrincipal.id)
    } catch {
        Write-PubLog -Level Warn -Message ('Could not read app role assignments: {0}' -f (Get-PubGraphErrorMessage -ErrorRecord $_))
        return $false
    }

    $assignedIds = @()
    foreach ($assignment in $assignments) { $assignedIds += $assignment.appRoleId }

    $resourceServicePrincipals = @{}
    foreach ($resourceName in @('Graph', 'SharePoint')) {
        $resourceServicePrincipals[$resourceName] = Get-PubServicePrincipal -AppId (Get-PubResourceAppId -Resource $resourceName)
    }

    $missing = @()
    foreach ($permission in (Get-PubPermissionCatalog -AuthMethod $AuthMethod)) {
        $roleId = Resolve-PubAppRole -Permission $permission -ResourceServicePrincipal $resourceServicePrincipals[$permission.Resource]
        if (-not $roleId -or $assignedIds -notcontains $roleId) { $missing += ('{0} ({1})' -f $permission.Name, $permission.Resource) }
    }

    if ($missing.Count -gt 0) {
        Write-PubLog -Level Warn -Message ('Still waiting on admin consent for: {0}' -f ($missing -join ', '))
        return $false
    }

    Write-PubLog -Level Success -Message 'All required application permissions are consented.'
    return $true
}

function New-PubAuthCertificate {
    <#
    .SYNOPSIS
        Creates the self-signed authentication certificate.

    .DESCRIPTION
        Windows: New-SelfSignedCertificate into CurrentUser\My, 2048-bit,
        two-year validity, plus a .cer export for upload.
        Non-Windows: OpenSSL fallback producing a .key/.cer/.pfx set under
        <project>\.certs, with the generated .pfx password stored via
        Set-PubSecret - never in config.json or the log.

        Returns an object with Thumbprint, NotAfter, CerPath, PfxPath, Base64.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $SubjectName,
        [int] $ValidityYears = 2
    )

    $certificateFolder = Join-Path (Get-PubProjectRoot) '.certs'
    if (-not (Test-Path -LiteralPath $certificateFolder)) {
        New-Item -Path $certificateFolder -ItemType Directory -Force | Out-Null
    }

    $safeName = $SubjectName -replace '[^A-Za-z0-9_\-]', '_'
    $cerPath  = Join-Path $certificateFolder ("{0}.cer" -f $safeName)
    $pfxPath  = Join-Path $certificateFolder ("{0}.pfx" -f $safeName)

    if (Test-PubIsWindows) {
        try {
            Write-PubLog -Level Info -Message ('Generating a self-signed certificate CN={0} (valid {1} years)...' -f $SubjectName, $ValidityYears)

            $certificate = New-SelfSignedCertificate -Subject ("CN={0}" -f $SubjectName) `
                                                     -CertStoreLocation 'Cert:\CurrentUser\My' `
                                                     -KeyExportPolicy Exportable `
                                                     -KeySpec Signature `
                                                     -KeyLength 2048 `
                                                     -KeyAlgorithm RSA `
                                                     -HashAlgorithm SHA256 `
                                                     -NotAfter (Get-Date).AddYears($ValidityYears) `
                                                     -ErrorAction Stop

            Export-Certificate -Cert $certificate -FilePath $cerPath -Force | Out-Null

            Write-PubLog -Level Success -Message ('Certificate created. Thumbprint {0}, expires {1}.' -f $certificate.Thumbprint, $certificate.NotAfter.ToString('yyyy-MM-dd'))
            Write-PubLog -Level Info -Message  ('Private key stays in Cert:\CurrentUser\My - only the public key ({0}) is uploaded.' -f $cerPath)

            return [pscustomobject] @{
                Thumbprint = $certificate.Thumbprint
                NotAfter   = $certificate.NotAfter
                CerPath    = $cerPath
                PfxPath    = ''
                Base64     = [System.Convert]::ToBase64String($certificate.RawData)
            }
        } catch {
            Write-PubLog -Level Error -Message ('Certificate generation failed: {0}' -f $_.Exception.Message)
            return $null
        }
    }

    # ---- OpenSSL fallback for PowerShell 7 on Linux / macOS ----
    if (-not (Get-Command -Name 'openssl' -ErrorAction SilentlyContinue)) {
        Write-PubLog -Level Error -Message 'OpenSSL was not found. Install OpenSSL, or run setup on a Windows host.'
        return $null
    }

    $keyPath  = Join-Path $certificateFolder ("{0}.key" -f $safeName)
    $days     = $ValidityYears * 365
    $password = [System.Guid]::NewGuid().ToString('N')

    try {
        Write-PubLog -Level Info -Message 'Generating a self-signed certificate with OpenSSL...'

        & openssl req -x509 -newkey rsa:2048 -sha256 -nodes `
            -keyout $keyPath -out $cerPath -days $days `
            -subj ("/CN={0}" -f $SubjectName) 2>$null
        if ($LASTEXITCODE -ne 0) { throw 'openssl req failed.' }

        & openssl pkcs12 -export -out $pfxPath -inkey $keyPath -in $cerPath -passout ("pass:{0}" -f $password) 2>$null
        if ($LASTEXITCODE -ne 0) { throw 'openssl pkcs12 failed.' }

        $securePassword = ConvertTo-SecureString -String $password -AsPlainText -Force
        Set-PubSecret -Name ('PfxPassword_{0}' -f $safeName) -Secret $securePassword | Out-Null
        $password = $null

        $x509 = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 ($cerPath)

        Write-PubLog -Level Success -Message ('Certificate created. Thumbprint {0}, expires {1}.' -f $x509.Thumbprint, $x509.NotAfter.ToString('yyyy-MM-dd'))
        Write-PubLog -Level Warn -Message  ('Protect {0} - it holds the private key. Never commit it or e-mail it.' -f $pfxPath)

        return [pscustomobject] @{
            Thumbprint = $x509.Thumbprint
            NotAfter   = $x509.NotAfter
            CerPath    = $cerPath
            PfxPath    = $pfxPath
            Base64     = [System.Convert]::ToBase64String($x509.RawData)
        }
    } catch {
        Write-PubLog -Level Error -Message ('OpenSSL certificate generation failed: {0}' -f $_.Exception.Message)
        return $null
    }
}

function Add-PubCertificateToApp {
    <#
    .SYNOPSIS
        Uploads the public key to the app registration's keyCredentials.

    .DESCRIPTION
        Existing key credentials are read first and preserved, so uploading a
        renewal certificate does not invalidate the one still in use.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $ApplicationObjectId,
        [Parameter(Mandatory)] [string] $Base64Certificate,
        [string] $DisplayName = 'SPO-PubConverter authentication certificate'
    )

    try {
        $application    = Invoke-PubGraph -Uri ("applications/{0}?`$select=keyCredentials" -f $ApplicationObjectId) -Method GET
        $keyCredentials = @()

        if ($application -and $application.PSObject.Properties['keyCredentials'] -and $application.keyCredentials) {
            foreach ($existing in $application.keyCredentials) {
                $keyCredentials += @{
                    type        = $existing.type
                    usage       = $existing.usage
                    key         = $existing.key
                    displayName = $existing.displayName
                }
            }
        }

        $keyCredentials += @{
            type        = 'AsymmetricX509Cert'
            usage       = 'Verify'
            key         = $Base64Certificate
            displayName = $DisplayName
        }

        Invoke-PubGraph -Uri ("applications/{0}" -f $ApplicationObjectId) -Method PATCH -Body @{ keyCredentials = $keyCredentials } -ContentType 'application/json' | Out-Null

        Write-PubLog -Level Success -Message 'Public key uploaded to the app registration.'
        return $true
    } catch {
        Write-PubLog -Level Error -Message ('Could not upload the certificate: {0}' -f (Get-PubGraphErrorMessage -ErrorRecord $_))
        return $false
    }
}

function Grant-PubSiteSelectedPermission {
    <#
    .SYNOPSIS
        Grants the app read/write on one specific site (Sites.Selected mode).

    .PARAMETER SiteUrl
        Full site URL, e.g. https://contoso.sharepoint.com/sites/Marketing

    .PARAMETER Role
        read or write. Discovery needs read; upload needs write.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $SiteUrl,

        [ValidateSet('read', 'write')]
        [string] $Role = 'write',

        $Config
    )

    if (-not $Config) { $Config = Get-PubConfig }

    try {
        $parsedUrl  = [uri] $SiteUrl.TrimEnd('/')
        $siteHost   = $parsedUrl.Host
        $sitePath   = $parsedUrl.AbsolutePath.TrimEnd('/')

        $site = Invoke-PubGraph -Uri ("sites/{0}:{1}" -f $siteHost, $sitePath) -Method GET

        $body = @{
            roles               = @($Role)
            grantedToIdentities = @(
                @{ application = @{ id = $Config['AppId']; displayName = $Config['AppDisplayName'] } }
            )
        }

        Invoke-PubGraph -Uri ("sites/{0}/permissions" -f $site.id) -Method POST -Body $body -ContentType 'application/json' | Out-Null

        Write-PubLog -Level Success -Message ('Granted "{0}" on {1}.' -f $Role, $SiteUrl)
        return $true
    } catch {
        Write-PubLog -Level Error -Message ('Could not grant site permission: {0}' -f (Get-PubGraphErrorMessage -ErrorRecord $_))
        Write-PubLog -Level Info -Message  'Granting Sites.Selected access needs an account with Sites.FullControl.All, or an administrator running this step.'
        return $false
    }
}

function Test-PubGraphAccess {
    <#
    .SYNOPSIS
        Menu option 3 - proves the app-only connection actually works.

    .DESCRIPTION
        Connects with the certificate, reads the tenant organisation, then
        performs a real SharePoint read (enumerate a site and its drives) so a
        missing SharePoint permission is caught here rather than mid-crawl.
    #>
    [CmdletBinding()]
    param(
        $Config
    )

    if (-not $Config) { $Config = Get-PubConfig }

    if (-not (Connect-PubGraphApp -Config $Config -Force)) { return $false }

    $ok = $true

    try {
        $organization = Invoke-PubGraph -Uri 'organization?$select=id,displayName,verifiedDomains' -Method GET
        if ($organization -and $organization.value) {
            $tenant = @($organization.value)[0]
            Write-PubLog -Level Success -Message ('Graph OK - tenant "{0}".' -f $tenant.displayName)
        }
    } catch {
        Write-PubLog -Level Warn -Message ('Directory read failed: {0}' -f (Get-PubGraphErrorMessage -ErrorRecord $_))
        Write-PubLog -Level Info -Message  'Directory.Read.All is optional - continuing with the SharePoint check.'
    }

    try {
        $rootSite = Invoke-PubGraph -Uri 'sites/root?$select=id,webUrl,displayName' -Method GET
        Write-PubLog -Level Success -Message ('SharePoint OK - root site {0}' -f $rootSite.webUrl)

        $drives = Invoke-PubGraph -Uri ("sites/{0}/drives?`$select=id,name&`$top=1" -f $rootSite.id) -Method GET
        if ($drives -and $drives.value) {
            Write-PubLog -Level Success -Message ('Document library read OK - "{0}".' -f (@($drives.value)[0].name))
        }
    } catch {
        $ok = $false
        Write-PubLog -Level Error -Message ('SharePoint read failed: {0}' -f (Get-PubGraphErrorMessage -ErrorRecord $_))

        if ($Config['AuthMethod'] -eq 'SitesSelected') {
            Write-PubLog -Level Info -Message 'In Sites.Selected mode this is expected until a site grant exists - use menu option 1 to grant access to a site, then re-test.'
        } else {
            Write-PubLog -Level Info -Message 'Check that admin consent has been granted for Sites.Read.All / Sites.ReadWrite.All.'
        }
    }

    $expiry = Test-PubCertificateExpiry -Config $Config
    Write-PubLog -Level Info -Message $expiry.Message

    return $ok
}

function Get-PubAppRegistrationStatus {
    <#
    .SYNOPSIS
        One-line app registration state for the menu header.
    #>
    [CmdletBinding()]
    param(
        $Config
    )

    if (-not $Config) { $Config = Get-PubConfig }

    if (-not (Test-PubConfigComplete -Config $Config)) {
        if ([string]::IsNullOrWhiteSpace($Config['AppId'])) { return 'not configured' }
        return 'app set, certificate missing'
    }

    $expiry = Test-PubCertificateExpiry -Config $Config -Quiet
    switch ($expiry.State) {
        'Valid'    { return ('OK (expires {0})'       -f $expiry.Expiry.ToString('yyyy-MM-dd')) }
        'Expiring' { return ('EXPIRING {0} ({1}d)'    -f $expiry.Expiry.ToString('yyyy-MM-dd'), $expiry.DaysLeft) }
        'Expired'  { return ('EXPIRED {0}'            -f $expiry.Expiry.ToString('yyyy-MM-dd')) }
        default    { return 'configured (expiry unknown)' }
    }
}

Export-ModuleMember -Function @(
    'Get-PubPermissionCatalog'
    'Get-PubResourceAppId'
    'Resolve-PubAppRole'
    'Show-PubPermissionTable'
    'Get-PubApplication'
    'Get-PubServicePrincipal'
    'New-PubAppRegistration'
    'Grant-PubAdminConsent'
    'Show-PubConsentUrl'
    'Test-PubAdminConsent'
    'New-PubAuthCertificate'
    'Add-PubCertificateToApp'
    'Grant-PubSiteSelectedPermission'
    'Test-PubGraphAccess'
    'Get-PubAppRegistrationStatus'
)
