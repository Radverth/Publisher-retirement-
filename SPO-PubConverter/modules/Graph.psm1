<#
.SYNOPSIS
    Microsoft Graph connection and request plumbing shared by every phase.

.DESCRIPTION
    One place for: connecting (certificate app-only, or interactive delegated
    for setup), and for issuing Graph calls with retry-with-backoff so a
    tenant-wide crawl survives Graph's per-app throttling (brief section 9).

    Every Graph/SharePoint call in this tool goes through Invoke-PubGraph or
    Get-PubGraphAll so the retry behaviour is impossible to forget.

    Compatible with Windows PowerShell 5.1 and PowerShell 7.
#>

Set-StrictMode -Version 2.0

foreach ($dependency in @('Logging', 'Config')) {
    if (-not (Get-Module -Name $dependency)) {
        Import-Module (Join-Path $PSScriptRoot ("{0}.psm1" -f $dependency)) -Force -DisableNameChecking
    }
}

$script:GraphBase          = 'https://graph.microsoft.com/v1.0'
$script:RetryStatusCodes   = @(429, 500, 502, 503, 504)
$script:ConnectedAppOnly   = $false

function Get-PubGraphBaseUri {
    [CmdletBinding()]
    param()
    return $script:GraphBase
}

function Test-PubGraphModule {
    <#
    .SYNOPSIS
        Verifies the Microsoft Graph PowerShell SDK is present.

    .PARAMETER Install
        Offer to install the required submodules from the PSGallery for the
        current user if they are missing.
    #>
    [CmdletBinding()]
    param(
        [switch] $Install
    )

    # Every call in this tool goes through Invoke-MgGraphRequest, which lives in
    # Microsoft.Graph.Authentication - the other SDK submodules are not needed.
    $required = @('Microsoft.Graph.Authentication')

    $missing = @()
    foreach ($module in $required) {
        if (-not (Get-Module -ListAvailable -Name $module)) { $missing += $module }
    }

    if ($missing.Count -eq 0) { return $true }

    Write-PubLog -Level Warn -Message ('Missing PowerShell module(s): {0}' -f ($missing -join ', '))

    if (-not $Install) {
        Write-PubLog -Level Info -Message 'Install with: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser'
        return $false
    }

    foreach ($module in $missing) {
        try {
            Write-PubLog -Level Info -Message ('Installing {0} for the current user...' -f $module)
            Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            Write-PubLog -Level Success -Message ('Installed {0}.' -f $module)
        } catch {
            Write-PubLog -Level Error -Message ('Could not install {0}: {1}' -f $module, $_.Exception.Message)
            return $false
        }
    }

    return $true
}

function Get-PubGraphContext {
    <#
    .SYNOPSIS
        Returns the current Graph context, or $null when not connected.
    #>
    [CmdletBinding()]
    param()

    if (-not (Get-Command -Name 'Get-MgContext' -ErrorAction SilentlyContinue)) { return $null }
    try { return Get-MgContext } catch { return $null }
}

function Test-PubGraphConnected {
    <#
    .SYNOPSIS
        True when a Graph session is live.

    .PARAMETER AppOnly
        Require the live session to be app-only (certificate) rather than an
        interactive delegated session.
    #>
    [CmdletBinding()]
    param(
        [switch] $AppOnly
    )

    $context = Get-PubGraphContext
    if (-not $context) { return $false }
    if ($AppOnly -and -not $script:ConnectedAppOnly) { return $false }
    return $true
}

function Get-PubCertificate {
    <#
    .SYNOPSIS
        Resolves the authentication certificate from the store or a .pfx file.

    .DESCRIPTION
        On Windows the thumbprint is looked up in CurrentUser\My then
        LocalMachine\My. On PowerShell 7 for Linux/macOS there is no certificate
        store, so the .pfx recorded in config is loaded instead (its password is
        read from SecretManagement / the protected local file).
    #>
    [CmdletBinding()]
    param(
        $Config
    )

    if (-not $Config) { $Config = Get-PubConfig }
    $thumbprint = $Config['CertificateThumbprint']

    if ((Test-PubIsWindows) -and -not [string]::IsNullOrWhiteSpace($thumbprint)) {
        foreach ($storePath in @('Cert:\CurrentUser\My', 'Cert:\LocalMachine\My')) {
            $certificate = Get-ChildItem -Path $storePath -ErrorAction SilentlyContinue |
                Where-Object { $_.Thumbprint -eq $thumbprint } |
                Select-Object -First 1
            if ($certificate) { return $certificate }
        }
        Write-PubLog -Level Warn -Message ("Certificate {0} not found in CurrentUser\My or LocalMachine\My." -f $thumbprint)
    }

    $pfxPath = [string] $Config['CertificatePfxPath']
    if (-not [string]::IsNullOrWhiteSpace($pfxPath) -and (Test-Path -LiteralPath $pfxPath)) {
        $password = Get-PubCertificatePassword -Config $Config
        if (-not $password) {
            # Opening a password-protected .pfx with no password fails with a
            # misleading "data cannot be read" error, so stop here instead.
            return $null
        }

        # EphemeralKeySet keeps the private key out of any on-disk store, but is
        # not supported on every platform - fall back to the default flags.
        foreach ($flags in @('EphemeralKeySet', 'DefaultKeySet')) {
            try {
                return New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 (
                    $pfxPath, $password, $flags
                )
            } catch {
                Write-PubLog -Level Debug -Message ('Loading {0} with {1} failed: {2}' -f $pfxPath, $flags, $_.Exception.Message)
            }
        }

        Write-PubLog -Level Error -Message ('Could not load the certificate from {0}. If it was created on another machine, copy the .secrets folder across as well, or re-run setup option 2.' -f $pfxPath)
    }

    return $null
}

function Connect-PubGraphApp {
    <#
    .SYNOPSIS
        Connects to Graph app-only using the certificate recorded in config.

    .DESCRIPTION
        This is the silent reconnect used by discovery, download and upload.
        Returns $true on success. Never throws - the menu reports the failure
        and stays open.
    #>
    [CmdletBinding()]
    param(
        $Config,
        [switch] $Force
    )

    if (-not $Config) { $Config = Get-PubConfig }

    if (-not (Test-PubConfigComplete -Config $Config)) {
        Write-PubLog -Level Error -Message 'No app registration configured yet - run menu option 1 (and 2) first.'
        return $false
    }

    if ((Test-PubGraphConnected -AppOnly) -and -not $Force) {
        return $true
    }

    if (-not (Test-PubGraphModule)) { return $false }

    $expiry = Test-PubCertificateExpiry -Config $Config
    if ($expiry.State -eq 'Expired') {
        Write-PubLog -Level Error -Message $expiry.Message
        return $false
    }

    try {
        if (Test-PubGraphConnected) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }

        if (Test-PubIsWindows) {
            Connect-MgGraph -ClientId $Config['AppId'] `
                            -TenantId $Config['TenantId'] `
                            -CertificateThumbprint $Config['CertificateThumbprint'] `
                            -NoWelcome -ErrorAction Stop
        } else {
            $certificate = Get-PubCertificate -Config $Config
            if (-not $certificate) {
                Write-PubLog -Level Error -Message 'No usable certificate found for app-only sign-in.'
                return $false
            }
            Connect-MgGraph -ClientId $Config['AppId'] `
                            -TenantId $Config['TenantId'] `
                            -Certificate $certificate `
                            -NoWelcome -ErrorAction Stop
        }

        $script:ConnectedAppOnly = $true
        Write-PubLog -Level Success -Message ('Connected to Microsoft Graph app-only as {0}.' -f $Config['AppId'])
        return $true
    } catch {
        $script:ConnectedAppOnly = $false
        Write-PubLog -Level Error -Message ('App-only sign-in failed: {0}' -f $_.Exception.Message)
        Write-PubLog -Level Info -Message 'If the app registration was only just consented, wait a minute and try menu option 3 again.'
        return $false
    }
}

function Connect-PubGraphInteractive {
    <#
    .SYNOPSIS
        Interactive delegated sign-in, used only by the setup phase.

    .PARAMETER Scopes
        Delegated scopes to request. Setup needs to create applications and
        grant app role assignments, so the signed-in account must hold
        Application Administrator / Global Administrator.
    #>
    [CmdletBinding()]
    param(
        [string]   $TenantId,
        [string[]] $Scopes = @(
            'Application.ReadWrite.All'
            'AppRoleAssignment.ReadWrite.All'
            'Directory.Read.All'
        )
    )

    if (-not (Test-PubGraphModule)) { return $false }

    try {
        if (Test-PubGraphConnected) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }

        $parameters = @{ Scopes = $Scopes; NoWelcome = $true; ErrorAction = 'Stop' }
        if (-not [string]::IsNullOrWhiteSpace($TenantId)) { $parameters['TenantId'] = $TenantId }

        Write-PubLog -Level Info -Message 'A browser sign-in window will open - use an account with Application Administrator or Global Administrator rights.'
        Connect-MgGraph @parameters

        $script:ConnectedAppOnly = $false
        $context = Get-PubGraphContext
        if ($context) {
            Write-PubLog -Level Success -Message ('Signed in as {0} on tenant {1}.' -f $context.Account, $context.TenantId)
        }
        return $true
    } catch {
        Write-PubLog -Level Error -Message ('Interactive sign-in failed: {0}' -f $_.Exception.Message)
        return $false
    }
}

function Disconnect-PubGraph {
    [CmdletBinding()]
    param()

    if (Test-PubGraphConnected) {
        try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
    }
    $script:ConnectedAppOnly = $false
}

function Get-PubHttpStatusCode {
    <#
    .SYNOPSIS
        Digs the HTTP status code out of whatever shape of error Graph threw.

    .DESCRIPTION
        The SDK surfaces failures differently between PowerShell 5.1 and 7 and
        between cmdlet and raw request calls, so probe several shapes and fall
        back to parsing the message.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $ErrorRecord
    )

    $exception = $ErrorRecord.Exception

    foreach ($candidate in @($exception, $exception.InnerException)) {
        if (-not $candidate) { continue }

        $response = $candidate.PSObject.Properties['Response']
        if ($response -and $response.Value) {
            $statusCode = $response.Value.PSObject.Properties['StatusCode']
            if ($statusCode -and $statusCode.Value) {
                try { return [int] $statusCode.Value } catch { }
            }
        }

        $direct = $candidate.PSObject.Properties['StatusCode']
        if ($direct -and $direct.Value) {
            try { return [int] $direct.Value } catch { }
        }

        $responseStatus = $candidate.PSObject.Properties['ResponseStatusCode']
        if ($responseStatus -and $responseStatus.Value) {
            try { return [int] $responseStatus.Value } catch { }
        }
    }

    if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message -match '"?status"?\s*:\s*(\d{3})') {
        return [int] $Matches[1]
    }
    if ($exception.Message -match '\b(4\d\d|5\d\d)\b') {
        return [int] $Matches[1]
    }

    return 0
}

function Get-PubRetryAfterSeconds {
    <#
    .SYNOPSIS
        Reads the Retry-After header Graph sends with a 429, if present.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $ErrorRecord
    )

    foreach ($candidate in @($ErrorRecord.Exception, $ErrorRecord.Exception.InnerException)) {
        if (-not $candidate) { continue }

        $response = $candidate.PSObject.Properties['Response']
        if (-not $response -or -not $response.Value) { continue }

        $headers = $response.Value.PSObject.Properties['Headers']
        if (-not $headers -or -not $headers.Value) { continue }

        try {
            $retryAfter = $headers.Value.RetryAfter
            if ($retryAfter -and $retryAfter.Delta) {
                return [int] $retryAfter.Delta.TotalSeconds
            }
        } catch { }

        try {
            $raw = $headers.Value['Retry-After']
            if ($raw) { return [int] ($raw | Select-Object -First 1) }
        } catch { }
    }

    return 0
}

function Invoke-PubGraph {
    <#
    .SYNOPSIS
        Issues one Graph request with throttling-aware retry.

    .DESCRIPTION
        Retries 429 (honouring Retry-After) and transient 5xx responses with
        exponential backoff - 2s, 4s, 8s, 16s - up to MaxAttempts. 4xx failures
        other than 429 are permanent and rethrown immediately so the caller can
        mark that file Failed and move on.

    .PARAMETER Uri
        Absolute URI, or a path relative to https://graph.microsoft.com/v1.0.

    .PARAMETER OutputFilePath
        Stream the response body to this file (used for downloads).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Uri,
        [ValidateSet('GET', 'POST', 'PATCH', 'PUT', 'DELETE')]
        [string]      $Method = 'GET',
        $Body,
        [hashtable]   $Headers,
        [string]      $ContentType,
        [string]      $OutputFilePath,
        [int]         $MaxAttempts = 5
    )

    if ($Uri -notmatch '^https?://') {
        $Uri = '{0}/{1}' -f $script:GraphBase, $Uri.TrimStart('/')
    }

    $attempt = 0
    while ($true) {
        $attempt++
        try {
            $parameters = @{
                Method      = $Method
                Uri         = $Uri
                ErrorAction = 'Stop'
            }

            if ($null -ne $Body) {
                if ($Body -is [string] -or $Body -is [byte[]]) { $parameters['Body'] = $Body }
                else { $parameters['Body'] = ($Body | ConvertTo-Json -Depth 10) }
            }
            if ($Headers)        { $parameters['Headers']        = $Headers }
            if ($ContentType)    { $parameters['ContentType']    = $ContentType }
            if ($OutputFilePath) { $parameters['OutputFilePath'] = $OutputFilePath }
            else                 { $parameters['OutputType']     = 'PSObject' }

            return Invoke-MgGraphRequest @parameters
        } catch {
            $status = Get-PubHttpStatusCode -ErrorRecord $_

            if (($script:RetryStatusCodes -notcontains $status) -or ($attempt -ge $MaxAttempts)) {
                if ($status -gt 0) {
                    Write-PubLog -Level Debug -Message ('Graph {0} {1} failed with HTTP {2} after {3} attempt(s).' -f $Method, $Uri, $status, $attempt)
                }
                throw
            }

            $wait = Get-PubRetryAfterSeconds -ErrorRecord $_
            if ($wait -le 0) { $wait = [math]::Pow(2, $attempt) }
            if ($wait -gt 120) { $wait = 120 }

            $reason = 'HTTP {0}' -f $status
            if ($status -eq 429) { $reason = 'throttled (HTTP 429)' }

            Write-PubLog -Level Warn -Message ('Graph {0} - waiting {1}s then retrying (attempt {2}/{3}).' -f $reason, [int] $wait, $attempt, $MaxAttempts)

            # Show the wait counting down, so a throttled run does not look
            # like a hang - this is the single most common reason a tenant-wide
            # crawl appears to stop dead.
            $remaining = [int] $wait
            while ($remaining -gt 0) {
                Write-Progress -Id 9 -Activity 'Microsoft Graph is throttling this app' `
                               -Status ('Waiting {0}s before retry {1} of {2}' -f $remaining, $attempt, $MaxAttempts) `
                               -PercentComplete ([int] ((([int] $wait - $remaining) / [math]::Max([int] $wait, 1)) * 100))
                Start-Sleep -Seconds 1
                $remaining--
            }
            Write-Progress -Id 9 -Activity 'Microsoft Graph is throttling this app' -Completed
        }
    }
}

function Get-PubGraphAll {
    <#
    .SYNOPSIS
        Runs a Graph collection request and follows @odata.nextLink to the end.

    .PARAMETER MaxItems
        Stop after this many items (0 = no limit). Useful for test runs.

    .PARAMETER OnPage
        Called with each page of results as it arrives, as
        & $OnPage $itemsInThisPage $pageNumber $itemsSoFar. Used so a caller can
        report progress while a large library is being paged through, instead
        of going silent for hundreds of sequential requests.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Uri,
        [int] $MaxItems = 0,
        [int] $MaxAttempts = 5,
        [scriptblock] $OnPage
    )

    $results = New-Object System.Collections.Generic.List[object]
    $next    = $Uri
    $page    = 0

    while ($next) {
        $response = Invoke-PubGraph -Uri $next -Method GET -MaxAttempts $MaxAttempts
        $page++

        if ($response -and $response.PSObject.Properties['value'] -and $null -ne $response.value) {
            $pageItems = New-Object System.Collections.Generic.List[object]

            foreach ($item in $response.value) {
                $results.Add($item)
                $pageItems.Add($item)
                if ($MaxItems -gt 0 -and $results.Count -ge $MaxItems) {
                    if ($OnPage) { & $OnPage $pageItems.ToArray() $page $results.Count }
                    return $results.ToArray()
                }
            }

            if ($OnPage) { & $OnPage $pageItems.ToArray() $page $results.Count }
        } elseif ($response) {
            $results.Add($response)
            if ($OnPage) { & $OnPage @($response) $page $results.Count }
        }

        $next = $null
        if ($response -and $response.PSObject.Properties['@odata.nextLink']) {
            $next = $response.'@odata.nextLink'
        }
    }

    return $results.ToArray()
}

function Get-PubGraphErrorMessage {
    <#
    .SYNOPSIS
        Produces a short, CSV-safe error string for the Notes column.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $ErrorRecord
    )

    $message = $ErrorRecord.Exception.Message

    if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
        try {
            $details = $ErrorRecord.ErrorDetails.Message | ConvertFrom-Json -ErrorAction Stop
            if ($details.PSObject.Properties['error'] -and $details.error.PSObject.Properties['message']) {
                $message = '{0}: {1}' -f $details.error.code, $details.error.message
            }
        } catch {
            $message = $ErrorRecord.ErrorDetails.Message
        }
    }

    $status = Get-PubHttpStatusCode -ErrorRecord $ErrorRecord
    if ($status -gt 0) { $message = 'HTTP {0} - {1}' -f $status, $message }

    $message = ($message -replace '\s+', ' ').Trim()
    if ($message.Length -gt 400) { $message = $message.Substring(0, 397) + '...' }

    return $message
}

Export-ModuleMember -Function @(
    'Get-PubGraphBaseUri'
    'Test-PubGraphModule'
    'Get-PubGraphContext'
    'Test-PubGraphConnected'
    'Get-PubCertificate'
    'Connect-PubGraphApp'
    'Connect-PubGraphInteractive'
    'Disconnect-PubGraph'
    'Invoke-PubGraph'
    'Get-PubGraphAll'
    'Get-PubHttpStatusCode'
    'Get-PubRetryAfterSeconds'
    'Get-PubGraphErrorMessage'
)
