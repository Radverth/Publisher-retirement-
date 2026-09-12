<#
.SYNOPSIS
    Local configuration store for the SharePoint Publisher File Converter.

.DESCRIPTION
    Persists tenant / app registration / certificate details and the operator's
    behaviour settings to config.json next to Start-Menu.ps1, so the tool can
    reconnect silently on later runs (brief section 6.4).

    NOTHING SECRET IS WRITTEN HERE. Only the certificate thumbprint (a public
    identifier used to look the certificate up in the Windows certificate
    store) and the public .cer path are persisted. Private keys and .pfx
    passwords go to SecretManagement or a DPAPI-protected file - see
    Set-PubSecret / Get-PubSecret below - and never to config.json, the CSV or
    the logs.

    Compatible with Windows PowerShell 5.1 and PowerShell 7.
#>

Set-StrictMode -Version 2.0

if (-not (Get-Module -Name 'Logging')) {
    Import-Module (Join-Path $PSScriptRoot 'Logging.psm1') -Force -DisableNameChecking
}

$script:ProjectRoot = Split-Path -Parent $PSScriptRoot
$script:ConfigPath  = Join-Path $script:ProjectRoot 'config.json'
$script:VaultName   = 'SPO-PubConverter'

function ConvertTo-PubArray {
    <#
    .SYNOPSIS
        Converts any collection to a plain object[], safely.

    .DESCRIPTION
        Wrapping a System.Collections.Generic.List in the array subexpression
        @( ) throws "Argument types do not match" on some PowerShell builds -
        a runtime code-generation defect rather than anything wrong with the
        script. It bites at run time, on the host the tool is deployed to,
        which is the worst place to find it, so collections that might be a
        List go through here instead of @( ).

    .EXAMPLE
        $rows = ConvertTo-PubArray $Rows
    #>
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)]
        $InputObject
    )

    # Every return is comma-wrapped. Returning an array from a PowerShell
    # function otherwise unrolls it in the pipeline: an empty array comes back
    # as $null and a one-item array as the bare item, so the caller's .Count
    # either throws under Set-StrictMode or reports the wrong thing. The unary
    # comma makes the array itself the single output.
    if ($null -eq $InputObject)      { return ,@() }
    if ($InputObject -is [object[]]) { return ,$InputObject }

    if (($InputObject -is [System.Collections.IEnumerable]) -and ($InputObject -isnot [string])) {
        $items = New-Object System.Collections.ArrayList
        foreach ($item in $InputObject) { [void] $items.Add($item) }
        return ,$items.ToArray()
    }

    return ,@($InputObject)
}

function Test-PubIsWindows {
    <#
    .SYNOPSIS
        True when running on Windows, on both PowerShell 5.1 and 7.

    .DESCRIPTION
        $IsWindows only exists on PowerShell 6+, so testing it directly breaks
        under Windows PowerShell 5.1 with Set-StrictMode enabled.
    #>
    [CmdletBinding()]
    param()

    if ($PSVersionTable.PSEdition -eq 'Desktop') { return $true }
    return [bool] (Get-Variable -Name 'IsWindows' -ValueOnly -ErrorAction SilentlyContinue)
}

function Get-PubProjectRoot {
    [CmdletBinding()]
    param()
    return $script:ProjectRoot
}

function Get-PubConfigPath {
    [CmdletBinding()]
    param()
    return $script:ConfigPath
}

function New-PubDefaultConfig {
    <#
    .SYNOPSIS
        The shape of config.json, with safe defaults for a first run.
    #>
    [CmdletBinding()]
    param()

    return [ordered] @{
        TenantId                = ''
        TenantDomain            = ''
        AppId                   = ''
        AppObjectId             = ''
        AppDisplayName          = ''
        CertificateThumbprint   = ''
        CertificateExpiry       = ''
        CertificatePublicPath   = ''
        CertificatePfxPath      = ''
        CertificateSecretName   = ''   # where the .pfx password is stored (never the password itself)
        AuthMethod              = 'AllSites'          # AllSites | SitesSelected | TenantAdmin
        EnumerationMethod       = 'Auto'              # Auto | PnP | Graph
        SharePointAdminUrl      = ''                  # override; derived from the tenant domain when blank
        DefaultWorkingFolder    = (Join-Path $script:ProjectRoot 'working')
        LastInventoryCsv        = ''
        ExistingPdfAction       = 'Version'           # Skip | Overwrite | Version
        UploadConflictAction    = 'Version'           # Skip | Overwrite | Version
        RemoveSourceAfterUpload = $false
        ScopeSiteListPath       = ''
        LastUpdated             = ''
    }
}

function Get-PubConfig {
    <#
    .SYNOPSIS
        Reads config.json, merged over the defaults so new keys always exist.

    .DESCRIPTION
        Returns an ordered hashtable. A missing or unreadable config file is not
        an error - the operator simply has not run setup yet.
    #>
    [CmdletBinding()]
    param(
        [string] $Path
    )

    if (-not $Path) { $Path = $script:ConfigPath }
    $config = New-PubDefaultConfig

    if (-not (Test-Path -LiteralPath $Path)) {
        return $config
    }

    try {
        $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) { return $config }

        $stored = $raw | ConvertFrom-Json -ErrorAction Stop
        foreach ($property in $stored.PSObject.Properties) {
            $config[$property.Name] = $property.Value
        }
    } catch {
        Write-PubLog -Level Warn -Message ("Could not read {0} ({1}). Using defaults." -f $Path, $_.Exception.Message)
    }

    return $config
}

function Save-PubConfig {
    <#
    .SYNOPSIS
        Writes config.json.

    .DESCRIPTION
        Refuses to persist anything that looks like secret material, so a future
        edit cannot accidentally start writing passwords to disk.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Config,

        [string] $Path
    )

    if (-not $Path) { $Path = $script:ConfigPath }

    $forbidden = @('Password', 'Secret', 'ClientSecret', 'PfxPassword', 'PrivateKey', 'Thumbprint_Password')
    foreach ($key in @($Config.Keys)) {
        foreach ($bad in $forbidden) {
            if ($key -like "*$bad*") {
                Write-PubLog -Level Warn -Message ("Refusing to persist '{0}' to config.json - secret material is not stored here." -f $key)
                $Config.Remove($key)
            }
        }
    }

    $Config['LastUpdated'] = (Get-Date).ToString('s')

    try {
        $json = $Config | ConvertTo-Json -Depth 6
        Set-Content -LiteralPath $Path -Value $json -Encoding UTF8 -ErrorAction Stop
        Write-PubLog -Level Debug -Message ("Configuration saved to {0}" -f $Path)
    } catch {
        Write-PubLog -Level Error -Message ("Failed to save configuration to {0}: {1}" -f $Path, $_.Exception.Message)
        throw
    }

    return $Path
}

function Set-PubConfigValue {
    <#
    .SYNOPSIS
        Convenience helper: read config, set one key, write it back.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Name,
        [Parameter(Mandatory)] [AllowEmptyString()] $Value,
        [string] $Path
    )

    $config = Get-PubConfig -Path $Path
    $config[$Name] = $Value
    Save-PubConfig -Config $config -Path $Path | Out-Null
    return $config
}

function Test-PubConfigComplete {
    <#
    .SYNOPSIS
        True when config.json holds enough detail to attempt a silent connect.
    #>
    [CmdletBinding()]
    param($Config)

    if (-not $Config) { $Config = Get-PubConfig }

    $hasIds  = -not ([string]::IsNullOrWhiteSpace($Config['TenantId']) -or [string]::IsNullOrWhiteSpace($Config['AppId']))
    $hasCert = -not ([string]::IsNullOrWhiteSpace($Config['CertificateThumbprint']))

    return ($hasIds -and $hasCert)
}

function Test-PubCertificateExpiry {
    <#
    .SYNOPSIS
        Checks the recorded certificate expiry and warns inside the threshold.

    .DESCRIPTION
        Called at every startup (brief section 9), not just during setup.
        Returns an object with State = Valid | Expiring | Expired | Unknown.
    #>
    [CmdletBinding()]
    param(
        $Config,
        [int] $WarnWithinDays = 30,
        [switch] $Quiet
    )

    if (-not $Config) { $Config = Get-PubConfig }

    $result = [pscustomobject] @{
        State      = 'Unknown'
        Expiry     = $null
        DaysLeft   = $null
        Message    = 'No certificate expiry recorded - run setup option 2.'
    }

    $expiryRaw = $Config['CertificateExpiry']
    if ([string]::IsNullOrWhiteSpace($expiryRaw)) {
        if (-not $Quiet) { Write-PubLog -Level Debug -Message $result.Message }
        return $result
    }

    $expiry = [datetime]::MinValue
    if (-not [datetime]::TryParse($expiryRaw, [ref] $expiry)) {
        $result.Message = "Certificate expiry '$expiryRaw' is not a readable date."
        if (-not $Quiet) { Write-PubLog -Level Warn -Message $result.Message }
        return $result
    }

    $daysLeft        = [int] ([math]::Floor(($expiry - (Get-Date)).TotalDays))
    $result.Expiry   = $expiry
    $result.DaysLeft = $daysLeft

    if ($daysLeft -lt 0) {
        $result.State   = 'Expired'
        $result.Message = 'Authentication certificate EXPIRED on {0} - re-run setup option 2 before doing anything else.' -f $expiry.ToString('yyyy-MM-dd')
        if (-not $Quiet) { Write-PubLog -Level Error -Message $result.Message }
    } elseif ($daysLeft -le $WarnWithinDays) {
        $result.State   = 'Expiring'
        $result.Message = 'Authentication certificate expires in {0} day(s) on {1} - renew it soon (setup option 2).' -f $daysLeft, $expiry.ToString('yyyy-MM-dd')
        if (-not $Quiet) { Write-PubLog -Level Warn -Message $result.Message }
    } else {
        $result.State   = 'Valid'
        $result.Message = 'Certificate valid until {0} ({1} days).' -f $expiry.ToString('yyyy-MM-dd'), $daysLeft
        if (-not $Quiet) { Write-PubLog -Level Debug -Message $result.Message }
    }

    return $result
}

function Get-PubWorkingFolder {
    <#
    .SYNOPSIS
        Returns the working folder, creating it and its subfolders on demand.

    .PARAMETER SubFolder
        'originals' for downloaded .pub files, 'converted' for produced PDFs,
        'inventory' for CSV exports. Omit for the root working folder.
    #>
    [CmdletBinding()]
    param(
        [ValidateSet('originals', 'converted', 'inventory')]
        [string] $SubFolder,

        $Config
    )

    if (-not $Config) { $Config = Get-PubConfig }

    $root = $Config['DefaultWorkingFolder']
    if ([string]::IsNullOrWhiteSpace($root)) {
        $root = Join-Path $script:ProjectRoot 'working'
    }

    $path = $root
    if ($SubFolder) { $path = Join-Path $root $SubFolder }

    if (-not (Test-Path -LiteralPath $path)) {
        New-Item -Path $path -ItemType Directory -Force | Out-Null
    }

    return (Resolve-Path -LiteralPath $path).Path
}

function Set-PubSecret {
    <#
    .SYNOPSIS
        Stores a secret (e.g. a generated .pfx password) outside config.json.

    .DESCRIPTION
        Prefers Microsoft.PowerShell.SecretManagement when a vault is
        registered. Falls back to Export-Clixml of a SecureString, which on
        Windows is DPAPI-encrypted to the current user and machine - still not
        plain text, but note the limitation in the README.

        Returns the storage method used so the caller can tell the operator.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Name,
        [Parameter(Mandatory)] [securestring] $Secret
    )

    if (Get-Command -Name 'Set-Secret' -ErrorAction SilentlyContinue) {
        try {
            $vault = Get-SecretVault -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($vault) {
                Set-Secret -Name $Name -SecureStringSecret $Secret -Vault $vault.Name -ErrorAction Stop
                Write-PubLog -Level Success -Message ("Secret '{0}' stored in SecretManagement vault '{1}'." -f $Name, $vault.Name)
                return 'SecretManagement'
            }
        } catch {
            Write-PubLog -Level Warn -Message ("SecretManagement store failed ({0}); falling back to a protected local file." -f $_.Exception.Message)
        }
    }

    $secretFolder = Join-Path $script:ProjectRoot '.secrets'
    if (-not (Test-Path -LiteralPath $secretFolder)) {
        New-Item -Path $secretFolder -ItemType Directory -Force | Out-Null
    }

    $secretFile = Join-Path $secretFolder ("{0}.xml" -f ($Name -replace '[^A-Za-z0-9_\-]', '_'))
    $Secret | Export-Clixml -LiteralPath $secretFile -Force

    if (-not (Test-PubIsWindows)) {
        Write-PubLog -Level Warn -Message ("Secret '{0}' written to {1}. On non-Windows hosts Export-Clixml does NOT encrypt - register a SecretManagement vault instead." -f $Name, $secretFile)
    } else {
        Write-PubLog -Level Info -Message ("Secret '{0}' written to {1} (DPAPI-protected for this user on this machine)." -f $Name, $secretFile)
    }

    return 'LocalFile'
}

function Get-PubCertificateSecretName {
    <#
    .SYNOPSIS
        The name the .pfx password is stored under, derived from the .pfx path.

    .DESCRIPTION
        Keyed off the certificate file itself so the writer and the reader
        always agree. An earlier build stored it under the certificate subject
        name but read it back under the App ID, which meant the password was
        never found and the .pfx was opened with none at all.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [AllowEmptyString()] [string] $PfxPath
    )

    if ([string]::IsNullOrWhiteSpace($PfxPath)) { return '' }
    return ('PfxPassword_{0}' -f [System.IO.Path]::GetFileNameWithoutExtension($PfxPath))
}

function Get-PubCertificatePassword {
    <#
    .SYNOPSIS
        Retrieves the .pfx password for the configured certificate.

    .DESCRIPTION
        Tries the name recorded in config first, then the name derived from the
        .pfx file, then the older App ID and thumbprint spellings, so a
        certificate created by any previous version of this tool still opens
        without being regenerated.

        Returns $null when nothing is found - callers must treat that as "do
        not attempt to open the .pfx", not as "use an empty password".
    #>
    [CmdletBinding()]
    param(
        $Config,
        [switch] $Quiet
    )

    if (-not $Config) { $Config = Get-PubConfig }

    $candidates = New-Object System.Collections.Generic.List[string]

    $recorded = [string] $Config['CertificateSecretName']
    if (-not [string]::IsNullOrWhiteSpace($recorded)) { $candidates.Add($recorded) }

    $fromFile = Get-PubCertificateSecretName -PfxPath ([string] $Config['CertificatePfxPath'])
    if (-not [string]::IsNullOrWhiteSpace($fromFile)) { $candidates.Add($fromFile) }

    foreach ($legacy in @($Config['AppId'], $Config['CertificateThumbprint'])) {
        if (-not [string]::IsNullOrWhiteSpace([string] $legacy)) { $candidates.Add(('PfxPassword_{0}' -f $legacy)) }
    }

    $tried = New-Object System.Collections.Generic.List[string]
    foreach ($name in $candidates) {
        if ($tried -contains $name) { continue }
        $tried.Add($name)

        $secret = Get-PubSecret -Name $name
        if ($secret) {
            Write-PubLog -Level Debug -Message ('Certificate password found under "{0}".' -f $name)
            return $secret
        }
    }

    if (-not $Quiet) {
        Write-PubLog -Level Error -Message ('No stored password found for {0}.' -f $Config['CertificatePfxPath'])
        Write-PubLog -Level Info  -Message ('Looked for: {0}' -f ($tried -join ', '))
        Write-PubLog -Level Info  -Message ('In the SecretManagement vault, and in {0}' -f (Join-Path $script:ProjectRoot '.secrets'))
        Write-PubLog -Level Info  -Message 'If the password is genuinely gone, re-run setup option 2 to generate and upload a replacement certificate.'
    }

    return $null
}

function Get-PubSecret {
    <#
    .SYNOPSIS
        Retrieves a secret stored by Set-PubSecret. Returns $null if absent.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Name
    )

    if (Get-Command -Name 'Get-Secret' -ErrorAction SilentlyContinue) {
        try {
            $value = Get-Secret -Name $Name -ErrorAction Stop
            if ($value) { return $value }
        } catch {
            Write-PubLog -Level Debug -Message ("Secret '{0}' not found in SecretManagement." -f $Name)
        }
    }

    $secretFile = Join-Path (Join-Path $script:ProjectRoot '.secrets') ("{0}.xml" -f ($Name -replace '[^A-Za-z0-9_\-]', '_'))
    if (Test-Path -LiteralPath $secretFile) {
        try { return (Import-Clixml -LiteralPath $secretFile) } catch { return $null }
    }

    return $null
}

Export-ModuleMember -Function @(
    'ConvertTo-PubArray'
    'Test-PubIsWindows'
    'Get-PubProjectRoot'
    'Get-PubConfigPath'
    'New-PubDefaultConfig'
    'Get-PubConfig'
    'Save-PubConfig'
    'Set-PubConfigValue'
    'Test-PubConfigComplete'
    'Test-PubCertificateExpiry'
    'Get-PubWorkingFolder'
    'Set-PubSecret'
    'Get-PubSecret'
    'Get-PubCertificateSecretName'
    'Get-PubCertificatePassword'
)
