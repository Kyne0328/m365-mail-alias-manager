$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"
$Marker = "__ALIAS_MANAGER_JSON__"
$AliasNamespace = "m365am"

function Write-Envelope {
    param(
        [string]$Id,
        [bool]$Ok,
        [object]$Data,
        [string]$ErrorMessage,
        [string]$Category
    )

    $payload = if ($Ok) {
        [ordered]@{ id = $Id; ok = $true; data = $Data }
    }
    else {
        [ordered]@{ id = $Id; ok = $false; error = $ErrorMessage; category = $Category }
    }

    [Console]::Out.WriteLine("$Marker$($payload | ConvertTo-Json -Depth 12 -Compress)")
}

function Get-Certificate {
    if (-not $env:EXCHANGE_CERTIFICATE_BASE64) {
        throw "EXCHANGE_CERTIFICATE_BASE64 is not configured."
    }

    $bytes = [Convert]::FromBase64String($env:EXCHANGE_CERTIFICATE_BASE64)
    return [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
        $bytes,
        $env:EXCHANGE_CERTIFICATE_PASSWORD,
        [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable
    )
}

function ConvertTo-Base64Url {
    param([byte[]]$Bytes)
    return [Convert]::ToBase64String($Bytes).TrimEnd("=").Replace("+", "-").Replace("/", "_")
}

function Get-GraphAccessToken {
    param(
        [string]$TenantId,
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate
    )

    if ($TenantId -notmatch "^[0-9a-fA-F-]{36}$") {
        throw "The signed-in tenant ID is invalid."
    }

    $tokenUri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    $thumbprintHash = [System.Security.Cryptography.SHA1]::HashData($Certificate.RawData)

    $headerJson = [ordered]@{
        alg = "RS256"
        typ = "JWT"
        x5t = ConvertTo-Base64Url -Bytes $thumbprintHash
    } | ConvertTo-Json -Compress

    $payloadJson = [ordered]@{
        aud = $tokenUri
        iss = $env:ENTRA_CLIENT_ID
        sub = $env:ENTRA_CLIENT_ID
        jti = [Guid]::NewGuid().ToString()
        nbf = $now - 60
        exp = $now + 600
    } | ConvertTo-Json -Compress

    $encodedHeader = ConvertTo-Base64Url -Bytes ([Text.Encoding]::UTF8.GetBytes($headerJson))
    $encodedPayload = ConvertTo-Base64Url -Bytes ([Text.Encoding]::UTF8.GetBytes($payloadJson))
    $unsignedAssertion = "$encodedHeader.$encodedPayload"

    $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Certificate)
    try {
        $signature = $rsa.SignData(
            [Text.Encoding]::UTF8.GetBytes($unsignedAssertion),
            [System.Security.Cryptography.HashAlgorithmName]::SHA256,
            [System.Security.Cryptography.RSASignaturePadding]::Pkcs1
        )
    }
    finally {
        $rsa.Dispose()
    }

    $clientAssertion = "$unsignedAssertion.$(ConvertTo-Base64Url -Bytes $signature)"
    $token = Invoke-RestMethod -Method Post -Uri $tokenUri -ContentType "application/x-www-form-urlencoded" -Body @{
        client_id = $env:ENTRA_CLIENT_ID
        scope = "https://graph.microsoft.com/.default"
        grant_type = "client_credentials"
        client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
        client_assertion = $clientAssertion
    }

    if (-not $token.access_token) {
        throw "Microsoft Graph did not return an application access token."
    }

    return $token.access_token
}

$script:OrganizationCache = @{}

function Get-PrimaryOrganization {
    param(
        [string]$TenantId,
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate
    )

    $key = $TenantId.ToLowerInvariant()
    if ($script:OrganizationCache.ContainsKey($key)) {
        return $script:OrganizationCache[$key]
    }

    $accessToken = Get-GraphAccessToken -TenantId $TenantId -Certificate $Certificate
    $result = Invoke-RestMethod -Method Get -Uri 'https://graph.microsoft.com/v1.0/domains?$select=id,isInitial,isDefault,isVerified' -Headers @{
        Authorization = "Bearer $accessToken"
        Accept = "application/json"
    }

    $initial = @(
        $result.value |
            Where-Object { $_.isInitial -eq $true -and $_.id -like "*.onmicrosoft.com" } |
            Select-Object -First 1
    )

    if ($initial.Count -eq 0) {
        $initial = @(
            $result.value |
                Where-Object { $_.isVerified -eq $true -and $_.id -like "*.onmicrosoft.com" } |
                Select-Object -First 1
        )
    }

    if ($initial.Count -ne 1) {
        throw "Microsoft Graph could not discover the tenant's primary onmicrosoft.com domain."
    }

    $organization = $initial[0].id.ToString().ToLowerInvariant()
    $script:OrganizationCache[$key] = $organization
    return $organization
}

function Assert-AcceptedDomain {
    param([string]$Domain)

    $accepted = Get-AcceptedDomain -Identity $Domain -ErrorAction Stop
    if (-not $accepted) {
        throw "The selected domain is not accepted by this Exchange organization."
    }
}

function Get-UserMailbox {
    param([string]$UserId)

    if ($UserId -notmatch "^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$") {
        throw "The signed-in user identifier is invalid."
    }

    $mailboxes = @(
        Get-Mailbox -Filter "ExternalDirectoryObjectId -eq '$UserId'" -ResultSize 2 -ErrorAction Stop
    )

    if ($mailboxes.Count -ne 1) {
        throw "The signed-in Microsoft Entra user does not have a unique Exchange Online mailbox."
    }

    return $mailboxes[0]
}

function Get-AliasSet {
    param(
        [object]$Mailbox,
        [int]$Limit
    )

    $escapedNamespace = [Regex]::Escape($AliasNamespace)
    $pattern = "^smtp:$escapedNamespace-(?<prefix>[a-z0-9](?:[a-z0-9._-]{0,22}[a-z0-9])?)-(?<sequence>\d{6})-(?<nonce>[a-f0-9]{12})@(?<domain>[^@]+)$"

    $managed = @(
        $Mailbox.EmailAddresses |
            ForEach-Object { $_.ToString() } |
            ForEach-Object {
                $match = [Regex]::Match($_, $pattern, [Text.RegularExpressions.RegexOptions]::IgnoreCase)
                if ($match.Success) {
                    [pscustomobject]@{
                        address = ($_ -replace "^[^:]+:", "").ToLowerInvariant()
                        sequence = [int]$match.Groups["sequence"].Value
                        prefix = $match.Groups["prefix"].Value.ToLowerInvariant()
                        domain = $match.Groups["domain"].Value.ToLowerInvariant()
                    }
                }
            } |
            Sort-Object sequence -Descending
    )

    $next = if ($managed.Count -gt 0) {
        ([int]($managed | Measure-Object sequence -Maximum).Maximum) + 1
    }
    else {
        1
    }

    return [ordered]@{
        mailbox = $Mailbox.UserPrincipalName.ToString().ToLowerInvariant()
        primaryAddress = $Mailbox.PrimarySmtpAddress.ToString().ToLowerInvariant()
        aliases = @($managed)
        count = $managed.Count
        limit = $Limit
        nextSequence = $next
    }
}

function Get-Domains {
    return @(
        Get-AcceptedDomain |
            Sort-Object @{ Expression = "Default"; Descending = $true }, DomainName |
            ForEach-Object {
                [pscustomobject]@{
                    domain = $_.DomainName.ToString().ToLowerInvariant()
                    isDefault = [bool]$_.Default
                    type = $_.DomainType.ToString()
                }
            }
    )
}

function Get-ErrorCategory {
    param([System.Management.Automation.ErrorRecord]$Record)

    $text = ($Record | Out-String)
    if ($text -match "(?i)access is denied|unauthorized|forbidden|not authorized|role assignment|cmdlet .* not available|insufficient privileges|authorization_requestdenied|admin consent|AADSTS65001") {
        return "authorization"
    }
    if ($text -match "(?i)connect-exchangeonline|token|certificate|authentication") {
        return "connection"
    }
    if ($text -match "(?i)mailbox .* couldn't be found|mailbox .* not found|recipient .* couldn't be found|does not have a unique exchange online mailbox") {
        return "mailbox"
    }
    if ($text -match "(?i)accepted domain|domain .* not found") {
        return "domain"
    }
    return "exchange"
}

$script:ConnectedOrganization = $null

function Ensure-ExchangeConnection {
    param(
        [string]$Organization,
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate
    )

    if ($script:ConnectedOrganization -eq $Organization) {
        return
    }

    if ($script:ConnectedOrganization) {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        $script:ConnectedOrganization = $null
    }

    $connectParams = @{
        AppId = $env:ENTRA_CLIENT_ID
        Certificate = $Certificate
        Organization = $Organization
        ShowBanner = $false
        ShowProgress = $false
        SkipLoadingFormatData = $true
        CommandName = @("Get-Mailbox", "Set-Mailbox", "Get-AcceptedDomain")
        ErrorAction = "Stop"
    }
    Connect-ExchangeOnline @connectParams

    $script:ConnectedOrganization = $Organization
}

function Invoke-ExchangeAction {
    param(
        [object]$Request,
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate
    )

    $organization = if ($Request.action -eq "bootstrap") {
        $usernameDomain = ""
        if ($Request.username) {
            $parts = $Request.username.ToString().ToLowerInvariant().Split("@")
            if ($parts.Count -eq 2 -and $parts[1] -like "*.onmicrosoft.com") {
                $usernameDomain = $parts[1]
            }
        }

        if ($usernameDomain) {
            $usernameDomain
        }
        else {
            Get-PrimaryOrganization -TenantId $Request.tenantId -Certificate $Certificate
        }
    }
    else {
        $Request.organization.ToString().ToLowerInvariant()
    }

    Ensure-ExchangeConnection -Organization $organization -Certificate $Certificate

    switch ($Request.action) {
        "bootstrap" {
            $mailbox = Get-UserMailbox -UserId $Request.userId
            return [ordered]@{
                organization = $organization
                domains = @(Get-Domains)
                aliasSet = Get-AliasSet -Mailbox $mailbox -Limit $Request.limit
            }
        }

        "domains" {
            return @(Get-Domains)
        }

        "aliases" {
            $mailbox = Get-UserMailbox -UserId $Request.userId
            return Get-AliasSet -Mailbox $mailbox -Limit $Request.limit
        }

        "create" {
            Assert-AcceptedDomain -Domain $Request.domain
            $mailbox = Get-UserMailbox -UserId $Request.userId
            $current = Get-AliasSet -Mailbox $mailbox -Limit $Request.limit
            $allAddresses = @($mailbox.EmailAddresses | ForEach-Object { ($_ -replace "^[^:]+:", "").ToLowerInvariant() })

            $sequence = [int]$current.nextSequence
            do {
                if ($sequence -gt 999999) {
                    throw "The alias sequence has reached its six-digit limit."
                }
                $nonce = [Guid]::NewGuid().ToString("N").Substring(0, 12)
                $candidate = ("{0}-{1}-{2:D6}-{3}@{4}" -f $AliasNamespace, $Request.prefix, $sequence, $nonce, $Request.domain).ToLowerInvariant()
                $sequence++
            } while ($allAddresses -contains $candidate)

            $removeCount = [Math]::Max(0, $current.count - [int]$Request.limit + 1)
            $primaryAddress = $mailbox.PrimarySmtpAddress.ToString().ToLowerInvariant()
            $toRemove = @(
                $current.aliases |
                    Where-Object { $_.address -ne $primaryAddress } |
                    Sort-Object sequence |
                    Select-Object -First $removeCount |
                    ForEach-Object { "smtp:$($_.address)" }
            )

            if ($toRemove.Count -lt $removeCount) {
                throw "The FIFO rotation cannot remove enough aliases because an app-managed address is currently the mailbox primary SMTP address."
            }

            $changes = @{ Add = "smtp:$candidate" }
            if ($toRemove.Count -gt 0) {
                $changes.Remove = $toRemove
            }

            Set-Mailbox -Identity $mailbox.Identity -EmailAddresses $changes -ErrorAction Stop

            $updated = Get-UserMailbox -UserId $Request.userId
            return Get-AliasSet -Mailbox $updated -Limit $Request.limit
        }

        "delete" {
            $mailbox = Get-UserMailbox -UserId $Request.userId
            $current = Get-AliasSet -Mailbox $mailbox -Limit $Request.limit
            $target = $Request.address.ToString().ToLowerInvariant()
            $match = @($current.aliases | Where-Object { $_.address -eq $target })

            if ($match.Count -ne 1) {
                throw "The address is not part of this app's managed alias set."
            }

            if ($mailbox.PrimarySmtpAddress.ToString().ToLowerInvariant() -eq $target) {
                throw "The primary SMTP address cannot be removed."
            }

            Set-Mailbox -Identity $mailbox.Identity -EmailAddresses @{ Remove = "smtp:$target" } -ErrorAction Stop

            $updated = Get-UserMailbox -UserId $Request.userId
            return Get-AliasSet -Mailbox $updated -Limit $Request.limit
        }

        default {
            throw "Unsupported Exchange action."
        }
    }
}

try {
    if (-not $env:ENTRA_CLIENT_ID) {
        throw "ENTRA_CLIENT_ID is not configured."
    }

    $certificate = Get-Certificate
    Import-Module ExchangeOnlineManagement -ErrorAction Stop

    while (($line = [Console]::In.ReadLine()) -ne $null) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        $id = ""
        try {
            $message = $line | ConvertFrom-Json -Depth 20
            $id = $message.id.ToString()
            $data = Invoke-ExchangeAction -Request $message.input -Certificate $certificate
            Write-Envelope -Id $id -Ok $true -Data $data
        }
        catch {
            $category = Get-ErrorCategory -Record $_
            if ($category -eq "connection") {
                $script:ConnectedOrganization = $null
            }
            Write-Envelope -Id $id -Ok $false -ErrorMessage $_.Exception.Message -Category $category
        }
    }
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}
