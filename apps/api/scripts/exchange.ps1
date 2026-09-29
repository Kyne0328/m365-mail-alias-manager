$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"
$Marker = "__ALIAS_MANAGER_JSON__"
$AliasNamespace = "m365am"

function Write-Envelope {
    param(
        [bool]$Ok,
        [object]$Data,
        [string]$ErrorMessage,
        [string]$Category
    )

    $payload = if ($Ok) {
        [ordered]@{ ok = $true; data = $Data }
    }
    else {
        [ordered]@{ ok = $false; error = $ErrorMessage; category = $Category }
    }

    $json = $payload | ConvertTo-Json -Depth 10 -Compress
    [Console]::Out.WriteLine("$Marker$json")
}

function Get-Certificate {
    if (-not $env:EXCHANGE_CERTIFICATE_BASE64) {
        throw "EXCHANGE_CERTIFICATE_BASE64 is not configured."
    }

    $bytes = [Convert]::FromBase64String($env:EXCHANGE_CERTIFICATE_BASE64)
    $password = $env:EXCHANGE_CERTIFICATE_PASSWORD

    return [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
        $bytes,
        $password,
        [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable
    )
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
        aliases = $managed
        count = $managed.Count
        limit = $Limit
        nextSequence = $next
    }
}

function Get-ErrorCategory {
    param([System.Management.Automation.ErrorRecord]$Record)

    $text = ($Record | Out-String)
    if ($text -match "(?i)access is denied|unauthorized|forbidden|not authorized|role assignment|cmdlet .* not available") {
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

try {
    $raw = [Console]::In.ReadToEnd()
    if (-not $raw) {
        throw "No Exchange request was provided."
    }

    $request = $raw | ConvertFrom-Json -Depth 10
    if (-not $env:ENTRA_CLIENT_ID) {
        throw "ENTRA_CLIENT_ID is not configured."
    }

    $certificate = Get-Certificate
    Import-Module ExchangeOnlineManagement -ErrorAction Stop

    $connectParams = @{
        AppId = $env:ENTRA_CLIENT_ID
        Certificate = $certificate
        Organization = $request.organization
        ShowBanner = $false
        SkipLoadingFormatData = $true
        CommandName = @("Get-Mailbox", "Set-Mailbox", "Get-AcceptedDomain")
    }
    Connect-ExchangeOnline @connectParams

    try {
        switch ($request.action) {
            "domains" {
                $domains = @(
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

                Write-Envelope -Ok $true -Data $domains
                break
            }

            "aliases" {
                $mailbox = Get-UserMailbox -UserId $request.userId
                $result = Get-AliasSet -Mailbox $mailbox -Limit $request.limit
                Write-Envelope -Ok $true -Data $result
                break
            }

            "create" {
                Assert-AcceptedDomain -Domain $request.domain
                $mailbox = Get-UserMailbox -UserId $request.userId
                $current = Get-AliasSet -Mailbox $mailbox -Limit $request.limit
                $allAddresses = @($mailbox.EmailAddresses | ForEach-Object { ($_ -replace "^[^:]+:", "").ToLowerInvariant() })

                $sequence = [int]$current.nextSequence
                do {
                    if ($sequence -gt 999999) {
                        throw "The alias sequence has reached its six-digit limit."
                    }
                    $nonce = [Guid]::NewGuid().ToString("N").Substring(0, 12)
                    $candidate = ("{0}-{1}-{2:D6}-{3}@{4}" -f $AliasNamespace, $request.prefix, $sequence, $nonce, $request.domain).ToLowerInvariant()
                    $sequence++
                } while ($allAddresses -contains $candidate)

                $removeCount = [Math]::Max(0, $current.count - [int]$request.limit + 1)
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

                if ($toRemove.Count -gt 0) {
                    Set-Mailbox -Identity $mailbox.Identity -EmailAddresses @{
                        Add = "smtp:$candidate"
                        Remove = $toRemove
                    } -ErrorAction Stop
                }
                else {
                    Set-Mailbox -Identity $mailbox.Identity -EmailAddresses @{
                        Add = "smtp:$candidate"
                    } -ErrorAction Stop
                }

                $updated = Get-UserMailbox -UserId $request.userId
                Write-Envelope -Ok $true -Data (Get-AliasSet -Mailbox $updated -Limit $request.limit)
                break
            }

            "delete" {
                $mailbox = Get-UserMailbox -UserId $request.userId
                $current = Get-AliasSet -Mailbox $mailbox -Limit $request.limit
                $target = $request.address.ToString().ToLowerInvariant()
                $match = @($current.aliases | Where-Object { $_.address -eq $target })

                if ($match.Count -ne 1) {
                    throw "The address is not part of this app's managed alias set."
                }

                if ($mailbox.PrimarySmtpAddress.ToString().ToLowerInvariant() -eq $target) {
                    throw "The primary SMTP address cannot be removed."
                }

                Set-Mailbox -Identity $mailbox.Identity -EmailAddresses @{
                    Remove = "smtp:$target"
                } -ErrorAction Stop

                $updated = Get-UserMailbox -UserId $request.userId
                Write-Envelope -Ok $true -Data (Get-AliasSet -Mailbox $updated -Limit $request.limit)
                break
            }

            default {
                throw "Unsupported Exchange action."
            }
        }
    }
    finally {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    }
}
catch {
    $category = Get-ErrorCategory -Record $_
    Write-Envelope -Ok $false -ErrorMessage $_.Exception.Message -Category $category
    exit 1
}
