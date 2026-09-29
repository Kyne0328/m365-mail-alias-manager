import { config } from "./config.js";

export function buildAdminConsentUrl(tenantId: string): string {
  const url = new URL(
    `https://login.microsoftonline.com/${tenantId}/adminconsent`
  );
  url.searchParams.set("client_id", config.clientId);
  url.searchParams.set("scope", "https://outlook.office365.com/.default");
  url.searchParams.set("redirect_uri", config.publicAppUrl);
  return url.toString();
}

export function buildSetupScript(): string {
  return `# Mail Alias Manager — one-time tenant setup
# Run as an Exchange Organization Management administrator after granting admin consent.

$ErrorActionPreference = "Stop"
$AppId = "${config.clientId}"
$DisplayName = "Mail Alias Manager"
$RecipientRole = "Mail Alias Manager - Recipients"
$DomainRole = "Mail Alias Manager - Domains"
$RoleGroup = "Mail Alias Manager"
$FallbackGroups = @("Recipient Management", "View-Only Organization Management")

if ($PSVersionTable.PSVersion.Major -lt 7) {
  throw "Run this script from PowerShell 7 (pwsh), not Windows PowerShell 5.1."
}

function Install-RequiredModule {
  param([Parameter(Mandatory = $true)][string]$Name)

  if (Get-Command Install-PSResource -ErrorAction SilentlyContinue) {
    Install-PSResource -Name $Name -Scope CurrentUser -TrustRepository -Quiet
  } else {
    Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber
  }
}

if (-not (Get-Module -ListAvailable Microsoft.Graph.Applications)) {
  Install-RequiredModule -Name "Microsoft.Graph.Applications"
}
Import-Module Microsoft.Graph.Applications -Force

$RequiredExchangeVersion = [version]"3.10.1"
$InstalledExchangeModule = Get-Module -ListAvailable ExchangeOnlineManagement |
  Sort-Object Version -Descending |
  Select-Object -First 1
if (-not $InstalledExchangeModule -or $InstalledExchangeModule.Version -lt $RequiredExchangeVersion) {
  Install-RequiredModule -Name "ExchangeOnlineManagement"
}
Import-Module ExchangeOnlineManagement -MinimumVersion $RequiredExchangeVersion -Force

Connect-MgGraph -Scopes "Application.Read.All" -NoWelcome
$EntraSp = Get-MgServicePrincipal -Filter "appId eq '$AppId'" | Select-Object -First 1
if (-not $EntraSp) {
  throw "Enterprise application not found. Grant tenant admin consent first, then rerun this script."
}

try {
  Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
} catch {
  Write-Warning "Interactive Exchange sign-in failed. Falling back to device-code sign-in."
  Connect-ExchangeOnline -ShowBanner:$false -Device
}

$OrgConfig = Get-OrganizationConfig -ErrorAction Stop
if ($OrgConfig.IsDehydrated) {
  Write-Host "Enabling Exchange organization customization..." -ForegroundColor Yellow
  Enable-OrganizationCustomization -Confirm:$false -ErrorAction Stop
}

$ExchangeSp = Get-ServicePrincipal -Identity $AppId -ErrorAction SilentlyContinue
if (-not $ExchangeSp) {
  New-ServicePrincipal -AppId $AppId -ObjectId $EntraSp.Id -DisplayName $DisplayName -ErrorAction Stop | Out-Null
  $ExchangeSp = Get-ServicePrincipal -Identity $AppId -ErrorAction Stop
}

$CustomRbacReady = $false
try {
  if (-not (Get-ManagementRole $RecipientRole -ErrorAction SilentlyContinue)) {
    New-ManagementRole -Name $RecipientRole -Parent "Mail Recipients" -ErrorAction Stop | Out-Null
    Get-ManagementRoleEntry "$RecipientRole\\*" -ErrorAction Stop |
      Where-Object Name -NotIn @("Get-Mailbox", "Set-Mailbox") |
      ForEach-Object { Remove-ManagementRoleEntry $_.Identity -Confirm:$false -ErrorAction Stop }
  }

  # Keep existing installations current when the app adds a narrowly-scoped
  # mailbox property. ExtensionCustomAttribute5 stores only m365am:v2:* values
  # and is modified with Add/Remove so unrelated tenant values are preserved.
  Set-ManagementRoleEntry "$RecipientRole\\Get-Mailbox" -Parameters Filter,ResultSize -ErrorAction Stop
  Set-ManagementRoleEntry "$RecipientRole\\Set-Mailbox" -Parameters Identity,EmailAddresses,ExtensionCustomAttribute5 -ErrorAction Stop

  if (-not (Get-ManagementRole $DomainRole -ErrorAction SilentlyContinue)) {
    New-ManagementRole -Name $DomainRole -Parent "View-Only Configuration" -ErrorAction Stop | Out-Null
    Get-ManagementRoleEntry "$DomainRole\\*" -ErrorAction Stop |
      Where-Object Name -ne "Get-AcceptedDomain" |
      ForEach-Object { Remove-ManagementRoleEntry $_.Identity -Confirm:$false -ErrorAction Stop }
  }

  if (-not (Get-RoleGroup $RoleGroup -ErrorAction SilentlyContinue)) {
    New-RoleGroup -Name $RoleGroup -Roles $RecipientRole,$DomainRole -Description "Least-privilege permissions for Mail Alias Manager." -ErrorAction Stop | Out-Null
  }

  $ExistingMember = Get-RoleGroupMember $RoleGroup -ResultSize Unlimited -ErrorAction Stop |
    Where-Object { $_.Identity -eq $ExchangeSp.Identity -or $_.Name -eq $ExchangeSp.ObjectId }

  if (-not $ExistingMember) {
    Add-RoleGroupMember -Identity $RoleGroup -Member $ExchangeSp.Identity -BypassSecurityGroupManagerCheck -ErrorAction Stop
  }

  $CustomRbacReady = $true
} catch {
  Write-Warning ("Least-privilege custom role creation is unavailable in this tenant: " + $_.Exception.Message)
  Write-Warning "Falling back to Microsoft's built-in Recipient Management and View-Only Organization Management role groups."
}

if (-not $CustomRbacReady) {
  foreach ($GroupName in $FallbackGroups) {
    $Member = Get-RoleGroupMember $GroupName -ResultSize Unlimited -ErrorAction Stop |
      Where-Object { $_.Identity -eq $ExchangeSp.Identity -or $_.Name -eq $ExchangeSp.ObjectId }

    if (-not $Member) {
      Add-RoleGroupMember -Identity $GroupName -Member $ExchangeSp.Identity -BypassSecurityGroupManagerCheck -ErrorAction Stop
    }
  }

  $RecipientFallback = Get-RoleGroupMember "Recipient Management" -ResultSize Unlimited -ErrorAction Stop |
    Where-Object { $_.Identity -eq $ExchangeSp.Identity -or $_.Name -eq $ExchangeSp.ObjectId }
  $ViewFallback = Get-RoleGroupMember "View-Only Organization Management" -ResultSize Unlimited -ErrorAction Stop |
    Where-Object { $_.Identity -eq $ExchangeSp.Identity -or $_.Name -eq $ExchangeSp.ObjectId }

  if (-not $RecipientFallback -or -not $ViewFallback) {
    throw "Tenant setup verification failed. The service principal is not present in both fallback Exchange role groups."
  }
}

Write-Host ""
if ($CustomRbacReady) {
  Write-Host "Mail Alias Manager tenant setup is complete (least-privilege custom RBAC)." -ForegroundColor Green
} else {
  Write-Host "Mail Alias Manager tenant setup is complete (built-in Exchange role-group fallback)." -ForegroundColor Green
}
Write-Host "Exchange permission caches can take time to refresh after a new assignment."
`;
}
