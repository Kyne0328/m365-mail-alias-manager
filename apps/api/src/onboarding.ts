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

if (-not (Get-Module -ListAvailable Microsoft.Graph.Applications)) {
  Install-Module Microsoft.Graph.Applications -Scope CurrentUser -Force
}
$RequiredExchangeVersion = [version]"3.10.1"
$InstalledExchangeModule = Get-Module -ListAvailable ExchangeOnlineManagement |
  Sort-Object Version -Descending |
  Select-Object -First 1
if (-not $InstalledExchangeModule -or $InstalledExchangeModule.Version -lt $RequiredExchangeVersion) {
  Install-Module ExchangeOnlineManagement -MinimumVersion $RequiredExchangeVersion -Scope CurrentUser -Force
}

Connect-MgGraph -Scopes "Application.Read.All"
$EntraSp = Get-MgServicePrincipal -Filter "appId eq '$AppId'" | Select-Object -First 1
if (-not $EntraSp) {
  throw "Enterprise application not found. Grant tenant admin consent first, then rerun this script."
}

Connect-ExchangeOnline -ShowBanner:$false

$ExchangeSp = Get-ServicePrincipal -Identity $AppId -ErrorAction SilentlyContinue
if (-not $ExchangeSp) {
  New-ServicePrincipal -AppId $AppId -ObjectId $EntraSp.Id -DisplayName $DisplayName | Out-Null
  $ExchangeSp = Get-ServicePrincipal -Identity $AppId
}

if (-not (Get-ManagementRole $RecipientRole -ErrorAction SilentlyContinue)) {
  New-ManagementRole -Name $RecipientRole -Parent "Mail Recipients" | Out-Null
  Get-ManagementRoleEntry "$RecipientRole\\*" |
    Where-Object Name -NotIn @("Get-Mailbox", "Set-Mailbox") |
    ForEach-Object { Remove-ManagementRoleEntry $_.Identity -Confirm:$false }

  Set-ManagementRoleEntry "$RecipientRole\\Get-Mailbox" -Parameters Filter,ResultSize
  Set-ManagementRoleEntry "$RecipientRole\\Set-Mailbox" -Parameters Identity,EmailAddresses
}

if (-not (Get-ManagementRole $DomainRole -ErrorAction SilentlyContinue)) {
  New-ManagementRole -Name $DomainRole -Parent "View-Only Configuration" | Out-Null
  Get-ManagementRoleEntry "$DomainRole\\*" |
    Where-Object Name -ne "Get-AcceptedDomain" |
    ForEach-Object { Remove-ManagementRoleEntry $_.Identity -Confirm:$false }
}

if (-not (Get-RoleGroup $RoleGroup -ErrorAction SilentlyContinue)) {
  New-RoleGroup -Name $RoleGroup -Roles $RecipientRole,$DomainRole -Description "Least-privilege permissions for Mail Alias Manager." | Out-Null
}

$ExistingMember = Get-RoleGroupMember $RoleGroup -ResultSize Unlimited |
  Where-Object { $_.ExternalDirectoryObjectId -eq $EntraSp.Id -or $_.Name -eq $DisplayName }

if (-not $ExistingMember) {
  Add-RoleGroupMember -Identity $RoleGroup -Member $ExchangeSp.Identity
}

Write-Host ""
Write-Host "Mail Alias Manager tenant setup is complete." -ForegroundColor Green
Write-Host "Exchange permission caches can take time to refresh after a new assignment."
`;
}
