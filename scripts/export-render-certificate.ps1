$ErrorActionPreference = "Stop"

$subject = "CN=Mail Alias Manager Exchange App"
$cert = Get-ChildItem Cert:\CurrentUser\My |
    Where-Object { $_.Subject -eq $subject -and $_.HasPrivateKey } |
    Sort-Object NotAfter -Descending |
    Select-Object -First 1

if (-not $cert) {
    throw "Could not find a certificate with subject '$subject' and an exportable private key."
}

$outDir = Join-Path (Split-Path $PSScriptRoot -Parent) ".local-secrets"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

$pfxPath = Join-Path $outDir "mail-alias-manager-exchange.pfx"
$password = Read-Host "Choose a PFX password for the Render secret" -AsSecureString

Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $password | Out-Null

$base64 = [Convert]::ToBase64String([IO.File]::ReadAllBytes($pfxPath))
Set-Clipboard -Value $base64

Write-Host ""
Write-Host "PFX exported successfully." -ForegroundColor Green
Write-Host "PFX path: $pfxPath"
Write-Host "Certificate thumbprint: $($cert.Thumbprint)"
Write-Host "The PFX base64 value is now in your clipboard."
Write-Host "Paste it directly into Render as EXCHANGE_CERTIFICATE_BASE64."
Write-Host "Use the same password you just entered for EXCHANGE_CERTIFICATE_PASSWORD."
Write-Host ""
Write-Host "Do not paste either secret into chat or commit the .local-secrets folder."
