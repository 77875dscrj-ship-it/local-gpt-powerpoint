$ErrorActionPreference = "Stop"

$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$certDir = Join-Path $root "certs"
$pfxPath = Join-Path $certDir "localhost.pfx"
$cerPath = Join-Path $certDir "localhost.cer"
$friendly = "Local GPT PowerPoint localhost"
$password = ConvertTo-SecureString "local-gpt-powerpoint" -AsPlainText -Force

New-Item -ItemType Directory -Force -Path $certDir | Out-Null

if (Test-Path -LiteralPath $pfxPath) {
  exit 0
}

$existing = Get-ChildItem Cert:\CurrentUser\My -ErrorAction SilentlyContinue |
  Where-Object { $_.FriendlyName -eq $friendly -and $_.NotAfter -gt (Get-Date).AddMonths(1) } |
  Select-Object -First 1

if ($null -eq $existing) {
  $existing = New-SelfSignedCertificate `
    -DnsName "localhost" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -FriendlyName $friendly `
    -KeyExportPolicy Exportable `
    -NotAfter (Get-Date).AddYears(3)
}

Export-PfxCertificate -Cert $existing -FilePath $pfxPath -Password $password | Out-Null
Export-Certificate -Cert $existing -FilePath $cerPath | Out-Null

$rootStoreHasCert = Get-ChildItem Cert:\CurrentUser\Root -ErrorAction SilentlyContinue |
  Where-Object { $_.Thumbprint -eq $existing.Thumbprint } |
  Select-Object -First 1

if ($null -eq $rootStoreHasCert) {
  Import-Certificate -FilePath $cerPath -CertStoreLocation Cert:\CurrentUser\Root | Out-Null
}
