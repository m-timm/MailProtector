<#
.SYNOPSIS
  Builds the Report Missed Spam manifest for your hosting URL and (optionally) deploys it to client tenants.

.DESCRIPTION
  1. Replaces {{BASE_URL}} in manifest.template.xml and writes dist\manifest.xml.
  2. If -TenantDomains is supplied, connects to each tenant's Exchange Online via GDAP
     (delegated admin) and installs the add-in org-wide with New-App -OrganizationApp.

  Use -ManifestOnly to just build the manifest, then upload dist\manifest.xml yourself in the
  M365 admin center (Settings > Integrated apps > Upload custom apps > Office Add-in).

  NOTE: The New-App path is untested. Try it on ONE tenant first and confirm the button appears
  in both classic and new Outlook before looping across clients.

.EXAMPLE
  .\Deploy-AddIn.ps1 -BaseUrl https://spamreport.example.com -ManifestOnly

.EXAMPLE
  .\Deploy-AddIn.ps1 -BaseUrl https://spamreport.example.com `
      -UserPrincipalName you@yourmsp.com `
      -TenantDomains contoso.onmicrosoft.com
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory)]
  [ValidatePattern('^https://')]
  [string]$BaseUrl,

  [string[]]$TenantDomains,

  [string]$UserPrincipalName,

  [switch]$ManifestOnly
)

$ErrorActionPreference = 'Stop'
$root      = $PSScriptRoot
$template  = Join-Path $root 'manifest.template.xml'
$distDir   = Join-Path $root 'dist'
$manifest  = Join-Path $distDir 'manifest.xml'
$addInName = 'Report Missed Spam'

# --- Build manifest -----------------------------------------------------------
$BaseUrl = $BaseUrl.TrimEnd('/')
New-Item -ItemType Directory -Path $distDir -Force | Out-Null
$BaseOrigin = ([System.Uri]$BaseUrl).GetLeftPart([System.UriPartial]::Authority)
$raw     = Get-Content -Raw -Path $template
$version = [regex]::Match($raw, '<Version>([^<]+)</Version>').Groups[1].Value
$raw.Replace('{{BASE_URL}}', $BaseUrl).Replace('{{BASE_ORIGIN}}', $BaseOrigin).Replace('{{VERSION}}', $version) |
  Set-Content -Path $manifest -Encoding UTF8
[xml]$null = Get-Content -Raw -Path $manifest   # throws if not well-formed
Write-Host "Manifest written: $manifest" -ForegroundColor Green

if ($ManifestOnly -or -not $TenantDomains) { return }

# --- Deploy to tenants --------------------------------------------------------
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
  throw 'ExchangeOnlineManagement module is required: Install-Module ExchangeOnlineManagement'
}
if (-not $UserPrincipalName) { throw '-UserPrincipalName (your partner admin UPN) is required for delegated connections.' }

$bytes = [System.IO.File]::ReadAllBytes($manifest)

foreach ($tenant in $TenantDomains) {
  Write-Host "`n[$tenant]" -ForegroundColor Cyan
  try {
    Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -DelegatedOrganization $tenant -ShowBanner:$false

    $existing = Get-App -OrganizationApp -ErrorAction SilentlyContinue |
      Where-Object { $_.DisplayName -eq $addInName }
    if ($existing) {
      Write-Host '  Already installed - skipping. To push a manifest change, remove it and re-run, or re-upload in the admin center.' -ForegroundColor Yellow
      continue
    }

    New-App -OrganizationApp -FileData $bytes -DefaultStateForUser Enabled -ProvidedTo Everyone | Out-Null
    Write-Host '  Installed for everyone (users may need to restart Outlook).' -ForegroundColor Green
  }
  catch {
    Write-Warning "  Failed: $($_.Exception.Message)"
  }
  finally {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
  }
}
