# Provisions the Work IQ service principal in the tenant.
# Run as a Global Administrator / Application Administrator.

$ErrorActionPreference = 'Stop'
$WorkIqAppId = 'fdcc1f02-fc51-4226-8753-f668596af7f7'
$TenantId    = '1a632370-87d8-4768-a8d0-7a9a728dd03d'

Write-Host 'Connecting to Microsoft Graph (device code)...' -ForegroundColor Cyan
Connect-MgGraph -UseDeviceCode -TenantId $TenantId -Scopes 'Application.ReadWrite.All' -NoWelcome

$ctx = Get-MgContext
Write-Host ("Signed in as: {0}" -f $ctx.Account) -ForegroundColor Green
Write-Host ("Scopes: {0}" -f ($ctx.Scopes -join ', '))

$existing = Get-MgServicePrincipal -Filter "appId eq '$WorkIqAppId'" -ErrorAction SilentlyContinue
if ($existing) {
    Write-Host ("Work IQ service principal already exists. objectId={0}" -f $existing.Id) -ForegroundColor Yellow
} else {
    Write-Host 'Creating Work IQ service principal...' -ForegroundColor Cyan
    $sp = New-MgServicePrincipal -AppId $WorkIqAppId
    Write-Host ("Created. objectId={0} displayName={1}" -f $sp.Id, $sp.DisplayName) -ForegroundColor Green
}
