# Builds the Teams app package (appPackage/build/peek-a-boo.zip) from manifest.template.json + .env
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$pkg = Join-Path $root 'appPackage'
$out = Join-Path $pkg 'build'

$vars = @{}
foreach ($file in '.env', '.env.local') {
    $p = Join-Path $root $file
    if (-not (Test-Path $p)) { continue }
    foreach ($line in Get-Content $p) {
        if ($line -match '^\s*([A-Z_][A-Z0-9_]*)\s*=\s*(.*?)\s*$') { $vars[$Matches[1]] = $Matches[2].Trim('"') }
    }
}
foreach ($k in 'APP_HOST', 'TEAMS_APP_ID') {
    $fromEnv = [Environment]::GetEnvironmentVariable($k)
    if ($fromEnv) { $vars[$k] = $fromEnv }
    if (-not $vars[$k]) { throw "$k is not set (add it to .env)" }
}
if ($vars.TEAMS_APP_ID -notmatch '^[0-9a-fA-F-]{36}$' -or $vars.TEAMS_APP_ID -eq '00000000-0000-0000-0000-000000000000') {
    throw "TEAMS_APP_ID must be a real GUID. Generate one with: [guid]::NewGuid()"
}
$appHost = $vars.APP_HOST -replace '^https?://', '' -replace '/+$', ''
$domain = ($appHost -split '/')[0]

$manifest = (Get-Content (Join-Path $pkg 'manifest.template.json') -Raw).
    Replace('{{APP_HOST}}', $appHost).
    Replace('{{APP_DOMAIN}}', $domain).
    Replace('{{TEAMS_APP_ID}}', $vars.TEAMS_APP_ID)

New-Item -ItemType Directory -Force $out | Out-Null
$staging = Join-Path $out 'staging'
Remove-Item -Recurse -Force $staging -ErrorAction SilentlyContinue
New-Item -ItemType Directory $staging | Out-Null
Set-Content -Path (Join-Path $staging 'manifest.json') -Value $manifest -Encoding utf8NoBOM
Copy-Item (Join-Path $pkg 'color.png'), (Join-Path $pkg 'outline.png') $staging

$zip = Join-Path $out 'peek-a-boo.zip'
Remove-Item $zip -ErrorAction SilentlyContinue
Compress-Archive -Path (Join-Path $staging '*') -DestinationPath $zip
Remove-Item -Recurse -Force $staging
Write-Host "Teams app package: $zip (host: https://$appHost)"
