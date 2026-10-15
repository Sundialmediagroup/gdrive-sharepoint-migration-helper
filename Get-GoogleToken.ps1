#Requires -Version 7.2
# Signs in through the browser with a Desktop OAuth client (loopback + PKCE) and sets
# $env:GDRIVE_ACCESS_TOKEN for this PowerShell session. Never prints or stores the token.
# Run as .\Get-GoogleToken.ps1 (not in a new pwsh process) so the variable persists.
[CmdletBinding()]
param(
    [string]$ClientSecretPath,
    [string[]]$Scope = @('https://www.googleapis.com/auth/drive.metadata', 'https://www.googleapis.com/auth/drive.file'),
    [ValidateRange(30, 900)][int]$TimeoutSeconds = 300
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (!$ClientSecretPath) {
    # Default: the git-ignored .secrets folder at the project root.
    $found = @(Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot '.secrets') -Filter 'client_secret*.json' -ErrorAction SilentlyContinue)
    if ($found.Count -ne 1) { throw 'Put exactly one client_secret*.json in .secrets\ at the project root, or pass -ClientSecretPath.' }
    $ClientSecretPath = $found[0].FullName
}
$config = Get-Content -LiteralPath $ClientSecretPath -Raw | ConvertFrom-Json
if (!$config.PSObject.Properties['installed']) { throw 'Client secret must be for a Desktop app OAuth client.' }
$client = $config.installed

function ConvertTo-Base64Url([byte[]]$Bytes) { [Convert]::ToBase64String($Bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_') }
$verifier = ConvertTo-Base64Url ([Security.Cryptography.RandomNumberGenerator]::GetBytes(48))
$challenge = ConvertTo-Base64Url ([Security.Cryptography.SHA256]::HashData([Text.Encoding]::ASCII.GetBytes($verifier)))
$state = ConvertTo-Base64Url ([Security.Cryptography.RandomNumberGenerator]::GetBytes(24))

# Pick a free loopback port; Desktop clients accept any port on localhost.
$probe = [Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback, 0)
$probe.Start(); $port = $probe.LocalEndpoint.Port; $probe.Stop()
$redirect = "http://localhost:$port/"
$listener = [Net.HttpListener]::new()
$listener.Prefixes.Add($redirect)
$listener.Start()
try {
    $query = @{
        client_id = $client.client_id; redirect_uri = $redirect; response_type = 'code'
        scope = $Scope -join ' '; state = $state; code_challenge = $challenge; code_challenge_method = 'S256'
    }
    $authUrl = $client.auth_uri + '?' + (($query.GetEnumerator() | ForEach-Object { $_.Key + '=' + [uri]::EscapeDataString($_.Value) }) -join '&')
    Write-Host 'Opening the browser for Google sign-in. Use the Workspace test account.'
    Start-Process $authUrl
    $pending = $listener.GetContextAsync()
    if (!$pending.Wait($TimeoutSeconds * 1000)) { throw 'Timed out waiting for Google sign-in.' }
    $context = $pending.Result
    $params = [Web.HttpUtility]::ParseQueryString($context.Request.Url.Query)
    $message = if ($params['code'] -and $params['state'] -ceq $state) { 'Signed in. You can close this tab.' } else { 'Sign-in failed. Return to PowerShell.' }
    $page = [Text.Encoding]::UTF8.GetBytes("<html><body><p>$message</p></body></html>")
    $context.Response.ContentType = 'text/html; charset=utf-8'
    $context.Response.ContentLength64 = $page.Length
    $context.Response.OutputStream.Write($page, 0, $page.Length)
    $context.Response.Close()
} finally { $listener.Stop(); $listener.Close() }

if ($params['error']) { throw "Google sign-in failed: $($params['error'])." }
if ($params['state'] -cne $state) { throw 'Sign-in response state mismatch; token not requested.' }
$token = Invoke-RestMethod -Method Post -Uri $client.token_uri -Body @{
    code = $params['code']; client_id = $client.client_id; client_secret = $client.client_secret
    redirect_uri = $redirect; grant_type = 'authorization_code'; code_verifier = $verifier
}
$env:GDRIVE_ACCESS_TOKEN = $token.access_token
$granted = @($token.scope -split ' ')
$missing = @($Scope | Where-Object { $_ -notin $granted })
if ($missing.Count) { Write-Warning "Not granted: $($missing -join ', '). Re-run and tick every permission on the consent screen." }
Write-Host "GDRIVE_ACCESS_TOKEN set for this session; expires in about $([int]($token.expires_in / 60)) minutes."
Write-Host 'Remove it when finished: Remove-Item Env:\GDRIVE_ACCESS_TOKEN'
