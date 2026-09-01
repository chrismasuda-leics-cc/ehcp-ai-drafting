<#
.SYNOPSIS
    Deploy EHCP Document Processor backend + frontend to Azure Container Apps.

.DESCRIPTION
    - Creates (or updates) two Container Apps: ehcp-backend and ehcp-frontend.
    - Environment variables from your local .env file are NEVER baked into the image.
      Instead they are injected at runtime:
        • Plain values   → passed as --env-vars
        • Sensitive keys → stored as Container App secrets, then referenced via secretref:
    - The frontend receives the backend's internal FQDN automatically.

.PARAMETER ResourceGroup
    Azure Resource Group that contains the Container Apps Environment.

.PARAMETER Environment
    Name of the existing Container Apps Environment (managed environment).

.PARAMETER Tag
    Image tag to deploy (default: latest).

.EXAMPLE
    .\deploy-aca.ps1 -ResourceGroup "rg-ehcp" -Environment "cae-ehcp"
    .\deploy-aca.ps1 -ResourceGroup "rg-ehcp" -Environment "cae-ehcp" -Tag "1.0.0"
#>

param(
    [Parameter(Mandatory)][string]$ResourceGroup,
    [Parameter(Mandatory)][string]$Environment,
    [string]$Tag = "latest"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ── Configuration ──────────────────────────────────────────────────────────
$ACR_SERVER = "" # Your Azure Container Registry Login Server

$BACKEND_APP   = "ehcp-containerapp-backend"
$FRONTEND_APP  = "ehcp-containerapp-frontend"

$BACKEND_IMAGE  = "$ACR_SERVER/ehcp-backend:$Tag"
$FRONTEND_IMAGE = "$ACR_SERVER/ehcp-frontend:$Tag"

$BACKEND_ENV_FILE = Join-Path $PSScriptRoot "backend\.env"
$FRONTEND_ENV_FILE = Join-Path $PSScriptRoot "frontend\.env"

# ── Helper ─────────────────────────────────────────────────────────────────
function Write-Step([string]$msg) {
    Write-Host "`n==> $msg" -ForegroundColor Cyan
}

function Assert-Command([string]$name) {
    if (-not (Get-Command $name -ErrorAction SilentlyContinue)) {
        Write-Error "'$name' is not installed or not on PATH."
    }
}

function Read-EnvFile([string]$path) {
    <#
    Parse a .env file and return a hashtable.
    Skips blank lines and comments (#). Strips inline comments.
    #>
    $result = @{}
    if (-not (Test-Path $path)) {
        Write-Warning ".env file not found at $path — env vars must be set manually."
        return $result
    }
    Get-Content $path | ForEach-Object {
        $line = $_.Trim()
        if ($line -and -not $line.StartsWith("#")) {
            $idx = $line.IndexOf("=")
            if ($idx -gt 0) {
                $key   = $line.Substring(0, $idx).Trim()
                $value = $line.Substring($idx + 1).Trim()
                # Strip surrounding quotes if present
                $value = $value -replace '^["'']|["'']$', ''
                # Strip inline comment  (e.g.  VALUE=foo  # comment)
                $value = ($value -split '\s+#')[0].Trim()
                if ($key -and $value -and -not $value.StartsWith("<")) {
                    $result[$key] = $value
                }
            }
        }
    }
    return $result
}

# ── Secret key names (stored as Container App secrets, not plain env vars) ──
$SECRET_KEYS = @(
    "AZURE_OPENAI_API_KEY",
    "AZURE_DOCUMENT_INTELLIGENCE_KEY",
    "AZURE_STORAGE_CONNECTION_STRING"
)

# ── Pre-flight ─────────────────────────────────────────────────────────────
Assert-Command "az"

Write-Step "Checking Azure CLI login"
$account = az account show 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Error "Not logged in. Run 'az login' first."
}

# ── Read .env ──────────────────────────────────────────────────────────────
Write-Step "Reading backend\.env"
$envVars = Read-EnvFile $BACKEND_ENV_FILE

Write-Step "Reading frontend\.env"
$frontendEnvVars = Read-EnvFile $FRONTEND_ENV_FILE

# ── Build secret args ──────────────────────────────────────────────────────
# Secrets: --secrets key1=value1 key2=value2 ...
# Env var references: --env-vars KEY=secretref:key1 ...
$secretArgs   = @()
$secretRefEnv = @()

foreach ($key in $SECRET_KEYS) {
    $secretName = $key.ToLower().Replace("_", "-")   # ACA secret names must be lowercase-kebab
    if ($envVars.ContainsKey($key) -and $envVars[$key]) {
        $secretArgs   += "$secretName=$($envVars[$key])"
        $secretRefEnv += "$key=secretref:$secretName"
        $envVars.Remove($key)   # remove from plain env vars
    }
}

# ── Build plain env-var args ───────────────────────────────────────────────
$plainEnvArgs = @()
foreach ($kv in $envVars.GetEnumerator()) {
    $plainEnvArgs += "$($kv.Key)=$($kv.Value)"
}

$allEnvArgs = $plainEnvArgs + $secretRefEnv

# ── Deploy / update backend ────────────────────────────────────────────────
Write-Step "Deploying backend Container App: $BACKEND_APP"

# Check if the app already exists
$exists = az containerapp show `
    --name $BACKEND_APP `
    --resource-group $ResourceGroup `
    2>$null

if ($LASTEXITCODE -eq 0) {
    Write-Host "   Updating existing Container App..." -ForegroundColor Yellow

    # Update secrets first (if any)
    if ($secretArgs.Count -gt 0) {
        az containerapp secret set `
            --name $BACKEND_APP `
            --resource-group $ResourceGroup `
            --secrets @secretArgs
        if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
    }

    az containerapp update `
        --name $BACKEND_APP `
        --resource-group $ResourceGroup `
        --image $BACKEND_IMAGE `
        --replace-env-vars @allEnvArgs
} else {
    Write-Host "   Creating new Container App..." -ForegroundColor Green

    $createArgs = @(
        "containerapp", "create",
        "--name",           $BACKEND_APP,
        "--resource-group", $ResourceGroup,
        "--environment",    $Environment,
        "--image",          $BACKEND_IMAGE,
        "--registry-server",$ACR_SERVER,
        "--target-port",    "8000",
        "--ingress",        "internal",       # internal only — frontend talks to it directly
        "--min-replicas",   "1",
        "--max-replicas",   "3"
    )

    if ($secretArgs.Count -gt 0) {
        $createArgs += "--secrets"
        $createArgs += $secretArgs
    }

    if ($allEnvArgs.Count -gt 0) {
        $createArgs += "--env-vars"
        $createArgs += $allEnvArgs
    }

    & az @createArgs
}

if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

# ── Get backend internal FQDN ──────────────────────────────────────────────
Write-Step "Retrieving backend internal FQDN"
$backendFqdn = az containerapp show `
    --name $BACKEND_APP `
    --resource-group $ResourceGroup `
    --query "properties.configuration.ingress.fqdn" `
    --output tsv

$backendUrl = "https://$backendFqdn"
Write-Host "   Backend URL (internal): $backendUrl" -ForegroundColor White

# ── Deploy / update frontend ───────────────────────────────────────────────
Write-Step "Deploying frontend Container App: $FRONTEND_APP"

$frontendPlainEnvArgs = @()
foreach ($kv in $frontendEnvVars.GetEnumerator()) {
    # Skip BACKEND_URL from .env — we inject the dynamic internal FQDN below
    if ($kv.Key -eq "BACKEND_URL") { continue }
    $frontendPlainEnvArgs += "$($kv.Key)=$($kv.Value)"
}

$frontendEnv = $frontendPlainEnvArgs + @("BACKEND_URL=$backendUrl")

$existsFrontend = az containerapp show `
    --name $FRONTEND_APP `
    --resource-group $ResourceGroup `
    2>$null

if ($LASTEXITCODE -eq 0) {
    Write-Host "   Updating existing Container App..." -ForegroundColor Yellow

    az containerapp update `
        --name $FRONTEND_APP `
        --resource-group $ResourceGroup `
        --image $FRONTEND_IMAGE `
        --replace-env-vars @frontendEnv
} else {
    Write-Host "   Creating new Container App..." -ForegroundColor Green

    az containerapp create `
        --name           $FRONTEND_APP `
        --resource-group $ResourceGroup `
        --environment    $Environment `
        --image          $FRONTEND_IMAGE `
        --registry-server $ACR_SERVER `
        --target-port    8501 `
        --ingress        "external" `
        --min-replicas   1 `
        --max-replicas   3 `
        --env-vars       @frontendEnv
}

if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

# ── Summary ────────────────────────────────────────────────────────────────
$frontendFqdn = az containerapp show `
    --name $FRONTEND_APP `
    --resource-group $ResourceGroup `
    --query "properties.configuration.ingress.fqdn" `
    --output tsv

Write-Host "`n✅ Deployment complete!" -ForegroundColor Green
Write-Host "   Backend  (internal) : $backendUrl"
Write-Host "   Frontend (public)   : https://$frontendFqdn" -ForegroundColor White
