<#
.SYNOPSIS
    Build Docker images for EHCP Document Processor (backend + frontend) and push to Azure ACR.

.DESCRIPTION
    1. Logs in to the Azure Container Registry using the Azure CLI.
    2. Builds the backend  image and tags it as <acr>/ehcp-backend:<tag>.
    3. Builds the frontend image and tags it as <acr>/ehcp-frontend:<tag>.
    4. Pushes both images to ACR.

.PARAMETER Tag
    Docker image tag (default: "latest").  Use a Git SHA or semver for production, e.g.:
        .\build-push.ps1 -Tag "1.0.0"
        .\build-push.ps1 -Tag (git rev-parse --short HEAD)

.EXAMPLE
    .\build-push.ps1
    .\build-push.ps1 -Tag "1.2.3"
#>

param(
    [string]$Tag = "latest"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ── Configuration ──────────────────────────────────────────────────────────
$ACR_NAME   = "" # Your Azure Container Registry name
$ACR_SERVER = "$ACR_NAME.azurecr.io"

$SCRIPT_DIR   = $PSScriptRoot
$BACKEND_DIR  = Join-Path $SCRIPT_DIR "backend"
$FRONTEND_DIR = Join-Path $SCRIPT_DIR "frontend"

$BACKEND_IMAGE  = "$ACR_SERVER/ehcp-backend:$Tag"
$FRONTEND_IMAGE = "$ACR_SERVER/ehcp-frontend:$Tag"

# ── Helper ─────────────────────────────────────────────────────────────────
function Write-Step([string]$msg) {
    Write-Host "`n==> $msg" -ForegroundColor Cyan
}

function Assert-Command([string]$name) {
    if (-not (Get-Command $name -ErrorAction SilentlyContinue)) {
        Write-Error "'$name' is not installed or not on PATH. Please install it first."
    }
}

# ── Pre-flight checks ──────────────────────────────────────────────────────
Assert-Command "docker"
Assert-Command "az"

# ── 1. ACR Login ───────────────────────────────────────────────────────────
Write-Step "Logging in to ACR: $ACR_SERVER"
az acr login --name $ACR_NAME
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

# ── 2. Build backend ───────────────────────────────────────────────────────
Write-Step "Building backend image → $BACKEND_IMAGE"
docker build `
    --file (Join-Path $BACKEND_DIR "Dockerfile") `
    --tag  $BACKEND_IMAGE `
    --tag  "$ACR_SERVER/ehcp-backend:latest" `
    $BACKEND_DIR
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

# ── 3. Build frontend ──────────────────────────────────────────────────────
Write-Step "Building frontend image → $FRONTEND_IMAGE"
docker build `
    --file (Join-Path $FRONTEND_DIR "Dockerfile") `
    --tag  $FRONTEND_IMAGE `
    --tag  "$ACR_SERVER/ehcp-frontend:latest" `
    $FRONTEND_DIR
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

# ── 4. Push backend ────────────────────────────────────────────────────────
Write-Step "Pushing backend image"
docker push $BACKEND_IMAGE
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

if ($Tag -ne "latest") {
    docker push "$ACR_SERVER/ehcp-backend:latest"
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}

# ── 5. Push frontend ───────────────────────────────────────────────────────
Write-Step "Pushing frontend image"
docker push $FRONTEND_IMAGE
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

if ($Tag -ne "latest") {
    docker push "$ACR_SERVER/ehcp-frontend:latest"
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}

# ── Summary ────────────────────────────────────────────────────────────────
Write-Host "`n✅ Done!" -ForegroundColor Green
Write-Host "   Backend  : $BACKEND_IMAGE" -ForegroundColor White
Write-Host "   Frontend : $FRONTEND_IMAGE" -ForegroundColor White
