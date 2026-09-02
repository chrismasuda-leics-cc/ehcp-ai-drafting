<#
.SYNOPSIS
    Full infrastructure deployment for EHCP Drafting Application.

.DESCRIPTION
    Provisions ALL Azure resources from scratch and deploys the application:
      1. Resource Group
      2. Azure Container Registry (ACR)
      3. Microsoft Foundry resource (with GPT deployment)
      4. Azure Document Intelligence
      5. Azure Storage Account + Blob Container
      6. Azure Cosmos DB (database + containers)
      7. Azure Container Apps Environment
      8. Build & push Docker images to ACR
      9. Deploy backend + frontend Container Apps with all env vars

    This is a ONE-TIME setup script. For subsequent code deployments, use
    build-push.ps1 and deploy-aca.ps1 instead.

    Replace {org name prefix} with shortname of your org.

.PARAMETER SubscriptionId
    Azure Subscription ID.

.PARAMETER ResourceGroup
    Resource Group name to create/use.

.PARAMETER Location
    Azure region (default: uksouth).

.PARAMETER Prefix
    Naming prefix for all resources.

.PARAMETER OpenAIModelName
    OpenAI model to deploy (default: gpt-5.2).

.PARAMETER SkipInfra
    Skip infrastructure creation and only build + deploy apps.

.EXAMPLE
    .\deploy-infrastructure.ps1 `
        -SubscriptionId "70195374-ae10-4ee7-a46d-67b8fd22d38c" `
        -ResourceGroup  "rg-ehcpdrafting-{org name prefix}" `
        -Location       "swedencentral" `
        -Prefix         "ehcp{org name prefix}"

    .\deploy-infrastructure.ps1 `
        -SubscriptionId "70195374-ae10-4ee7-a46d-67b8fd22d38c" `
        -ResourceGroup  "rg-ehcpdrafting-{org name prefix}" `
        -SkipInfra
#>

param(
    [Parameter(Mandatory)][string]$SubscriptionId,
    [Parameter(Mandatory)][string]$ResourceGroup,
    [string]$Location        = "swedencentral",
    [string]$Prefix          = "ehcp-{org name prefix}",
    [string]$OpenAIModelName = "gpt-5.2",
    [switch]$SkipInfra
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
if ($null -ne (Get-Variable -Name PSNativeCommandUseErrorActionPreference -ErrorAction SilentlyContinue)) {
    $PSNativeCommandUseErrorActionPreference = $false
}

# ══════════════════════════════════════════════════════════════════════════════
# RESOURCE NAMING
# ══════════════════════════════════════════════════════════════════════════════
$ACR_NAME           = ("${Prefix}draftregistry".ToLower() -replace '[^a-z0-9]', '')
if ($ACR_NAME.Length -gt 50) { $ACR_NAME = $ACR_NAME.Substring(0, 50) }
$ACR_SERVER         = "$ACR_NAME.azurecr.io"
$FOUNDRY_NAME        = "${Prefix}-foundry"
$OPENAI_DEPLOYMENT   = "gpt-5.2"
$DOC_INTEL_NAME      = "${Prefix}-docintel"
$STORAGE_ACCOUNT     = "${Prefix}storage"
$STORAGE_CONTAINER   = "ehcp-outputs"
$COSMOS_ACCOUNT      = "${Prefix}-cosmos"
$COSMOS_DATABASE     = "ehcp-audit"
$COSMOS_CONTAINER_ACTIVITY = "activity-logs"
$COSMOS_CONTAINER_JOB      = "job-logs"
$CAE_NAME            = "${Prefix}-environment"
$BACKEND_APP         = "${Prefix}-backend"
$FRONTEND_APP        = "${Prefix}-frontend"
$LOG_ANALYTICS       = "${Prefix}-logs"

# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════
function Write-Step([string]$msg) {
    Write-Host "`n$('='*60)" -ForegroundColor Cyan
    Write-Host "  $msg" -ForegroundColor Cyan
    Write-Host "$('='*60)" -ForegroundColor Cyan
}

function Write-SubStep([string]$msg) {
    Write-Host "  -> $msg" -ForegroundColor Yellow
}

function Write-Done([string]$msg) {
    Write-Host "  ✅ $msg" -ForegroundColor Green
}

function Test-ContainerAppExists([string]$name, [string]$resourceGroup) {
    $count = az containerapp list --resource-group $resourceGroup --query "[?name=='$name'] | length(@)" -o tsv 2>$null
    return ($LASTEXITCODE -eq 0 -and [int]$count -gt 0)
}

# ══════════════════════════════════════════════════════════════════════════════
# PRE-FLIGHT
# ══════════════════════════════════════════════════════════════════════════════
Write-Step "Pre-flight checks"

if (-not (Get-Command "az" -ErrorAction SilentlyContinue)) {
    Write-Error "Azure CLI (az) is not installed. Install from https://aka.ms/installazurecli"
}

$account = az account show 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Error "Not logged in. Run 'az login' first."
}

Write-SubStep "Setting subscription: $SubscriptionId"
az account set --subscription $SubscriptionId
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
Write-Done "Subscription set"

# ══════════════════════════════════════════════════════════════════════════════
# 1. RESOURCE GROUP
# ══════════════════════════════════════════════════════════════════════════════
if (-not $SkipInfra) {

Write-Step "1. Resource Group: $ResourceGroup"
az group create --name $ResourceGroup --location $Location -o none
Write-Done "Resource Group ready"

# ══════════════════════════════════════════════════════════════════════════════
# 2. AZURE CONTAINER REGISTRY
# ══════════════════════════════════════════════════════════════════════════════
Write-Step "2. Azure Container Registry: $ACR_NAME"

$acrInRg = az acr list --resource-group $ResourceGroup --query "[0].name" -o tsv 2>$null
if ($LASTEXITCODE -eq 0 -and $acrInRg) {
    $ACR_NAME = $acrInRg
    $ACR_SERVER = "$ACR_NAME.azurecr.io"
    Write-SubStep "Using existing ACR in resource group: $ACR_NAME"
} else {
    $acrNameAvailable = az acr check-name --name $ACR_NAME --query "nameAvailable" -o tsv
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

    if ($acrNameAvailable -ne "true") {
        $resourceToken = ($ResourceGroup.ToLower() -replace '[^a-z0-9]', '')
        if ([string]::IsNullOrWhiteSpace($resourceToken)) { $resourceToken = "rg" }

        $subToken = ($SubscriptionId -replace '-', '').ToLower()
        if ($subToken.Length -gt 6) { $subToken = $subToken.Substring(0, 6) }

        $altAcrName = ("{0}{1}acr{2}" -f $Prefix, $resourceToken, $subToken).ToLower() -replace '[^a-z0-9]', ''
        if ($altAcrName.Length -gt 50) { $altAcrName = $altAcrName.Substring(0, 50) }

        $altAvailable = az acr check-name --name $altAcrName --query "nameAvailable" -o tsv
        if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
        if ($altAvailable -ne "true") {
            Write-Error "No available ACR name found for '$ACR_NAME' or fallback '$altAcrName'."
        }

        Write-SubStep "ACR name '$ACR_NAME' is unavailable globally. Using '$altAcrName'."
        $ACR_NAME = $altAcrName
        $ACR_SERVER = "$ACR_NAME.azurecr.io"
    }

    az acr create `
        --name $ACR_NAME `
        --resource-group $ResourceGroup `
        --location $Location `
        --sku Basic `
        --admin-enabled true `
        -o none
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}
Write-Done "ACR created: $ACR_SERVER"

# ══════════════════════════════════════════════════════════════════════════════
# 3. MICROSOFT FOUNDRY
# ══════════════════════════════════════════════════════════════════════════════
Write-Step "3. Microsoft Foundry: $FOUNDRY_NAME"
$foundryExists = az cognitiveservices account show `
    --name $FOUNDRY_NAME --resource-group $ResourceGroup `
    --query "name" -o tsv 2>$null

if ($LASTEXITCODE -eq 0 -and -not [string]::IsNullOrWhiteSpace($foundryExists)) {
    Write-SubStep "Using existing Foundry resource: $FOUNDRY_NAME"
} else {
    az cognitiveservices account create `
        --name $FOUNDRY_NAME `
        --resource-group $ResourceGroup `
        --location $Location `
        --kind AIServices `
        --sku S0 `
        --allow-project-management true `
        --custom-domain $FOUNDRY_NAME `
        -o none
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
    Write-Done "Foundry resource created"
}

Write-SubStep "Deploying model: $OpenAIModelName -> $OPENAI_DEPLOYMENT"
az cognitiveservices account deployment create `
    --name $FOUNDRY_NAME `
    --resource-group $ResourceGroup `
    --deployment-name $OPENAI_DEPLOYMENT `
    --model-name $OpenAIModelName `
    --model-version "2025-12-11" `
    --model-format OpenAI `
    --sku-capacity 80 `
    --sku-name GlobalStandard `
    -o none
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
Write-Done "Model deployed: $OPENAI_DEPLOYMENT"

$OPENAI_ENDPOINT = az cognitiveservices account show `
    --name $FOUNDRY_NAME --resource-group $ResourceGroup `
    --query "properties.endpoint" -o tsv
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

$OPENAI_KEY = az cognitiveservices account keys list `
    --name $FOUNDRY_NAME --resource-group $ResourceGroup `
    --query "key1" -o tsv 2>$null
if ($LASTEXITCODE -ne 0) {
    $OPENAI_KEY = ""
    Write-SubStep "Foundry local auth key unavailable. Backend will use managed identity for Foundry access."
}

# ══════════════════════════════════════════════════════════════════════════════
# 4. AZURE DOCUMENT INTELLIGENCE
# ══════════════════════════════════════════════════════════════════════════════
Write-Step "4. Azure Document Intelligence: $DOC_INTEL_NAME"
az cognitiveservices account create `
    --name $DOC_INTEL_NAME `
    --resource-group $ResourceGroup `
    --location $Location `
    --kind FormRecognizer `
    --sku S0 `
    --custom-domain $DOC_INTEL_NAME `
    -o none
Write-Done "Document Intelligence created"

$DI_ENDPOINT = az cognitiveservices account show `
    --name $DOC_INTEL_NAME --resource-group $ResourceGroup `
    --query "properties.endpoint" -o tsv

$DI_KEY = az cognitiveservices account keys list `
    --name $DOC_INTEL_NAME --resource-group $ResourceGroup `
    --query "key1" -o tsv

# ══════════════════════════════════════════════════════════════════════════════
# 5. AZURE STORAGE ACCOUNT + BLOB CONTAINER
# ══════════════════════════════════════════════════════════════════════════════
Write-Step "5. Azure Storage Account: $STORAGE_ACCOUNT"
az storage account create `
    --name $STORAGE_ACCOUNT `
    --resource-group $ResourceGroup `
    --location $Location `
    --sku Standard_LRS `
    --kind StorageV2 `
    -o none
Write-Done "Storage account created"

$STORAGE_CONN_STRING = az storage account show-connection-string `
    --name $STORAGE_ACCOUNT --resource-group $ResourceGroup `
    --query "connectionString" -o tsv

Write-SubStep "Creating blob container: $STORAGE_CONTAINER"
az storage container create `
    --name $STORAGE_CONTAINER `
    --connection-string $STORAGE_CONN_STRING `
    -o none
Write-Done "Blob container created"

# ══════════════════════════════════════════════════════════════════════════════
# 6. AZURE COSMOS DB
# ══════════════════════════════════════════════════════════════════════════════
Write-Step "6. Azure Cosmos DB: $COSMOS_ACCOUNT"
az cosmosdb create `
    --name $COSMOS_ACCOUNT `
    --resource-group $ResourceGroup `
    --locations regionName=$Location failoverPriority=0 `
    --default-consistency-level Session `
    --kind GlobalDocumentDB `
    -o none
Write-Done "Cosmos DB account created"

$COSMOS_ENDPOINT = az cosmosdb show `
    --name $COSMOS_ACCOUNT --resource-group $ResourceGroup `
    --query "documentEndpoint" -o tsv

$COSMOS_KEY = az cosmosdb keys list `
    --name $COSMOS_ACCOUNT --resource-group $ResourceGroup `
    --query "primaryMasterKey" -o tsv

Write-SubStep "Creating database: $COSMOS_DATABASE"
az cosmosdb sql database create `
    --account-name $COSMOS_ACCOUNT `
    --resource-group $ResourceGroup `
    --name $COSMOS_DATABASE `
    -o none

Write-SubStep "Creating container: $COSMOS_CONTAINER_ACTIVITY (partition: /partitionKey)"
az cosmosdb sql container create `
    --account-name $COSMOS_ACCOUNT `
    --resource-group $ResourceGroup `
    --database-name $COSMOS_DATABASE `
    --name $COSMOS_CONTAINER_ACTIVITY `
    --partition-key-path "/partitionKey" `
    --throughput 400 `
    -o none

Write-SubStep "Creating container: $COSMOS_CONTAINER_JOB (partition: /partitionKey)"
az cosmosdb sql container create `
    --account-name $COSMOS_ACCOUNT `
    --resource-group $ResourceGroup `
    --database-name $COSMOS_DATABASE `
    --name $COSMOS_CONTAINER_JOB `
    --partition-key-path "/partitionKey" `
    --throughput 400 `
    -o none
Write-Done "Cosmos DB fully provisioned"

# ══════════════════════════════════════════════════════════════════════════════
# 7. LOG ANALYTICS + CONTAINER APPS ENVIRONMENT
# ══════════════════════════════════════════════════════════════════════════════
Write-Step "7. Container Apps Environment: $CAE_NAME"

Write-SubStep "Creating Log Analytics workspace: $LOG_ANALYTICS"
az monitor log-analytics workspace create `
    --workspace-name $LOG_ANALYTICS `
    --resource-group $ResourceGroup `
    --location $Location `
    -o none

$LOG_ANALYTICS_ID = az monitor log-analytics workspace show `
    --workspace-name $LOG_ANALYTICS --resource-group $ResourceGroup `
    --query "customerId" -o tsv

$LOG_ANALYTICS_KEY = az monitor log-analytics workspace get-shared-keys `
    --workspace-name $LOG_ANALYTICS --resource-group $ResourceGroup `
    --query "primarySharedKey" -o tsv

Write-SubStep "Creating Container Apps Environment"
az containerapp env create `
    --name $CAE_NAME `
    --resource-group $ResourceGroup `
    --location $Location `
    --logs-workspace-id $LOG_ANALYTICS_ID `
    --logs-workspace-key $LOG_ANALYTICS_KEY `
    -o none
Write-Done "Container Apps Environment ready"

} else {
    # ── SkipInfra: retrieve existing keys ──────────────────────────────────
    Write-Step "SkipInfra: Retrieving existing resource keys"

    $acrInRg = az acr list --resource-group $ResourceGroup --query "[0].name" -o tsv 2>$null
    if ($LASTEXITCODE -eq 0 -and $acrInRg) {
        $ACR_NAME = $acrInRg
        $ACR_SERVER = "$ACR_NAME.azurecr.io"
        Write-SubStep "Using existing ACR in resource group: $ACR_NAME"
    }

    $OPENAI_ENDPOINT = az cognitiveservices account show `
        --name $FOUNDRY_NAME --resource-group $ResourceGroup `
        --query "properties.endpoint" -o tsv 2>$null
    $OPENAI_KEY = az cognitiveservices account keys list `
        --name $FOUNDRY_NAME --resource-group $ResourceGroup `
        --query "key1" -o tsv 2>$null

    $DI_ENDPOINT = az cognitiveservices account show `
        --name $DOC_INTEL_NAME --resource-group $ResourceGroup `
        --query "properties.endpoint" -o tsv 2>$null
    $DI_KEY = az cognitiveservices account keys list `
        --name $DOC_INTEL_NAME --resource-group $ResourceGroup `
        --query "key1" -o tsv 2>$null

    $STORAGE_CONN_STRING = az storage account show-connection-string `
        --name $STORAGE_ACCOUNT --resource-group $ResourceGroup `
        --query "connectionString" -o tsv 2>$null

    $COSMOS_ENDPOINT = az cosmosdb show `
        --name $COSMOS_ACCOUNT --resource-group $ResourceGroup `
        --query "documentEndpoint" -o tsv 2>$null
    $COSMOS_KEY = az cosmosdb keys list `
        --name $COSMOS_ACCOUNT --resource-group $ResourceGroup `
        --query "primaryMasterKey" -o tsv 2>$null

    Write-Done "Keys retrieved"
}

# ══════════════════════════════════════════════════════════════════════════════
# 8. BUILD & PUSH DOCKER IMAGES
# ══════════════════════════════════════════════════════════════════════════════
Write-Step "8. Building & pushing Docker images to ACR"

$SCRIPT_DIR = $PSScriptRoot
$APP_ROOT = Join-Path $SCRIPT_DIR "ehcp_agent_final"
if (-not (Test-Path $APP_ROOT)) {
    $APP_ROOT = $SCRIPT_DIR
}

$BACKEND_DIR = Join-Path $APP_ROOT "backend"
$FRONTEND_DIR = Join-Path $APP_ROOT "frontend"

if (-not (Test-Path $BACKEND_DIR)) { Write-Error "Backend directory not found: $BACKEND_DIR" }
if (-not (Test-Path $FRONTEND_DIR)) { Write-Error "Frontend directory not found: $FRONTEND_DIR" }

Write-SubStep "Building backend image"
az acr build `
    --registry $ACR_NAME `
    --image "${Prefix}-backend:latest" `
    --file (Join-Path $BACKEND_DIR "Dockerfile") `
    $BACKEND_DIR `
    -o none
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
Write-Done "Backend image pushed"

Write-SubStep "Building frontend image"
az acr build `
    --registry $ACR_NAME `
    --image "${Prefix}-frontend:latest" `
    --file (Join-Path $FRONTEND_DIR "Dockerfile") `
    $FRONTEND_DIR `
    -o none
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
Write-Done "Frontend image pushed"

# ══════════════════════════════════════════════════════════════════════════════
# 9. DEPLOY BACKEND CONTAINER APP
# ══════════════════════════════════════════════════════════════════════════════
Write-Step "9. Deploying Backend: $BACKEND_APP"

# Secrets (sensitive values stored securely)
$backendSecrets = @()

# Environment variables
$backendEnvVars = @(
    "AZURE_OPENAI_ENDPOINT=$OPENAI_ENDPOINT",
    "AZURE_OPENAI_DEPLOYMENT=$OPENAI_DEPLOYMENT",
    "AZURE_OPENAI_API_VERSION=2025-04-01-preview",
    "MODEL_TEMPERATURE=0",
    "MODEL_MAX_TOKENS=30",
    "AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT=$DI_ENDPOINT",
    "AZURE_STORAGE_CONTAINER_NAME=$STORAGE_CONTAINER",
    "AUTH_ENABLED=false",
    "ENTRA_TENANT_ID=",
    "ENTRA_CLIENT_ID=",
    "AUDIT_LOG_ENABLED=false",
    "BACKEND_HOST=0.0.0.0",
    "BACKEND_PORT=8000",
    "BACKEND_WORKERS=4"
)

$useManagedIdentity = [string]::IsNullOrWhiteSpace($OPENAI_KEY) -or [string]::IsNullOrWhiteSpace($DI_KEY) -or [string]::IsNullOrWhiteSpace($STORAGE_CONN_STRING)

if ($useManagedIdentity) {
    Write-SubStep "One or more service keys are unavailable. Configuring backend for managed identity."
    $backendEnvVars += "USE_MANAGED_IDENTITY=true"
    $backendEnvVars += "AZURE_STORAGE_ACCOUNT_URL=https://$STORAGE_ACCOUNT.blob.core.windows.net"
    $backendEnvVars += "COSMOS_DB_ENDPOINT=$COSMOS_ENDPOINT"
    $backendEnvVars += "COSMOS_DB_DATABASE=$COSMOS_DATABASE"
    $backendEnvVars += "COSMOS_DB_CONTAINER=$COSMOS_CONTAINER_ACTIVITY"
    $backendEnvVars += "COSMOS_DB_JOB_CONTAINER=$COSMOS_CONTAINER_JOB"
} else {
    $backendSecrets += "openai-api-key=$OPENAI_KEY"
    $backendSecrets += "doc-intel-key=$DI_KEY"
    $backendSecrets += "storage-conn-string=$STORAGE_CONN_STRING"
    if (-not [string]::IsNullOrWhiteSpace($COSMOS_KEY)) {
        $backendSecrets += "cosmos-db-key=$COSMOS_KEY"
        $backendEnvVars += "COSMOS_DB_KEY=secretref:cosmos-db-key"
    }

    $backendEnvVars += "USE_MANAGED_IDENTITY=false"
    $backendEnvVars += "AZURE_OPENAI_API_KEY=secretref:openai-api-key"
    $backendEnvVars += "AZURE_DOCUMENT_INTELLIGENCE_KEY=secretref:doc-intel-key"
    $backendEnvVars += "AZURE_STORAGE_CONNECTION_STRING=secretref:storage-conn-string"
    $backendEnvVars += "COSMOS_DB_ENDPOINT=$COSMOS_ENDPOINT"
    $backendEnvVars += "COSMOS_DB_DATABASE=$COSMOS_DATABASE"
    $backendEnvVars += "COSMOS_DB_CONTAINER=$COSMOS_CONTAINER_ACTIVITY"
    $backendEnvVars += "COSMOS_DB_JOB_CONTAINER=$COSMOS_CONTAINER_JOB"
}

if (Test-ContainerAppExists -name $BACKEND_APP -resourceGroup $ResourceGroup) {
    Write-SubStep "Updating existing backend..."
    $backendSecretArgs = @(
        "containerapp", "secret", "set",
        "--name", $BACKEND_APP,
        "--resource-group", $ResourceGroup,
        "--secrets"
    ) + $backendSecrets + @("-o", "none")
    & az @backendSecretArgs
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

    $backendUpdateArgs = @(
        "containerapp", "update",
        "--name", $BACKEND_APP,
        "--resource-group", $ResourceGroup,
        "--image", "$ACR_SERVER/${Prefix}-backend:latest",
        "--replace-env-vars"
    ) + $backendEnvVars + @("-o", "none")
    & az @backendUpdateArgs
} else {
    Write-SubStep "Creating new backend Container App..."
    $backendCreateArgs = @(
        "containerapp", "create",
        "--name", $BACKEND_APP,
        "--resource-group", $ResourceGroup,
        "--environment", $CAE_NAME,
        "--image", "$ACR_SERVER/${Prefix}-backend:latest",
        "--registry-server", $ACR_SERVER,
        "--target-port", "8000",
        "--ingress", "internal",
        "--min-replicas", "1",
        "--max-replicas", "3",
        "--cpu", "2.0",
        "--memory", "4.0Gi"
    )

    if ($backendSecrets.Count -gt 0) {
        $backendCreateArgs += "--secrets"
        $backendCreateArgs += $backendSecrets
    }

    if ($backendEnvVars.Count -gt 0) {
        $backendCreateArgs += "--env-vars"
        $backendCreateArgs += $backendEnvVars
    }

    $backendCreateArgs += @("-o", "none")
    & az @backendCreateArgs
}
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
Write-Done "Backend deployed"

# ── Get backend internal FQDN ─────────────────────────────────────────────
$backendFqdn = az containerapp show `
    --name $BACKEND_APP --resource-group $ResourceGroup `
    --query "properties.configuration.ingress.fqdn" -o tsv
$backendUrl = "https://$backendFqdn"

# ══════════════════════════════════════════════════════════════════════════════
# 10. DEPLOY FRONTEND CONTAINER APP
# ══════════════════════════════════════════════════════════════════════════════
Write-Step "10. Deploying Frontend: $FRONTEND_APP"

$frontendEnvVars = @(
    "BACKEND_URL=$backendUrl",
    "AUTH_ENABLED=false",
    "ENTRA_TENANT_ID=",
    "ENTRA_CLIENT_ID=",
    "ENTRA_CLIENT_SECRET=",
    "ENTRA_SCOPE="
)

if (Test-ContainerAppExists -name $FRONTEND_APP -resourceGroup $ResourceGroup) {
    Write-SubStep "Updating existing frontend..."
    $frontendUpdateArgs = @(
        "containerapp", "update",
        "--name", $FRONTEND_APP,
        "--resource-group", $ResourceGroup,
        "--image", "$ACR_SERVER/${Prefix}-frontend:latest",
        "--replace-env-vars"
    ) + $frontendEnvVars + @("-o", "none")
    & az @frontendUpdateArgs
} else {
    Write-SubStep "Creating new frontend Container App..."
    $frontendCreateArgs = @(
        "containerapp", "create",
        "--name", $FRONTEND_APP,
        "--resource-group", $ResourceGroup,
        "--environment", $CAE_NAME,
        "--image", "$ACR_SERVER/${Prefix}-frontend:latest",
        "--registry-server", $ACR_SERVER,
        "--target-port", "8501",
        "--ingress", "external",
        "--min-replicas", "1",
        "--max-replicas", "3",
        "--cpu", "1.0",
        "--memory", "2.0Gi"
    )

    if ($frontendEnvVars.Count -gt 0) {
        $frontendCreateArgs += "--env-vars"
        $frontendCreateArgs += $frontendEnvVars
    }

    $frontendCreateArgs += @("-o", "none")
    & az @frontendCreateArgs
}
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
Write-Done "Frontend deployed"

# ══════════════════════════════════════════════════════════════════════════════
# SUMMARY
# ══════════════════════════════════════════════════════════════════════════════
$frontendFqdn = az containerapp show `
    --name $FRONTEND_APP --resource-group $ResourceGroup `
    --query "properties.configuration.ingress.fqdn" -o tsv

Write-Host ""
Write-Host "$('='*60)" -ForegroundColor Green
Write-Host "  DEPLOYMENT COMPLETE" -ForegroundColor Green
Write-Host "$('='*60)" -ForegroundColor Green
Write-Host ""
Write-Host "  RESOURCES CREATED:" -ForegroundColor White
Write-Host "  ────────────────────────────────────────────" -ForegroundColor DarkGray
Write-Host "  Resource Group          : $ResourceGroup"
Write-Host "  Container Registry      : $ACR_SERVER"
Write-Host "  Microsoft Foundry       : $OPENAI_ENDPOINT"
Write-Host "  Document Intelligence   : $DI_ENDPOINT"
Write-Host "  Storage Account         : $STORAGE_ACCOUNT"
Write-Host "  Cosmos DB               : $COSMOS_ENDPOINT"
Write-Host "  Container Apps Env      : $CAE_NAME"
Write-Host ""
Write-Host "  APPLICATION URLs:" -ForegroundColor White
Write-Host "  ────────────────────────────────────────────" -ForegroundColor DarkGray
Write-Host "  Frontend (public)       : https://$frontendFqdn" -ForegroundColor Cyan
Write-Host "  Backend  (internal)     : $backendUrl" -ForegroundColor Cyan
Write-Host ""
Write-Host "  NEXT STEPS:" -ForegroundColor White
Write-Host "  ────────────────────────────────────────────" -ForegroundColor DarkGray
Write-Host "  1. Create Entra ID App Registration (see EHCP_AppRegistration_Customer_Setup.md)"
Write-Host "  2. Update ENTRA_TENANT_ID, ENTRA_CLIENT_ID, ENTRA_CLIENT_SECRET in Container App env vars"
Write-Host "  3. Set ENTRA_REDIRECT_URI to: https://$frontendFqdn"
Write-Host "  4. Test the application at: https://$frontendFqdn"
Write-Host ""
Write-Host "  For subsequent deployments, use:" -ForegroundColor DarkGray
Write-Host "    az acr build --registry $ACR_NAME --image ${Prefix}-backend:latest --file backend/Dockerfile backend/"
Write-Host "    az acr build --registry $ACR_NAME --image ${Prefix}-frontend:latest --file frontend/Dockerfile frontend/"
Write-Host "    az containerapp update --name $BACKEND_APP --resource-group $ResourceGroup --image $ACR_SERVER/${Prefix}-backend:latest"
Write-Host "    az containerapp update --name $FRONTEND_APP --resource-group $ResourceGroup --image $ACR_SERVER/${Prefix}-frontend:latest"
Write-Host ""
