import os
from functools import lru_cache
from dotenv import load_dotenv

load_dotenv()

# ── Managed Identity Toggle ─────────────────────────────
USE_MANAGED_IDENTITY = os.getenv(
    "USE_MANAGED_IDENTITY", "false").lower() in ("true", "1", "yes")
# for user-assigned managed identity
AZURE_CLIENT_ID = os.getenv("AZURE_CLIENT_ID", "")

# Azure OpenAI
AZURE_OPENAI_ENDPOINT = os.getenv("AZURE_OPENAI_ENDPOINT")
# used only when managed identity is off
AZURE_OPENAI_API_KEY = os.getenv("AZURE_OPENAI_API_KEY")
AZURE_OPENAI_DEPLOYMENT = os.getenv("AZURE_OPENAI_DEPLOYMENT")
AZURE_OPENAI_API_VERSION = os.getenv("AZURE_OPENAI_API_VERSION")

# Azure OpenAI model settings
MODEL_TEMPERATURE = float(os.getenv("MODEL_TEMPERATURE", "0"))
MODEL_MAX_TOKENS = int(os.getenv("MODEL_MAX_TOKENS", "30"))

# Azure Document Intelligence
AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT = os.getenv(
    "AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT")
AZURE_DOCUMENT_INTELLIGENCE_KEY = os.getenv(
    "AZURE_DOCUMENT_INTELLIGENCE_KEY")  # used only when managed identity is off

# Azure Blob Storage
AZURE_STORAGE_CONNECTION_STRING = os.getenv(
    "AZURE_STORAGE_CONNECTION_STRING", "")  # used only when managed identity is off
# e.g. https://<account>.blob.core.windows.net
AZURE_STORAGE_ACCOUNT_URL = os.getenv("AZURE_STORAGE_ACCOUNT_URL", "")
AZURE_STORAGE_CONTAINER_NAME = os.getenv(
    "AZURE_STORAGE_CONTAINER_NAME", "ehcp-outputs")

# Paths
TEMPLATE_DOCX = "EHCP_LCC_Template.docx"
MAPPING_WORKBOOK = "ehcp_mapping.xlsx"
EXPECTED_OUTPUT_DOCX = "EHCP Output - Child A.docx"
TEMP_DIR = "temp"
OUTPUT_DIR = "output"

# Backend API
BACKEND_HOST = os.getenv("BACKEND_HOST", "0.0.0.0")
BACKEND_PORT = int(os.getenv("BACKEND_PORT", "8000"))
BACKEND_WORKERS = int(os.getenv("BACKEND_WORKERS", "4"))

# ── Microsoft Entra ID (Azure AD) Authentication ─────────
AUTH_ENABLED = os.getenv(
    "AUTH_ENABLED", "false").lower() in ("true", "1", "yes")
ENTRA_TENANT_ID = os.getenv("ENTRA_TENANT_ID", "")
# Backend API app registration client ID
ENTRA_CLIENT_ID = os.getenv("ENTRA_CLIENT_ID", "")
ENTRA_AUTHORITY = f"https://login.microsoftonline.com/{ENTRA_TENANT_ID}"

# ── Azure Cosmos DB (Audit Logging) ──────────────────────
AUDIT_LOG_ENABLED = os.getenv(
    "AUDIT_LOG_ENABLED", "false").lower() in ("true", "1", "yes")
COSMOS_DB_ENDPOINT = os.getenv("COSMOS_DB_ENDPOINT", "")
# used only when managed identity is off
COSMOS_DB_KEY = os.getenv("COSMOS_DB_KEY", "")
COSMOS_DB_DATABASE = os.getenv("COSMOS_DB_DATABASE", "ehcp-audit")
COSMOS_DB_CONTAINER = os.getenv("COSMOS_DB_CONTAINER", "activity-logs")


# ── Credential Helpers ──────────────────────────────────

@lru_cache(maxsize=1)
def get_azure_credential():
    """Return a DefaultAzureCredential (cached). Only imported when needed."""
    from azure.identity import DefaultAzureCredential
    kwargs = {}
    if AZURE_CLIENT_ID:
        kwargs["managed_identity_client_id"] = AZURE_CLIENT_ID
    return DefaultAzureCredential(**kwargs)


@lru_cache(maxsize=1)
def get_openai_token_provider():
    """Return a bearer-token provider for Azure OpenAI (cached)."""
    from azure.identity import get_bearer_token_provider
    return get_bearer_token_provider(
        get_azure_credential(),
        "https://cognitiveservices.azure.com/.default",
    )
