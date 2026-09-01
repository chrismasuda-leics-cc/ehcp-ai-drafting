from app.settings import USE_MANAGED_IDENTITY, get_azure_credential
import os
import json
from dotenv import load_dotenv
from azure.core.exceptions import ResourceExistsError
from azure.storage.blob import BlobServiceClient

load_dotenv()


AZURE_STORAGE_CONNECTION_STRING = os.getenv(
    "AZURE_STORAGE_CONNECTION_STRING", "")
AZURE_STORAGE_ACCOUNT_URL = os.getenv("AZURE_STORAGE_ACCOUNT_URL", "")
AZURE_STORAGE_CONTAINER_NAME = os.getenv(
    "AZURE_STORAGE_CONTAINER_NAME", "ehcp-outputs")


def get_blob_service_client():
    """Return a BlobServiceClient using managed identity or connection string."""
    if USE_MANAGED_IDENTITY:
        if not AZURE_STORAGE_ACCOUNT_URL:
            raise ValueError(
                "AZURE_STORAGE_ACCOUNT_URL is not set. "
                "Add it to your .env file to use managed identity with blob storage. "
                "Example: https://<account>.blob.core.windows.net"
            )
        return BlobServiceClient(
            account_url=AZURE_STORAGE_ACCOUNT_URL,
            credential=get_azure_credential(),
        )
    if not AZURE_STORAGE_CONNECTION_STRING:
        raise ValueError(
            "AZURE_STORAGE_CONNECTION_STRING is not set. "
            "Add it to your .env file to enable blob storage."
        )
    return BlobServiceClient.from_connection_string(AZURE_STORAGE_CONNECTION_STRING)


def _ensure_container(blob_service_client):
    """Return the container, creating it safely under concurrent requests."""
    container_client = blob_service_client.get_container_client(
        AZURE_STORAGE_CONTAINER_NAME)
    if not container_client.exists():
        try:
            container_client.create_container()
        except ResourceExistsError:
            pass
    return container_client


def upload_json_to_blob(local_path: str, blob_name: str = None) -> str:
    """Upload a local JSON file to Azure Blob Storage.

    Args:
        local_path: Path to the local JSON file.
        blob_name: Optional blob name (defaults to the file's basename).

    Returns:
        The blob URL of the uploaded file.
    """
    if not is_blob_storage_enabled():
        return ""

    if blob_name is None:
        blob_name = os.path.basename(local_path)

    blob_service_client = get_blob_service_client()
    container_client = _ensure_container(blob_service_client)
    blob_client = container_client.get_blob_client(blob_name)

    with open(local_path, "rb") as f:
        blob_client.upload_blob(f, overwrite=True)

    return blob_client.url


def upload_bytes_to_blob(data: bytes, blob_name: str, content_type: str = "application/json") -> str:
    """Upload raw bytes to Azure Blob Storage.

    Args:
        data: Bytes to upload.
        blob_name: Name for the blob.
        content_type: MIME type for the blob.

    Returns:
        The blob URL of the uploaded file.
    """
    if not is_blob_storage_enabled():
        return ""

    from azure.storage.blob import ContentSettings

    blob_service_client = get_blob_service_client()
    container_client = _ensure_container(blob_service_client)
    blob_client = container_client.get_blob_client(blob_name)

    blob_client.upload_blob(
        data,
        overwrite=True,
        content_settings=ContentSettings(content_type=content_type),
    )

    return blob_client.url


def is_blob_storage_enabled() -> bool:
    """Return True if blob storage is configured (via connection string or managed identity)."""
    if USE_MANAGED_IDENTITY:
        return bool(AZURE_STORAGE_ACCOUNT_URL)
    return bool(AZURE_STORAGE_CONNECTION_STRING)


# ---------------------------------------------------------
# Session-scoped, path-derived blob helpers
#
# The backend runs on multiple replicas (Azure Container Apps scales to
# several instances with no sticky sessions).  Per-user artifacts are written
# to session-scoped local directories (temp/<session_id>/..., output/<session_id>/...).
# Because a user's requests can be load-balanced to different replicas, blob
# storage is used as the single shared source of truth.  The blob key is
# derived from the file's path *relative to the working directory*, so the
# session_id embedded in the path guarantees keys are unique per user and never
# collide across concurrent uploads.
# ---------------------------------------------------------

_CONTENT_TYPE_BY_EXT = {
    ".json": "application/json",
    ".txt": "text/plain",
    ".pdf": "application/pdf",
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
}


def blob_key_for_path(local_path: str) -> str:
    """Derive a stable, session-scoped blob key from a local file path.

    Uses the path relative to the current working directory so that
    session-scoped directories (temp/<sid>/..., output/<sid>/...) map to
    unique blob keys.  Falls back to the basename if the path resolves to
    something outside the working directory.
    """
    abspath = os.path.abspath(local_path)
    try:
        rel = os.path.relpath(abspath, os.getcwd())
    except ValueError:
        rel = os.path.basename(local_path)
    rel = rel.replace(os.sep, "/")
    if rel.startswith("../") or rel == "..":
        rel = os.path.basename(local_path)
    return rel


def upload_file_to_blob(local_path: str, content_type: str = None) -> str:
    """Upload a local file to blob storage using a session-scoped key derived
    from its path.  No-op (returns "") if blob storage is disabled or the file
    does not exist locally."""
    if not is_blob_storage_enabled():
        return ""
    if not os.path.exists(local_path):
        return ""

    from azure.storage.blob import ContentSettings

    blob_name = blob_key_for_path(local_path)
    if content_type is None:
        ext = os.path.splitext(local_path)[1].lower()
        content_type = _CONTENT_TYPE_BY_EXT.get(
            ext, "application/octet-stream")

    blob_service_client = get_blob_service_client()
    container_client = _ensure_container(blob_service_client)
    blob_client = container_client.get_blob_client(blob_name)

    with open(local_path, "rb") as f:
        blob_client.upload_blob(
            f,
            overwrite=True,
            content_settings=ContentSettings(content_type=content_type),
        )

    return blob_client.url


def download_file_from_blob(local_path: str) -> bool:
    """Download a blob (keyed by the local path) to ``local_path``.

    Returns True if the file was downloaded successfully, False otherwise."""
    if not is_blob_storage_enabled():
        return False

    blob_name = blob_key_for_path(local_path)
    blob_service_client = get_blob_service_client()
    container_client = blob_service_client.get_container_client(
        AZURE_STORAGE_CONTAINER_NAME)
    blob_client = container_client.get_blob_client(blob_name)

    try:
        if not blob_client.exists():
            return False
        os.makedirs(os.path.dirname(os.path.abspath(local_path))
                    or ".", exist_ok=True)
        with open(local_path, "wb") as f:
            f.write(blob_client.download_blob().readall())
        return True
    except Exception as exc:
        print(f"[blob] download failed for {blob_name}: {exc}")
        return False


def ensure_local_file(local_path: str) -> bool:
    """Ensure a file exists locally, pulling it from blob storage if missing.

    A previous pipeline step may have produced the file on a different replica,
    so when it is absent on the current replica it is retrieved from blob using
    the session-scoped key.  Returns True if the file is available locally
    afterwards."""
    if os.path.exists(local_path):
        return True
    return download_file_from_blob(local_path)


def delete_blob(blob_name: str) -> bool:
    """Delete a blob from Azure Blob Storage.

    Args:
        blob_name: Name of the blob to delete.

    Returns:
        True if deleted, False if not found or blob storage is disabled.
    """
    if not is_blob_storage_enabled():
        return False

    blob_service_client = get_blob_service_client()
    container_client = blob_service_client.get_container_client(
        AZURE_STORAGE_CONTAINER_NAME)
    blob_client = container_client.get_blob_client(blob_name)

    if blob_client.exists():
        blob_client.delete_blob()
        return True
    return False
