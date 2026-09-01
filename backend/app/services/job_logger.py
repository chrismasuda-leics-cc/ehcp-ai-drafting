"""
Job-level Audit Logger for EHCP Document Processor.

Logs a single comprehensive record per processing run to Cosmos DB,
capturing the full lifecycle: upload → analyse → create EHCP.
"""
import os
import uuid
from datetime import datetime, timezone
from typing import Optional

from app.settings import (
    USE_MANAGED_IDENTITY,
    COSMOS_DB_ENDPOINT,
    COSMOS_DB_KEY,
    COSMOS_DB_DATABASE,
    AUDIT_LOG_ENABLED,
)

# Job logs go to a separate container
COSMOS_DB_JOB_CONTAINER = os.getenv("COSMOS_DB_JOB_CONTAINER", "job-logs")

_cosmos_client = None
_container_client = None


def _get_container_client():
    """Lazy-init Cosmos DB container client for job logs."""
    global _cosmos_client, _container_client
    if _container_client is not None:
        return _container_client
    if not COSMOS_DB_ENDPOINT:
        return None
    from azure.cosmos.aio import CosmosClient
    if USE_MANAGED_IDENTITY:
        from azure.identity.aio import DefaultAzureCredential
        credential = DefaultAzureCredential()
        _cosmos_client = CosmosClient(
            COSMOS_DB_ENDPOINT, credential=credential)
    else:
        _cosmos_client = CosmosClient(
            COSMOS_DB_ENDPOINT, credential=COSMOS_DB_KEY)
    database = _cosmos_client.get_database_client(COSMOS_DB_DATABASE)
    _container_client = database.get_container_client(COSMOS_DB_JOB_CONTAINER)
    return _container_client


def create_job_record(user: Optional[dict] = None, session_id: str = None, job_id: str = None) -> dict:
    """Create a new job record with a unique ID and initial fields.

    ``job_id`` may be supplied by the caller (generated once per case on the
    frontend) so the same case always maps to one record.  The Cosmos document
    ``id`` is kept identical to ``job_id`` so the record can be fetched with a
    strongly-consistent point read, avoiding query indexing lag that could
    otherwise create duplicate job records."""
    now = datetime.now(timezone.utc).isoformat()
    jid = job_id or str(uuid.uuid4())
    return {
        "id": jid,
        "partitionKey": session_id or "anonymous",
        "job_id": jid,
        "user_name": user.get("name", "anonymous") if user else "anonymous",
        "user_email": (user.get("preferred_username") or user.get("upn") or user.get("unique_name") or user.get("email") or "") if user else "",
        "user_oid": user.get("oid", "") if user else "",
        "session_id": session_id or "",

        # Upload phase
        "upload_datetime": now,
        "upload_documents": [],  # [{name, size, doc_category}]
        "upload_failures": [],   # [{name, reason}]

        # Analyse phase
        "analyse_start_datetime": None,
        "analyse_completed_datetime": None,

        # Create EHCP phase
        "create_ehcp_completed_datetime": None,
        "create_ehcp_document_name": None,
        "create_ehcp_document_size": None,
        "create_ehcp_file_path": None,  # storage account path

        # Status
        # uploaded | analysing | analysed | creating | completed | error | timeout
        "status": "uploaded",
        "error": "",

        # Token usage
        "token_usage": {
            "prompt_tokens": 0,
            "completion_tokens": 0,
            "total_tokens": 0,
            "prompt_tokens_display": "0",
            "completion_tokens_display": "0",
            "total_tokens_display": "0",
        },

        # Per-document completeness report
        "completeness_report": [],
        # [{doc_category, total_fields, complete_fields, missing_fields, missing_field_names}]
    }


def _format_size(size_bytes: int) -> str:
    """Convert bytes to human-readable string (e.g. 217.5 KB)."""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    elif size_bytes < 1024 * 1024 * 1024:
        return f"{size_bytes / (1024 * 1024):.1f} MB"
    else:
        return f"{size_bytes / (1024 * 1024 * 1024):.1f} GB"


def _format_tokens(count: int) -> str:
    """Convert token count to human-readable string (e.g. 12.5K)."""
    if count < 1000:
        return str(count)
    elif count < 1_000_000:
        return f"{count / 1000:.1f}K"
    else:
        return f"{count / 1_000_000:.2f}M"


def add_upload_document(job: dict, filename: str, size_bytes: int, doc_category: str):
    """Add an uploaded document entry to the job record."""
    job["upload_documents"].append({
        "name": filename,
        "size": _format_size(size_bytes),
        "doc_category": doc_category,
        "upload": "uploaded",
    })


def add_upload_failure(job: dict, filename: str, reason: str):
    """Record an upload failure."""
    job["upload_failures"].append({
        "name": filename,
        "reason": reason,
        "upload": "not uploaded",
    })


def set_analyse_start(job: dict):
    """Mark analysis as started."""
    job["analyse_start_datetime"] = datetime.now(timezone.utc).isoformat()
    job["status"] = "analysing"


def set_analyse_complete(job: dict, error: str = ""):
    """Mark analysis as completed."""
    job["analyse_completed_datetime"] = datetime.now(timezone.utc).isoformat()
    if error:
        job["status"] = "error"
        job["error"] = error
    else:
        job["status"] = "analysed"


def set_create_ehcp_complete(job: dict, doc_name: str, doc_size: int, file_path: str):
    """Mark EHCP creation as completed."""
    job["create_ehcp_completed_datetime"] = datetime.now(
        timezone.utc).isoformat()
    job["create_ehcp_document_name"] = doc_name
    job["create_ehcp_document_size"] = _format_size(doc_size)
    job["create_ehcp_file_path"] = file_path
    job["status"] = "completed"


def add_token_usage(job: dict, prompt_tokens: int = 0, completion_tokens: int = 0, total_tokens: int = 0):
    """Accumulate token usage from an LLM call."""
    job["token_usage"]["prompt_tokens"] += prompt_tokens
    job["token_usage"]["completion_tokens"] += completion_tokens
    job["token_usage"]["total_tokens"] += total_tokens
    # Update formatted values
    job["token_usage"]["prompt_tokens_display"] = _format_tokens(
        job["token_usage"]["prompt_tokens"])
    job["token_usage"]["completion_tokens_display"] = _format_tokens(
        job["token_usage"]["completion_tokens"])
    job["token_usage"]["total_tokens_display"] = _format_tokens(
        job["token_usage"]["total_tokens"])


def add_completeness_entry(job: dict, doc_category: str, total_fields: int,
                           complete_fields: int, missing_fields: int,
                           missing_field_names: list, source_file: str = ""):
    """Add a per-document completeness report entry."""
    pct = round((complete_fields / total_fields)
                * 100) if total_fields > 0 else 0
    job["completeness_report"].append({
        "doc_category": doc_category,
        "source_file": source_file,
        "total_fields": total_fields,
        "filled_fields": complete_fields,
        "missing_fields": missing_fields,
        "completed_pct": f"{pct}%",
        "missing_field_names": missing_field_names,
    })


def set_error(job: dict, error: str):
    """Set job status to error with reason."""
    job["status"] = "error"
    job["error"] = error


async def save_job_record(job: dict):
    """Save the job record to Cosmos DB. Never raises."""
    if not AUDIT_LOG_ENABLED:
        return
    try:
        container = _get_container_client()
        if container is None:
            print(f"  [JobLog] No Cosmos container configured, skipping job log")
            return
        # Strip raw token counts before persisting — only keep display values
        save_copy = dict(job)
        if "token_usage" in save_copy:
            save_copy["token_usage"] = {
                k: v for k, v in save_copy["token_usage"].items()
                if k.endswith("_display")
            }
        await container.upsert_item(body=save_copy)
    except Exception as e:
        print(f"  [JobLog] Failed to save job record: {e}")


async def load_job_record(job_id: str, session_id: str = None) -> Optional[dict]:
    """Load an existing job record from Cosmos DB by job_id.

    Uses a strongly-consistent point read (``read_item``) keyed on the document
    id, which now equals the job_id.  This avoids the query indexing lag that
    could momentarily hide a just-created record and cause a duplicate job to
    be created for the same case."""
    if not AUDIT_LOG_ENABLED or not job_id:
        return None
    try:
        container = _get_container_client()
        if container is None:
            return None
        partition_key = session_id or "anonymous"
        try:
            item = await container.read_item(
                item=job_id, partition_key=partition_key)
        except Exception:
            # Not found (or not yet created) — caller will create it.
            return None
        # Restore raw token counters (stripped before save) so accumulation works
        if "token_usage" in item:
            item["token_usage"].setdefault("prompt_tokens", 0)
            item["token_usage"].setdefault("completion_tokens", 0)
            item["token_usage"].setdefault("total_tokens", 0)
        else:
            item["token_usage"] = {
                "prompt_tokens": 0, "completion_tokens": 0, "total_tokens": 0,
                "prompt_tokens_display": "0", "completion_tokens_display": "0", "total_tokens_display": "0",
            }
        return item
    except Exception as e:
        print(f"  [JobLog] Failed to load job record: {e}")
        return None
