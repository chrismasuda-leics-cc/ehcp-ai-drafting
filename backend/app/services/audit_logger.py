"""
Cosmos DB Audit Logger for EHCP Document Processor.

Logs user actions (uploads, analysis, writes, downloads, deletes) to
an Azure Cosmos DB container for compliance and activity tracking.
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
    COSMOS_DB_CONTAINER,
    AUDIT_LOG_ENABLED,
)

_cosmos_client = None
_container_client = None


def _get_container_client():
    """Lazily initialize and return the Cosmos DB container client."""
    global _cosmos_client, _container_client

    if _container_client is not None:
        return _container_client

    if not COSMOS_DB_ENDPOINT:
        return None

    from azure.cosmos.aio import CosmosClient

    if USE_MANAGED_IDENTITY:
        from azure.identity.aio import DefaultAzureCredential
        credential = DefaultAzureCredential()
        _cosmos_client = CosmosClient(COSMOS_DB_ENDPOINT, credential=credential)
    else:
        _cosmos_client = CosmosClient(COSMOS_DB_ENDPOINT, credential=COSMOS_DB_KEY)

    database = _cosmos_client.get_database_client(COSMOS_DB_DATABASE)
    _container_client = database.get_container_client(COSMOS_DB_CONTAINER)
    return _container_client


async def log_action(
    action: str,
    session_id: Optional[str] = None,
    user: Optional[dict] = None,
    details: Optional[dict] = None,
):
    """
    Log an action to Cosmos DB.

    Args:
        action: The action type (e.g., "upload", "analyze", "write_ehcp", "download", "delete")
        session_id: The user's session ID
        user: User claims from Entra ID token (if auth is enabled)
        details: Additional details about the action (filenames, doc types, results, etc.)
    """
    if not AUDIT_LOG_ENABLED:
        return

    try:
        container = _get_container_client()
        if container is None:
            return

        record = {
            "id": str(uuid.uuid4()),
            "partitionKey": session_id or "anonymous",
            "action": action,
            "timestamp": datetime.now(timezone.utc).isoformat(),
            "sessionId": session_id,
            "user": {
                "name": user.get("name", "unknown") if user else "anonymous",
                "email": (user.get("preferred_username") or user.get("upn") or user.get("unique_name") or user.get("email") or "") if user else "",
                "oid": user.get("oid", "") if user else "",
            },
            "details": details or {},
        }

        await container.create_item(body=record)

    except Exception as e:
        # Audit logging should never break the main application
        print(f"  [AuditLog] Failed to log action '{action}': {e}")


