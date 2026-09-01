"""
Microsoft Entra ID (Azure AD) authentication for the EHCP backend.

Validates JWT Bearer tokens issued by Entra ID. Tokens are obtained by the
Streamlit frontend via MSAL and sent in the Authorization header.
"""

from typing import Optional

from fastapi import Depends, HTTPException, Request, status
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
import jwt
from jwt import PyJWKClient

from app.settings import (
    ENTRA_TENANT_ID,
    ENTRA_CLIENT_ID,
    ENTRA_AUTHORITY,
    AUTH_ENABLED,
)

# Bearer token extraction
_bearer_scheme = HTTPBearer(auto_error=False)

# JWKS client for fetching Microsoft's signing keys (cached internally)
_jwks_client: Optional[PyJWKClient] = None


def _get_jwks_client() -> PyJWKClient:
    """Lazily initialize the JWKS client."""
    global _jwks_client
    if _jwks_client is None:
        jwks_url = f"{ENTRA_AUTHORITY}/discovery/v2.0/keys"
        _jwks_client = PyJWKClient(jwks_url, cache_keys=True)
    return _jwks_client


def _decode_token(token: str) -> dict:
    """Decode and validate a JWT token from Entra ID."""
    jwks_client = _get_jwks_client()
    signing_key = jwks_client.get_signing_key_from_jwt(token)
    valid_audiences = [ENTRA_CLIENT_ID, f"api://{ENTRA_CLIENT_ID}"]
    valid_issuers = [
        f"https://login.microsoftonline.com/{ENTRA_TENANT_ID}/v2.0",
        f"https://sts.windows.net/{ENTRA_TENANT_ID}/",
    ]

    # Entra ID token validation
    payload = jwt.decode(
        token,
        signing_key.key,
        algorithms=["RS256"],
        audience=valid_audiences,
        issuer=valid_issuers,
        options={
            "verify_exp": True,
            "verify_aud": True,
            "verify_iss": True,
        },
    )
    return payload


async def get_current_user(
    request: Request,
    credentials: Optional[HTTPAuthorizationCredentials] = Depends(_bearer_scheme),
) -> Optional[dict]:
    """
    FastAPI dependency that validates the Bearer token.

    If AUTH_ENABLED is False, returns user info from headers (X-User-Name, X-User-Email).
    If AUTH_ENABLED is True, validates the token and returns user claims.
    """
    if not AUTH_ENABLED:
        # Fall back to user info passed via headers from the frontend
        user_name = request.headers.get("x-user-name")
        user_email = request.headers.get("x-user-email")
        if user_name:
            return {"name": user_name, "preferred_username": user_email or ""}
        return None

    if credentials is None:
        auth_header = request.headers.get("authorization", "NONE")
        print(f"[AUTH DEBUG] No credentials found. Authorization header: {auth_header[:50] if auth_header else 'MISSING'}")
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Missing authorization token",
            headers={"WWW-Authenticate": "Bearer"},
        )

    try:
        payload = _decode_token(credentials.credentials)
        return payload
    except jwt.ExpiredSignatureError:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Token has expired",
            headers={"WWW-Authenticate": "Bearer"},
        )
    except jwt.InvalidAudienceError as e:
        # Debug: decode without verification to see actual audience
        try:
            unverified = jwt.decode(credentials.credentials, options={"verify_signature": False, "verify_aud": False, "verify_exp": False, "verify_iss": False})
            actual_aud = unverified.get("aud", "MISSING")
            actual_iss = unverified.get("iss", "MISSING")
            print(f"[AUTH DEBUG] Token aud={actual_aud}, iss={actual_iss}, expected_aud={[ENTRA_CLIENT_ID, f'api://{ENTRA_CLIENT_ID}']}")
        except Exception:
            pass
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail=f"Invalid token audience: {e}",
            headers={"WWW-Authenticate": "Bearer"},
        )
    except jwt.InvalidIssuerError as e:
        try:
            unverified = jwt.decode(credentials.credentials, options={"verify_signature": False, "verify_aud": False, "verify_exp": False, "verify_iss": False})
            actual_iss = unverified.get("iss", "MISSING")
            print(f"[AUTH DEBUG] Token iss={actual_iss}, expected=https://login.microsoftonline.com/{ENTRA_TENANT_ID}/v2.0")
        except Exception:
            pass
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail=f"Invalid token issuer: {e}",
            headers={"WWW-Authenticate": "Bearer"},
        )
    except Exception as e:
        print(f"[AUTH DEBUG] Generic error: {type(e).__name__}: {e}")
        try:
            unverified = jwt.decode(credentials.credentials, options={"verify_signature": False, "verify_aud": False, "verify_exp": False, "verify_iss": False})
            print(f"[AUTH DEBUG] Token claims: aud={unverified.get('aud')}, iss={unverified.get('iss')}, sub={unverified.get('sub')}")
        except Exception:
            print(f"[AUTH DEBUG] Could not decode token at all. Token starts with: {credentials.credentials[:20]}...")
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail=f"Invalid token: {str(e)}",
            headers={"WWW-Authenticate": "Bearer"},
        )
