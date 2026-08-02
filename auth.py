import asyncio
import base64
import json
import os
import time

import msal

from excelmcp.storage import atomic_write_text, get_config_dir

SCOPES = ["Files.Read.All"]
_PROACTIVE_REFRESH_SECONDS = 600  # refresh if token expires within 10 minutes

# Public client ID for the ExcelMCP app registration. This is not a secret —
# device-code flow uses a public client with no client secret, and the value is
# visible in every auth request. Override to point at your own tenant's app.
_DEFAULT_CLIENT_ID = "dd6408d2-c7cf-42fb-926e-7438d6c2aad8"


def get_auth_config() -> tuple[str, str]:
    client_id = os.environ.get("EXCELMCP_CLIENT_ID", _DEFAULT_CLIENT_ID)
    tenant_id = os.environ.get("EXCELMCP_TENANT_ID", "common")
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    return client_id, authority


def get_token_cache_path():
    return get_config_dir() / "token.json"


def _seconds_until_expiry(token: str) -> int:
    """Decode the JWT exp claim (no signature verification) and return seconds remaining.

    Only used to decide whether to proactively refresh; the token itself is
    validated by Microsoft Graph on every call, so no verification is needed here.
    """
    try:
        payload_b64 = token.split(".")[1]
        payload_b64 += "=" * (-len(payload_b64) % 4)
        claims = json.loads(base64.urlsafe_b64decode(payload_b64))
        return max(0, int(claims["exp"]) - int(time.time()))
    except Exception:
        return 0


def _persist(cache: msal.SerializableTokenCache) -> None:
    """Writes the MSAL cache atomically with 0600 permissions."""
    if not cache.has_state_changed:
        return
    try:
        atomic_write_text(get_token_cache_path(), cache.serialize(), sensitive=True)
    except OSError as exc:
        # A cache we cannot persist costs us a re-auth next run — not fatal now.
        print(f"[Auth] Warning: could not save token cache: {exc}")


async def get_token(force_refresh: bool = False) -> str:
    client_id, authority = get_auth_config()

    cache = msal.SerializableTokenCache()
    cache_path = get_token_cache_path()

    if cache_path.exists():
        try:
            cache.deserialize(cache_path.read_text(encoding="utf-8"))
        except Exception:
            print("[Auth] Existing token cache is unreadable — re-authenticating.")

    app = msal.PublicClientApplication(client_id, authority=authority, token_cache=cache)

    accounts = app.get_accounts()
    if accounts:
        result = await asyncio.to_thread(
            app.acquire_token_silent, SCOPES, accounts[0], None, force_refresh
        )
        if result and "access_token" in result:
            if (
                not force_refresh
                and _seconds_until_expiry(result["access_token"])
                < _PROACTIVE_REFRESH_SECONDS
            ):
                refreshed = await asyncio.to_thread(
                    app.acquire_token_silent, SCOPES, accounts[0], None, True
                )
                if refreshed and "access_token" in refreshed:
                    result = refreshed
            _persist(cache)
            return result["access_token"]

    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise ValueError(f"Failed to create device flow: {flow.get('error')}")

    print(
        f"\nGo to https://microsoft.com/devicelogin and enter code: {flow['user_code']}\n",
        flush=True,
    )

    result = await asyncio.to_thread(app.acquire_token_by_device_flow, flow)

    if "access_token" in result:
        _persist(cache)
        print("[Auth] Authenticated successfully.", flush=True)
        return result["access_token"]

    raise RuntimeError(
        f"Authentication failed: {result.get('error')} - {result.get('error_description')}"
    )
