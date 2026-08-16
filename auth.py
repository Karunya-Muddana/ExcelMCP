import asyncio
import base64
import json
import os
import time
from typing import Any, Callable, Optional

import msal
import requests

from excelmcp.storage import atomic_write_text, get_config_dir, log

SCOPES = ["Files.Read.All"]
_PROACTIVE_REFRESH_SECONDS = 600  # refresh if token expires within 10 minutes

# A device code is single-use at the token endpoint, and only this client ever
# redeems it. So "already been used" (AADSTS70000) means our own poll redeemed it
# twice: the first response carried the tokens but was lost in transit, MSAL
# treated the dropped response as retryable, and the retry hit a spent code. The
# user sees "All done!" in the browser while the CLI reports a hard failure.
# The code is gone either way, so the only recovery is a fresh flow.
_DEVICE_FLOW_ATTEMPTS = 3

# Public client ID for the ExcelMCP app registration. This is not a secret —
# device-code flow uses a public client with no client secret, and the value is
# visible in every auth request. Override to point at your own tenant's app.
_DEFAULT_CLIENT_ID = "dd6408d2-c7cf-42fb-926e-7438d6c2aad8"

# MSAL's default HTTP client (requests.Session, no timeout) means a stalled
# TCP handshake to login.microsoftonline.com — flaky wifi, a filtering
# campus/corporate proxy, a dropped SSL packet — blocks forever with no way
# to recover short of killing the process. Every MSAL HTTP call is routed
# through this session so a bad network fails loud after N seconds instead
# of hanging the whole scan indefinitely.
_MSAL_HTTP_TIMEOUT_SECONDS = 20


class _TimeoutSession(requests.Session):
    def request(self, *args, **kwargs):
        kwargs.setdefault("timeout", _MSAL_HTTP_TIMEOUT_SECONDS)
        return super().request(*args, **kwargs)


def get_auth_config() -> tuple[str, str]:
    client_id = os.environ.get("EXCELMCP_CLIENT_ID", _DEFAULT_CLIENT_ID)
    tenant_id = os.environ.get("EXCELMCP_TENANT_ID", "common")
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    return client_id, authority


# msal.PublicClientApplication's constructor performs OIDC tenant discovery —
# a synchronous network round trip to login.microsoftonline.com — every time
# it's instantiated. get_token() used to build a brand-new app on every
# single call and that constructor call was never offloaded to a thread, so
# it ran directly on the event loop. Scanning N sheets concurrently
# (structure.py's asyncio.gather over read_sheet) fired N simultaneous
# blocking HTTP requests, none of them time-bounded, each one capable of
# freezing every other in-flight task the moment it blocked on
# socket.connect(). One slow network path froze the entire scan.
#
# The app (and its token cache) is now built once per process, lazily,
# behind a lock so concurrent callers await the same build instead of racing
# to construct their own — and the construction itself runs in a thread so
# even a slow first call can't block the loop.
_app: Optional[msal.PublicClientApplication] = None
_app_cache: Optional[msal.SerializableTokenCache] = None
_app_lock = asyncio.Lock()


def _build_app() -> tuple[msal.PublicClientApplication, msal.SerializableTokenCache]:
    """Synchronous MSAL app construction. Always run this via asyncio.to_thread."""
    client_id, authority = get_auth_config()

    cache = msal.SerializableTokenCache()
    cache_path = get_token_cache_path()
    if cache_path.exists():
        try:
            cache.deserialize(cache_path.read_text(encoding="utf-8"))
        except Exception:
            log("[Auth] Existing token cache is unreadable — re-authenticating.")

    app = msal.PublicClientApplication(
        client_id,
        authority=authority,
        token_cache=cache,
        http_client=_TimeoutSession(),
    )
    return app, cache


async def _get_app() -> tuple[msal.PublicClientApplication, msal.SerializableTokenCache]:
    """Returns the process-wide MSAL app, building it at most once."""
    global _app, _app_cache
    if _app is not None:
        return _app, _app_cache
    async with _app_lock:
        if _app is None:
            _app, _app_cache = await asyncio.to_thread(_build_app)
    return _app, _app_cache


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
        # The cache is now a process-wide singleton reused across every
        # get_token() call, not a throwaway per-call object, so this flag
        # must be cleared after a successful write — otherwise every future
        # call re-persists the same state to disk forever.
        cache.has_state_changed = False
    except OSError as exc:
        # A cache we cannot persist costs us a re-auth next run — not fatal now.
        log(f"[Auth] Warning: could not save token cache: {exc}")


def _is_spent_code(result: dict) -> bool:
    """True when the device code was consumed before our poll could read the tokens."""
    description = (result.get("error_description") or "").lower()
    return "aadsts70000" in description or "has already been used" in description


def _is_expired_code(result: dict) -> bool:
    """True when the code timed out before the user finished signing in."""
    description = (result.get("error_description") or "").lower()
    return result.get("error") == "expired_token" or "aadsts70019" in description


def _print_device_code(flow: dict) -> None:
    log(
        f"\nGo to {flow.get('verification_uri', 'https://microsoft.com/devicelogin')} "
        f"and enter code: {flow['user_code']}\n",
    )


async def _run_device_flow(
    app: msal.PublicClientApplication,
    cache: msal.SerializableTokenCache,
    on_device_code: Optional[Callable[[dict], None]],
) -> dict:
    """Runs device-code auth, starting a fresh flow if a code is spent or expires."""
    announce = on_device_code or _print_device_code
    last: dict[str, Any] = {}

    for attempt in range(1, _DEVICE_FLOW_ATTEMPTS + 1):
        flow = await asyncio.to_thread(app.initiate_device_flow, scopes=SCOPES)
        if "user_code" not in flow:
            raise RuntimeError(
                f"Could not start device sign-in: {flow.get('error')} "
                f"- {flow.get('error_description')}"
            )

        announce(flow)
        last = await asyncio.to_thread(app.acquire_token_by_device_flow, flow)

        if "access_token" in last:
            return last

        # MSAL may have cached tokens even on a path that reported failure.
        _persist(cache)

        if attempt == _DEVICE_FLOW_ATTEMPTS:
            break

        if _is_spent_code(last):
            log(
                "[Auth] The sign-in completed but the confirmation did not reach us, "
                "so that code is now spent. Starting a new one.",
            )
            continue

        if _is_expired_code(last):
            log("[Auth] That code expired before sign-in finished. Here is a new one.")
            continue

        break

    raise RuntimeError(_device_flow_message(last))


def _device_flow_message(result: dict) -> str:
    error = result.get("error")
    description = result.get("error_description") or ""

    if _is_spent_code(result):
        return (
            "Sign-in kept failing at the last step. Microsoft accepted your login "
            "but the token never reached this machine, which usually means a flaky "
            "or filtered connection to login.microsoftonline.com. This is common "
            "under WSL and behind corporate proxies. Try again on a stable "
            "connection, or run the wizard from the Windows host rather than WSL."
        )
    if _is_expired_code(result):
        return "The device code expired before sign-in finished. Run the command again."
    if error == "authorization_declined":
        return "Sign-in was declined in the browser. Run the command again to retry."
    if "AADSTS65004" in description:
        return (
            "Consent was not granted. ExcelMCP needs the Files.Read.All permission "
            "to read your workbooks. In a managed tenant an administrator may need "
            "to approve it first."
        )
    return f"Authentication failed: {error} - {description}"


async def get_token(
    force_refresh: bool = False,
    on_device_code: Optional[Callable[[dict], None]] = None,
) -> str:
    app, cache = await _get_app()

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

    result = await _run_device_flow(app, cache, on_device_code)
    _persist(cache)
    log("[Auth] Authenticated successfully.")
    return result["access_token"]
