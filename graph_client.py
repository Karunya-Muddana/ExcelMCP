# IMMUTABLE RULE: ExcelMCP never serves stale data.
# If live data cannot be fetched, the operation fails.
# No caching. No fallback. No local file reads.
# Stale data is worse than no data.
# All errors must be loud, clear, and actionable.

import asyncio
import email.utils
import urllib.parse
from typing import Any, Optional

import aiohttp

from excelmcp.auth import get_token

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"

_RETRYABLE_STATUSES = {500, 502, 503, 504}
_BACKOFF_SECONDS = [2, 4, 8]
_MAX_SERVER_ERROR_RETRIES = 3
_MAX_RATE_LIMIT_RETRIES = 3

# Graph error bodies are echoed back to the calling agent. Cap them so a large
# or unexpectedly detailed payload cannot flood the model's context.
_MAX_ERROR_BODY_CHARS = 500

# No timeout at all (aiohttp's default is total=None) meant one unresponsive
# connection could hang a tool call forever. sock_connect/sock_read bound the
# handshake and each read; total bounds the whole request including retries
# inside aiohttp.
_TIMEOUT = aiohttp.ClientTimeout(total=120, sock_connect=15, sock_read=60)

# OneDrive trees can be deep and wide. Recursing unbounded and fanning out with
# an unbounded asyncio.gather produced hundreds of simultaneous requests, which
# reliably triggered Graph throttling. Both dimensions are now capped.
_MAX_SCAN_DEPTH = 12
_MAX_CONCURRENT_REQUESTS = 8

_session: Optional[aiohttp.ClientSession] = None
_session_loop: Optional[asyncio.AbstractEventLoop] = None
_scan_semaphore: Optional[asyncio.Semaphore] = None
_scan_semaphore_loop: Optional[asyncio.AbstractEventLoop] = None


class GraphAPIError(Exception):
    """General Graph API error — server error, connection error, or unhandled status."""


class GraphAuthError(Exception):
    """Authentication or authorization failure."""


class GraphNotFoundError(Exception):
    """Resource not found (404)."""


class GraphRateLimitError(Exception):
    """Rate limit exceeded after max retries."""


class GraphSessionExpiredError(Exception):
    """Workbook session expired."""


def _truncate(text: str) -> str:
    text = " ".join(text.split())
    if len(text) <= _MAX_ERROR_BODY_CHARS:
        return text
    return text[:_MAX_ERROR_BODY_CHARS] + f"... [truncated, {len(text)} chars total]"


def quote_odata_literal(value: str) -> str:
    """Escapes a string for use inside an OData single-quoted literal in a URL path.

    Two distinct escapes are needed and both were missing, which let a sheet name
    steer the request URL:

    1. OData escaping — a literal ``'`` terminates the quoted string, so it must
       be doubled. A sheet named ``Bob's Data`` produced ``worksheets('Bob's Data')``,
       which Graph rejects or misparses.
    2. Percent-encoding with ``safe=""`` — ``urllib.parse.quote`` leaves ``/``
       untouched by default, so a sheet named ``Q1/Q2`` injected an extra path
       segment and addressed a different Graph resource entirely.

    Order matters: double the quotes first, then percent-encode everything.
    """
    return urllib.parse.quote(value.replace("'", "''"), safe="")


def quote_drive_path(path: str) -> str:
    """Percent-encodes a OneDrive folder path for the ``/root:<path>:`` addressing form.

    Each segment is encoded with ``safe=""`` and rejoined with literal ``/`` so
    that reserved characters — notably ``:``, which delimits the path in this
    Graph URL form, plus ``?`` and ``#`` — cannot break out of the path.
    """
    segments = [s for s in path.split("/") if s]
    return "/" + "/".join(urllib.parse.quote(s, safe="") for s in segments)


def _parse_retry_after(header_value: Optional[str], attempt: int) -> float:
    """Returns a sleep duration from a Retry-After header.

    The header is legally either delta-seconds or an HTTP-date; the previous
    ``int(...)`` call raised ValueError on the date form and crashed the request.
    The result is clamped so a hostile or buggy value cannot stall the server.
    """
    fallback = float(_BACKOFF_SECONDS[min(attempt, len(_BACKOFF_SECONDS) - 1)])
    if not header_value:
        return fallback

    try:
        return max(1.0, min(60.0, float(int(header_value))))
    except (TypeError, ValueError):
        pass

    try:
        retry_at = email.utils.parsedate_to_datetime(header_value)
    except (TypeError, ValueError):
        return fallback
    if retry_at is None:
        return fallback

    import datetime as _dt

    now = _dt.datetime.now(_dt.timezone.utc)
    if retry_at.tzinfo is None:
        retry_at = retry_at.replace(tzinfo=_dt.timezone.utc)
    return max(1.0, min(60.0, (retry_at - now).total_seconds()))


async def _get_session() -> aiohttp.ClientSession:
    """Returns a shared ClientSession, rebuilding it if closed or bound to a dead loop.

    A session per request meant a fresh TCP + TLS handshake for every Graph call,
    which is the dominant cost when cross_file_aggregate fans out over many files.
    """
    global _session, _session_loop
    loop = asyncio.get_running_loop()
    if _session is None or _session.closed or _session_loop is not loop:
        if _session is not None and not _session.closed and _session_loop is loop:
            await _session.close()
        _session = aiohttp.ClientSession(timeout=_TIMEOUT)
        _session_loop = loop
    return _session


async def close_client() -> None:
    """Closes the shared session. Safe to call when none is open."""
    global _session, _session_loop
    if _session is not None and not _session.closed:
        await _session.close()
    _session = None
    _session_loop = None


def _get_scan_semaphore() -> asyncio.Semaphore:
    global _scan_semaphore, _scan_semaphore_loop
    loop = asyncio.get_running_loop()
    if _scan_semaphore is None or _scan_semaphore_loop is not loop:
        _scan_semaphore = asyncio.Semaphore(_MAX_CONCURRENT_REQUESTS)
        _scan_semaphore_loop = loop
    return _scan_semaphore


async def _request(
    method: str,
    url: str,
    headers: Optional[dict[str, str]] = None,
    json_data: Optional[dict[str, Any]] = None,
    params: Optional[dict[str, str]] = None,
) -> Optional[dict[str, Any]]:
    req_headers: dict[str, str] = {}
    if headers:
        req_headers.update(headers)

    session = await _get_session()
    token = await get_token()

    server_error_retries = 0
    rate_limit_retries = 0
    connection_retries = 0
    did_force_refresh = False

    while True:
        req_headers["Authorization"] = f"Bearer {token}"
        try:
            async with session.request(
                method, url, headers=req_headers, json=json_data, params=params
            ) as resp:

                if resp.status == 429:
                    if rate_limit_retries >= _MAX_RATE_LIMIT_RETRIES:
                        raise GraphRateLimitError(
                            f"Microsoft Graph rate limit hit and still throttled after "
                            f"{_MAX_RATE_LIMIT_RETRIES} retries. Wait a minute and try "
                            f"again, or scan a smaller folder."
                        )
                    delay = _parse_retry_after(
                        resp.headers.get("Retry-After"), rate_limit_retries
                    )
                    rate_limit_retries += 1
                    print(
                        f"[Graph] Rate limited — retrying in {delay:.0f}s "
                        f"({rate_limit_retries}/{_MAX_RATE_LIMIT_RETRIES})..."
                    )
                    await asyncio.sleep(delay)
                    continue

                if resp.status == 401:
                    if did_force_refresh:
                        raise GraphAuthError(
                            "Microsoft auth token invalid or expired. "
                            "Run excelmcp-setup to re-authenticate."
                        )
                    did_force_refresh = True
                    token = await get_token(force_refresh=True)
                    continue

                if resp.status == 403:
                    text = _truncate(await resp.text())
                    raise GraphAuthError(
                        f"Access denied (403). Ensure the app has Files.Read.All "
                        f"permission. Details: {text}"
                    )

                if resp.status == 404:
                    if "workbook-session-id" in req_headers:
                        raise GraphSessionExpiredError(
                            "Workbook session not found (404). Session may have expired."
                        )
                    raise GraphNotFoundError(url)

                if resp.status in _RETRYABLE_STATUSES:
                    if server_error_retries < _MAX_SERVER_ERROR_RETRIES - 1:
                        delay = _BACKOFF_SECONDS[
                            min(server_error_retries, len(_BACKOFF_SECONDS) - 1)
                        ]
                        server_error_retries += 1
                        print(
                            f"[Graph] Server error {resp.status} — retrying in {delay}s "
                            f"(attempt {server_error_retries}/{_MAX_SERVER_ERROR_RETRIES})..."
                        )
                        await asyncio.sleep(delay)
                        continue
                    text = _truncate(await resp.text())
                    raise GraphAPIError(
                        f"Graph API returned {resp.status} after "
                        f"{_MAX_SERVER_ERROR_RETRIES} retries. OneDrive may be "
                        f"temporarily unavailable. Try again in 30 seconds. "
                        f"Details: {text}"
                    )

                if resp.status >= 400:
                    text = _truncate(await resp.text())
                    raise GraphAPIError(f"Graph API Error {resp.status}: {text}")

                if resp.status == 204 or resp.content_length == 0:
                    return None

                return await resp.json()

        except (aiohttp.ClientConnectionError, asyncio.TimeoutError) as exc:
            if connection_retries < _MAX_SERVER_ERROR_RETRIES - 1:
                delay = _BACKOFF_SECONDS[
                    min(connection_retries, len(_BACKOFF_SECONDS) - 1)
                ]
                connection_retries += 1
                print(
                    f"[Graph] Connection error — retrying in {delay}s "
                    f"(attempt {connection_retries}/{_MAX_SERVER_ERROR_RETRIES}): {exc}"
                )
                await asyncio.sleep(delay)
                continue
            raise GraphAPIError(
                f"Failed to connect to Microsoft Graph API after "
                f"{_MAX_SERVER_ERROR_RETRIES} retries: {exc}. "
                f"Check your internet connection and try again."
            ) from exc


async def _get_children(url: str) -> list[dict[str, Any]]:
    """Fetches all children at a URL, following pagination."""
    items: list[dict[str, Any]] = []
    seen_pages = 0
    while url:
        data = await _request("GET", url)
        if not data:
            break
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
        seen_pages += 1
        if seen_pages > 1000:  # guard against a pathological pagination loop
            print("[Graph] Stopping pagination after 1000 pages.")
            break
    return items


async def _scan_recursive(children_url: str, depth: int = 0) -> list[dict[str, Any]]:
    """Recursively collects .xlsx files from a folder and all its subfolders.

    Depth-limited and throttled by a shared semaphore so a deep tree cannot
    recurse without bound or storm Graph with concurrent requests.
    """
    if depth >= _MAX_SCAN_DEPTH:
        print(f"[Graph] Max scan depth {_MAX_SCAN_DEPTH} reached — not descending further.")
        return []

    semaphore = _get_scan_semaphore()
    async with semaphore:
        items = await _get_children(children_url)

    excel_files = [
        i
        for i in items
        if i.get("file") and i.get("name", "").lower().endswith(".xlsx")
    ]
    subfolders = [i for i in items if i.get("folder")]

    if subfolders:
        subfolder_results = await asyncio.gather(
            *[
                _scan_recursive(
                    f"{GRAPH_BASE_URL}/me/drive/items/{urllib.parse.quote(f['id'], safe='')}/children",
                    depth + 1,
                )
                for f in subfolders
            ]
        )
        for result in subfolder_results:
            excel_files.extend(result)

    return excel_files


async def suggest_paths(attempted_path: str) -> list[str]:
    try:
        parts = attempted_path.strip("/").split("/")
        parent = "/" + "/".join(parts[:-1])
        if parent == "/":
            url = f"{GRAPH_BASE_URL}/me/drive/root/children"
        else:
            url = f"{GRAPH_BASE_URL}/me/drive/root:{quote_drive_path(parent)}:/children"
        items = await _get_children(url)
        return [i["name"] for i in items if i.get("folder")]
    except Exception:
        return []


async def list_excel_files(folder_path: str) -> list[dict[str, Any]]:
    if not folder_path.startswith("/"):
        folder_path = "/" + folder_path

    if folder_path.strip("/") == "":
        root_url = f"{GRAPH_BASE_URL}/me/drive/root/children"
    else:
        root_url = (
            f"{GRAPH_BASE_URL}/me/drive/root:{quote_drive_path(folder_path)}:/children"
        )

    try:
        return await _scan_recursive(root_url)
    except GraphNotFoundError:
        suggestions = await suggest_paths(folder_path)
        msg = f"Folder '{folder_path}' not found on OneDrive."
        if suggestions:
            msg += (
                f" Available folders at this level: "
                f"{', '.join(suggestions)}. "
                f"Try one of these as your folder_path."
            )
        raise GraphAPIError(msg)


def _item_url(item_id: str) -> str:
    return f"{GRAPH_BASE_URL}/me/drive/items/{urllib.parse.quote(item_id, safe='')}"


async def create_session(item_id: str) -> str:
    url = f"{_item_url(item_id)}/workbook/createSession"
    data = await _request("POST", url, json_data={"persistChanges": False})
    return (data or {}).get("id", "")


async def refresh_session(item_id: str, session_id: str) -> None:
    url = f"{_item_url(item_id)}/workbook/refreshSession"
    await _request("POST", url, headers={"workbook-session-id": session_id})


async def close_session(item_id: str, session_id: str) -> None:
    url = f"{_item_url(item_id)}/workbook/closeSession"
    await _request("POST", url, headers={"workbook-session-id": session_id})


async def get_sheet_names(item_id: str, session_id: Optional[str] = None) -> list[str]:
    url = f"{_item_url(item_id)}/workbook/worksheets"
    headers: dict[str, str] = {}
    if session_id:
        headers["workbook-session-id"] = session_id
    data = await _request("GET", url, headers=headers)
    return [sheet["name"] for sheet in (data or {}).get("value", [])]


async def get_used_range(
    item_id: str, sheet_name: str, session_id: Optional[str] = None
) -> dict[str, Any]:
    url = (
        f"{_item_url(item_id)}/workbook/worksheets"
        f"('{quote_odata_literal(sheet_name)}')/usedRange"
    )
    headers: dict[str, str] = {}
    if session_id:
        headers["workbook-session-id"] = session_id
    data = await _request("GET", url, headers=headers)
    return data or {}
