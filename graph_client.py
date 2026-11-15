"""Minimal Microsoft Graph client using delegated tokens from TokenManager."""
from __future__ import annotations

import asyncio
import logging

import httpx

from auth import TokenManager

log = logging.getLogger("msteams-mcp.graph")

GRAPH = "https://graph.microsoft.com/v1.0"


class GraphError(Exception):
    """A Graph call failed in a way worth explaining to the user."""

    def __init__(self, message: str, status: int = 0, code: str = ""):
        super().__init__(message)
        self.status = status
        self.code = code


class GraphClient:
    def __init__(self, tokens: TokenManager, timeout: float = 30,
                 transport: httpx.AsyncBaseTransport | None = None):
        self.tokens = tokens
        self._http = httpx.AsyncClient(
            base_url=GRAPH,
            headers={"Accept": "application/json"},
            timeout=timeout,
            follow_redirects=False,
            transport=transport,
        )

    async def aclose(self) -> None:
        await self._http.aclose()

    async def request(self, method: str, path: str, *, params: dict | None = None,
                      json: dict | None = None, headers: dict | None = None) -> dict:
        # Absolute URLs only come from @odata.nextLink; never send the token elsewhere.
        if path.startswith("http") and not path.startswith(GRAPH + "/"):
            raise GraphError(f"Refusing to follow a link outside Microsoft Graph: {path[:100]}")
        force, retried_401 = False, False
        for attempt in range(4):
            token = await self.tokens.access_token(force=force)
            force = False
            try:
                r = await self._http.request(
                    method, path, params=params, json=json,
                    headers={**(headers or {}), "Authorization": f"Bearer {token}"},
                )
            except httpx.HTTPError as e:
                raise GraphError(f"Could not reach Microsoft Graph: {e}") from e
            if r.status_code == 401 and not retried_401:
                retried_401, force = True, True
                continue
            if r.status_code in (429, 503, 504) and attempt < 3:
                try:
                    delay = float(r.headers.get("Retry-After", "2"))
                except ValueError:
                    delay = 2.0
                log.warning("Graph %s on %s, retrying in %.0fs", r.status_code, path[:80], delay)
                await asyncio.sleep(min(max(delay, 1.0), 15.0))
                continue
            break
        if r.status_code >= 400:
            raise self._error(r, method, path)
        if r.status_code == 204 or not r.content:
            return {}
        return r.json()

    @staticmethod
    def _error(r: httpx.Response, method: str, path: str) -> GraphError:
        try:
            err = r.json().get("error") or {}
        except ValueError:
            err = {}
        code = err.get("code") or ""
        msg = (err.get("message") or r.text[:300] or "").strip()
        where = f"{method} {path.split('?')[0][:120]}"
        if r.status_code == 401:
            text = f"401 Unauthorized on {where}: {msg}"
        elif r.status_code == 403:
            text = (f"403 Forbidden on {where}: {msg} (the signed-in user, or the app's "
                    f"delegated permissions, don't allow this)")
        elif r.status_code == 404:
            text = f"404 Not found on {where}: {msg} (check the id; the user may not be a member)"
        else:
            text = f"Microsoft Graph returned {r.status_code} {code} on {where}: {msg}"
        return GraphError(text, r.status_code, code)

    async def get(self, path: str, params: dict | None = None, headers: dict | None = None) -> dict:
        return await self.request("GET", path, params=params, headers=headers)

    async def post(self, path: str, json: dict) -> dict:
        return await self.request("POST", path, json=json)

    async def get_all(self, path: str, params: dict | None = None, limit: int = 100,
                      headers: dict | None = None) -> tuple[list[dict], bool]:
        """Follow @odata.nextLink until `limit` items. Returns (items, more_available)."""
        items: list[dict] = []
        url: str | None = path
        while url and len(items) < limit:
            page = await self.request("GET", url, params=params, headers=headers)
            items.extend(page.get("value") or [])
            url = page.get("@odata.nextLink")
            params = None  # the next link already carries the query
        return items[:limit], bool(url) or len(items) > limit

    async def scan(self, path: str, params: dict | None = None, limit: int = 100, *,
                   keep=None, stop=None, max_scan: int = 1000,
                   headers: dict | None = None) -> tuple[list[dict], str]:
        """Page through a collection in server order, keeping items that pass
        `keep`, until one of these happens (the second value says which):
          'stop'  - `stop(item)` was true (e.g. an ordered cutoff was passed);
          'limit' - `limit` items were kept and at least one more would be kept;
          'scan'  - `max_scan` items were examined (more may exist);
          'end'   - the collection ran out.
        After reaching `limit` it keeps reading (within `max_scan`) until it can tell
        'limit' (more exist) from 'stop'/'end' (the list is complete)."""
        kept: list[dict] = []
        seen = 0
        url: str | None = path
        while url:
            page = await self.request("GET", url, params=params, headers=headers)
            params = None  # the next link already carries the query
            for item in page.get("value") or []:
                if stop and stop(item):
                    return kept, "stop"
                if keep is None or keep(item):
                    if len(kept) >= limit:
                        return kept, "limit"
                    kept.append(item)
                seen += 1
                if seen >= max_scan:
                    return kept, "scan"
            url = page.get("@odata.nextLink")
        return kept, "end"
