# OneDrive Re-Authentication MCP Tool

## Context

OneDrive tokens expire and currently re-auth requires manually running `authenticate(force_new=True)` in a terminal. When a 401 happens during normal MCP tool usage, `_api_get()` calls `authenticate(force_new=True)` which blocks the entire MCP server for up to 15 minutes waiting for device flow completion. We need MCP tools to trigger and complete re-auth from within Claude Code.

## Approach: Two-Tool Non-Blocking Device Flow

Use MSAL's `exit_condition` parameter to poll without blocking. Two tools:

1. **`onedrive_start_reauth`** — Initiates device flow, persists flow state to disk, returns URL + code
2. **`onedrive_complete_reauth`** — Polls once (non-blocking), returns success/pending/error

Claude calls `start_reauth`, shows the user the URL + code, then calls `complete_reauth` in a loop until done.

## Codex Review Findings (R1)

1. **`exit_condition` signature**: Must be `lambda flow: True` (not `lambda: True`) — MSAL passes the flow dict as an argument
2. **Flow dict serialization**: Safe — `initiate_device_flow()` returns only ints/floats/strings
3. **`_get_headers()` also blocks**: When `_access_token` is None (e.g., after MCP server restart), `_get_headers()` calls `authenticate()` which triggers device flow. Must fix `_get_headers()` too — try silent acquisition only, raise if that fails
4. **Re-persist flow on each poll**: MSAL mutates the flow dict during polling (`latest_attempt_at`, `interval` slow-down). Must write updated flow back to disk after each `poll_reauth()` call
5. **Thread safety**: Low risk today (stdio MCP is single-threaded sync), but note for future
6. **MSAL version**: `pyproject.toml` has `msal>=1.20.0`. `exit_condition` is supported in installed 1.34.0. Consider pinning `>=1.30.0` for this feature

## Files to Modify

### 1. `src/onedrive.py` — Add three helper functions

```python
def check_auth_status() -> dict:
    """Non-destructive check: try silent token acquisition.
    Returns {authenticated: bool, account: str|None}"""

def start_reauth() -> dict:
    """Initiate device flow, save flow state to AUTH_FLOW_FILE.
    Returns {verification_uri, user_code, expires_in}"""

def poll_reauth() -> dict:
    """Load flow state, poll once with exit_condition=lambda flow: True.
    Returns {status: "success"|"pending"|"error", ...}
    On success: save token cache, delete flow file.
    On pending: re-persist mutated flow dict, return as-is.
    On error/expired: delete flow file, return error."""
```

- Flow state persisted to `drive-mcp/onedrive_auth_flow.json` (same dir as token cache)
- `poll_reauth` reconstructs `msal.PublicClientApplication` + cache from scratch each call
- Re-persist flow dict after each poll (MSAL mutates `latest_attempt_at`, `interval`)

### 2. `src/server.py` — Add two MCP tools

```python
@mcp.tool()
def onedrive_start_reauth() -> str:
    """Start OneDrive re-authentication. Returns URL and code for user."""
    result = onedrive.start_reauth()
    # Format as readable message with URL + code

@mcp.tool()
def onedrive_complete_reauth() -> str:
    """Check if OneDrive re-authentication completed. Call after user visits URL."""
    result = onedrive.poll_reauth()
    # Return status message
```

### 3. `src/onedrive.py` — Fix `_get_headers()` and `_api_get()` to never block

**`_get_headers()`** — currently calls `authenticate()` when `_access_token` is None. Change to try silent acquisition only:

```python
def _get_headers():
    """Get authorization headers. Never triggers interactive auth."""
    global _access_token
    if not _access_token:
        # Try silent refresh from cache
        cache = _get_token_cache()
        app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)
        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(SCOPES, account=accounts[0])
            if result and "access_token" in result:
                _access_token = result["access_token"]
                _save_token_cache(cache)
        # Covers: cache exists but no accounts, or silent refresh fails
        if not _access_token:
            raise Exception("OneDrive not authenticated. Call onedrive_start_reauth() to connect.")
    return {"Authorization": f"Bearer {_access_token}"}
```

**`_api_get()`** — on 401, try silent refresh, raise if fails:

```python
def _api_get(url: str) -> dict:
    response = requests.get(url, headers=_get_headers())
    if response.status_code == 401:
        # Try silent refresh (no user interaction)
        cache = _get_token_cache()
        app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)
        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(SCOPES, account=accounts[0])
            if result and "access_token" in result:
                global _access_token
                _access_token = result["access_token"]
                _save_token_cache(cache)
                response = requests.get(url, headers=_get_headers())
                response.raise_for_status()
                return response.json()
        raise Exception("OneDrive token expired. Call onedrive_start_reauth() to renew.")
    response.raise_for_status()
    return response.json()
```

**`_download_file_content()`** — same pattern, silent refresh on 401:

```python
def _download_file_content(download_url: str) -> bytes:
    response = requests.get(download_url, headers=_get_headers())
    if response.status_code == 401:
        cache = _get_token_cache()
        app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)
        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(SCOPES, account=accounts[0])
            if result and "access_token" in result:
                global _access_token
                _access_token = result["access_token"]
                _save_token_cache(cache)
                response = requests.get(download_url, headers=_get_headers())
                response.raise_for_status()
                return response.content
        raise Exception("OneDrive token expired. Call onedrive_start_reauth() to renew.")
    response.raise_for_status()
    return response.content
```

**Important**: After these changes, `authenticate()` is only called from `start_reauth()` flow and the existing `__main__` test block. No MCP-reachable code path triggers interactive/blocking auth.

### 4. `pyproject.toml` — Pin MSAL minimum

Change `msal>=1.20.0` to `msal>=1.30.0` (for `exit_condition` support).

## Usage Flow

```
Claude: calls onedrive_start_reauth()
  → "Visit https://microsoft.com/devicelogin and enter code: ABC123"

User: completes auth in browser

Claude: calls onedrive_complete_reauth()
  → "pending" → waits a few seconds → calls again
  → "success" → done, OneDrive tools work again
```

## Verification

1. Delete `onedrive_token_cache.json` to simulate expired auth
2. Call any OneDrive tool → should get clear error message pointing to `onedrive_start_reauth`
3. Call `onedrive_start_reauth` → get URL + code
4. Complete auth in browser
5. Call `onedrive_complete_reauth` → should return success
6. Call `onedrive_list_root` → should work
