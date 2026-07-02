"""OAuth + multi-account token management for microsoft-365-mcp.

Tokens live in the MSAL token cache, serialized to one of two backends (see
_token_backend): the OS keyring (default on macOS — service name
"microsoft-365-mcp" / KEYRING_SERVICE, encrypted at rest), or a chmod-600 JSON
file when M365_TOKEN_FILE is set or token_cache.json exists next to this
module. The file backend is REQUIRED on Windows (Credential Manager caps a
credential blob at ~2.5 KB; the MSAL cache exceeds that with one account) and
avoids the macOS Keychain's per-app authorization prompts on ad-hoc-signed
interpreters. Accounts are keyed by email (Entra "preferred_username").

The app identity is a PUBLIC client: an Entra Application (client) ID with no
secret. It comes from M365_CLIENT_ID (env) or client_config.json next to this
module ({"client_id": "...", "tenant": "common"}). Tenant: "common" works for
both work and personal accounts; "consumers" for personal-only;
"organizations" or a tenant GUID for a locked corporate tenant. See SETUP.md.

Design: every Outlook/Calendar/OneDrive call takes an `account` email. We
resolve that account in the MSAL cache, acquire_token_silent (auto-refresh),
and hand back a bearer access token for Microsoft Graph.
"""

from __future__ import annotations

import json
import os
import threading
from pathlib import Path
from typing import Iterable

import keyring
import msal

BASE_DIR = Path(__file__).parent
CLIENT_CONFIG_PATH = BASE_DIR / "client_config.json"
ACCOUNTS_INDEX_PATH = BASE_DIR / "accounts_index.json"

KEYRING_SERVICE = "microsoft-365-mcp"
KEYRING_CACHE_KEY = "msal_token_cache"

AUTHORITY_BASE = "https://login.microsoftonline.com"

# Delegated Graph scopes. MSAL adds the reserved scopes (openid, profile,
# offline_access) automatically — passing them explicitly is an MSAL error, so
# they are deliberately absent. offline_access is what yields refresh tokens.
SCOPES = [
    "Mail.ReadWrite",
    "Mail.Send",
    "Calendars.ReadWrite",
    "Files.ReadWrite.All",
    "User.Read",
]

DEFAULT_ACCOUNT_ENV = "M365_DEFAULT_ACCOUNT"
CLIENT_ID_ENV = "M365_CLIENT_ID"
TENANT_ENV = "M365_TENANT_ID"
TOKEN_FILE_ENV = "M365_TOKEN_FILE"
DEFAULT_TOKEN_FILE = BASE_DIR / "token_cache.json"

_lock = threading.Lock()
_app_singleton: "msal.PublicClientApplication | None" = None
_cache_singleton: "msal.SerializableTokenCache | None" = None
_pending_device_flow: dict | None = None


class AccountError(Exception):
    """Raised when an account is missing, unauthorized, or can't be refreshed."""


# --- Client (Entra app) config ---------------------------------------------


def _client_config() -> tuple[str, str]:
    """Return (client_id, tenant). Env wins over client_config.json."""
    client_id = os.environ.get(CLIENT_ID_ENV, "").strip()
    tenant = os.environ.get(TENANT_ENV, "").strip()
    if not client_id and CLIENT_CONFIG_PATH.exists():
        try:
            cfg = json.loads(CLIENT_CONFIG_PATH.read_text())
        except json.JSONDecodeError as e:
            raise AccountError(f"{CLIENT_CONFIG_PATH} is not valid JSON: {e}") from e
        client_id = str(cfg.get("client_id", "")).strip()
        tenant = tenant or str(cfg.get("tenant", "")).strip()
    if not client_id:
        raise AccountError(
            f"No Entra Application (client) ID. Set {CLIENT_ID_ENV} or create "
            f"{CLIENT_CONFIG_PATH} with {{\"client_id\": \"...\"}}. See SETUP.md."
        )
    return client_id, tenant or "common"


# --- Token cache backend ----------------------------------------------------
# Default: OS keyring (encrypted at rest). Opt-in / fallback: a chmod-600 JSON
# file holding the serialized MSAL cache. File mode is active when
# M365_TOKEN_FILE is set OR token_cache.json exists next to this module.
# Keyring write failures (e.g. Windows Credential Manager blob cap) fall back
# to the file automatically. Keep the file gitignored.


def _token_file() -> "Path | None":
    env = os.environ.get(TOKEN_FILE_ENV, "").strip()
    if env:
        return Path(env).expanduser()
    if DEFAULT_TOKEN_FILE.exists():
        return DEFAULT_TOKEN_FILE
    # Windows Credential Manager caps a credential blob at ~2.5 KB — smaller
    # than the MSAL cache once one account is signed in — so default to the
    # file backend there deterministically, rather than relying on a keyring
    # write to fail and fall back mid-session (which would re-prompt login).
    if os.name == "nt":
        return DEFAULT_TOKEN_FILE
    return None


def _cache_read() -> str:
    p = _token_file()
    if p is not None:
        try:
            return p.read_text()
        except OSError:
            return ""
    try:
        return keyring.get_password(KEYRING_SERVICE, KEYRING_CACHE_KEY) or ""
    except Exception:
        return ""


def _cache_write(serialized: str) -> None:
    p = _token_file()
    if p is None:
        try:
            keyring.set_password(KEYRING_SERVICE, KEYRING_CACHE_KEY, serialized)
            return
        except Exception:
            # Keyring unusable (e.g. Windows blob-size cap, headless Linux).
            # Fall back to the file so tokens are never silently dropped.
            p = DEFAULT_TOKEN_FILE
    p.write_text(serialized)
    try:
        p.chmod(0o600)
    except OSError:
        pass


def _cache() -> msal.SerializableTokenCache:
    global _cache_singleton
    if _cache_singleton is None:
        _cache_singleton = msal.SerializableTokenCache()
        existing = _cache_read()
        if existing:
            try:
                _cache_singleton.deserialize(existing)
            except ValueError:
                pass  # corrupt cache -> start fresh; accounts re-add
    return _cache_singleton


def _persist_cache() -> None:
    cache = _cache()
    if cache.has_state_changed:
        _cache_write(cache.serialize())


def _app() -> msal.PublicClientApplication:
    global _app_singleton
    if _app_singleton is None:
        client_id, tenant = _client_config()
        _app_singleton = msal.PublicClientApplication(
            client_id,
            authority=f"{AUTHORITY_BASE}/{tenant}",
            token_cache=_cache(),
        )
    return _app_singleton


# --- Accounts index ---------------------------------------------------------
# Written by add_account / removed by remove_account. Source of truth for
# enumeration: reading the MSAL cache from the keyring just to list accounts
# would decrypt secrets (and on macOS can prompt); the index file avoids that.


def _load_index() -> list[str]:
    if not ACCOUNTS_INDEX_PATH.exists():
        return []
    try:
        return json.loads(ACCOUNTS_INDEX_PATH.read_text())
    except json.JSONDecodeError:
        return []


def _save_index(emails: Iterable[str]) -> None:
    ACCOUNTS_INDEX_PATH.write_text(json.dumps(sorted(set(emails)), indent=2))


def list_accounts() -> list[str]:
    """Return all account emails from the on-disk index."""
    return _load_index()


def default_account() -> str:
    """Return the default account: env override, else first account in index."""
    env_default = os.environ.get(DEFAULT_ACCOUNT_ENV, "").strip()
    if env_default:
        return env_default
    accounts = list_accounts()
    if not accounts:
        raise AccountError(
            "No accounts configured. Run m365_account_add first, or see SETUP.md."
        )
    return accounts[0]


def _resolve(account: str | None) -> str:
    return (account or "").strip() or default_account()


# --- OAuth flows ------------------------------------------------------------


def _register(result: dict) -> dict:
    """Persist a successful token result and index the account."""
    if "error" in result:
        raise AccountError(
            f"OAuth failed: {result.get('error')}: {result.get('error_description', '')}"
        )
    claims = result.get("id_token_claims") or {}
    email = (claims.get("preferred_username") or claims.get("email") or "").lower()
    if not email:
        raise AccountError("OAuth succeeded but the token has no email claim.")
    _persist_cache()
    index = _load_index()
    if email not in index:
        index.append(email)
        _save_index(index)
    return {"email": email, "scopes": result.get("scope", "").split(), "status": "added"}


def add_account(method: str = "interactive") -> dict:
    """Authorize a new account. interactive = browser flow (default).

    method="device_code" starts a device-code flow instead: returns the
    verification URL + code; call complete_device_flow() after entering it.
    """
    if method not in ("interactive", "device_code"):
        raise AccountError("method must be 'interactive' or 'device_code'")
    app = _app()
    if method == "device_code":
        global _pending_device_flow
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise AccountError(f"Device flow failed to start: {flow.get('error_description', flow)}")
        _pending_device_flow = flow
        return {
            "status": "awaiting_user",
            "verification_uri": flow["verification_uri"],
            "user_code": flow["user_code"],
            "expires_in": flow.get("expires_in"),
            "next_step": "Open the URL, enter the code, sign in, then call m365_account_complete.",
        }
    result = app.acquire_token_interactive(SCOPES, prompt="select_account")
    return _register(result)


def complete_device_flow() -> dict:
    """Finish a device-code flow started by add_account(method='device_code').

    Blocks until the user completes sign-in in the browser (or the code expires).
    """
    global _pending_device_flow
    if not _pending_device_flow:
        raise AccountError("No pending device flow. Call m365_account_add with method='device_code' first.")
    result = _app().acquire_token_by_device_flow(_pending_device_flow)
    _pending_device_flow = None
    return _register(result)


def remove_account(email: str) -> dict:
    email = email.lower().strip()
    app = _app()
    for acct in app.get_accounts():
        if (acct.get("username") or "").lower() == email:
            app.remove_account(acct)
    _persist_cache()
    _save_index(e for e in _load_index() if e != email)
    return {"email": email, "status": "removed"}


def get_token(account: str | None = None) -> str:
    """Return a valid Graph access token for the account (silent refresh)."""
    email = _resolve(account).lower()
    app = _app()
    with _lock:
        target = None
        for acct in app.get_accounts():
            if (acct.get("username") or "").lower() == email:
                target = acct
                break
        if target is None:
            raise AccountError(
                f"No cached credentials for {email}. Run m365_account_add to authorize it."
            )
        result = app.acquire_token_silent(SCOPES, account=target)
        _persist_cache()
    if not result or "access_token" not in result:
        detail = (result or {}).get("error_description", "refresh failed")
        raise AccountError(
            f"Could not refresh token for {email}: {detail}. "
            "Re-run m365_account_add for this account."
        )
    return result["access_token"]
