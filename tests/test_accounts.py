"""Account store tests — file token backend + index, no keyring, no network."""

from __future__ import annotations

import json

import pytest

import accounts


@pytest.fixture
def store(tmp_path, monkeypatch):
    """Route the token cache to a temp file and the index to a temp path."""
    token_file = tmp_path / "token_cache.json"
    monkeypatch.setenv(accounts.TOKEN_FILE_ENV, str(token_file))
    monkeypatch.setattr(accounts, "ACCOUNTS_INDEX_PATH", tmp_path / "accounts_index.json")
    monkeypatch.setattr(accounts, "_cache_singleton", None)
    monkeypatch.setattr(accounts, "_app_singleton", None)
    return token_file


def test_index_round_trip(store):
    assert accounts.list_accounts() == []
    accounts._save_index(["b@example.com", "a@example.com", "a@example.com"])
    assert accounts.list_accounts() == ["a@example.com", "b@example.com"]


def test_default_account_env_override(store, monkeypatch):
    accounts._save_index(["a@example.com"])
    assert accounts.default_account() == "a@example.com"
    monkeypatch.setenv(accounts.DEFAULT_ACCOUNT_ENV, "z@example.com")
    assert accounts.default_account() == "z@example.com"


def test_default_account_empty_raises(store):
    with pytest.raises(accounts.AccountError, match="m365_account_add"):
        accounts.default_account()


def test_cache_write_creates_chmod_600_file(store):
    accounts._cache_write('{"AccessToken": {}}')
    assert store.exists()
    assert oct(store.stat().st_mode & 0o777) == "0o600"
    assert accounts._cache_read() == '{"AccessToken": {}}'


def test_client_config_env_wins(store, monkeypatch):
    monkeypatch.setenv(accounts.CLIENT_ID_ENV, "client-from-env")
    monkeypatch.setenv(accounts.TENANT_ENV, "consumers")
    assert accounts._client_config() == ("client-from-env", "consumers")


def test_client_config_missing_raises_with_setup_pointer(store, monkeypatch):
    monkeypatch.delenv(accounts.CLIENT_ID_ENV, raising=False)
    monkeypatch.setattr(accounts, "CLIENT_CONFIG_PATH", store.parent / "client_config.json")
    with pytest.raises(accounts.AccountError, match="SETUP.md"):
        accounts._client_config()


def test_client_config_file_fallback(store, monkeypatch):
    monkeypatch.delenv(accounts.CLIENT_ID_ENV, raising=False)
    monkeypatch.delenv(accounts.TENANT_ENV, raising=False)
    cfg = store.parent / "client_config.json"
    cfg.write_text(json.dumps({"client_id": "abc-123", "tenant": "orgtenant"}))
    monkeypatch.setattr(accounts, "CLIENT_CONFIG_PATH", cfg)
    assert accounts._client_config() == ("abc-123", "orgtenant")


def test_register_rejects_error_result(store):
    with pytest.raises(accounts.AccountError, match="OAuth failed"):
        accounts._register({"error": "access_denied", "error_description": "nope"})


def test_register_requires_email_claim(store):
    with pytest.raises(accounts.AccountError, match="no email claim"):
        accounts._register({"id_token_claims": {}})


def test_register_indexes_account(store):
    out = accounts._register(
        {"id_token_claims": {"preferred_username": "User@Example.COM"}, "scope": "Mail.ReadWrite"}
    )
    assert out["email"] == "user@example.com"
    assert accounts.list_accounts() == ["user@example.com"]


def test_add_account_rejects_unknown_method(store):
    with pytest.raises(accounts.AccountError, match="interactive"):
        accounts.add_account(method="magic")


def test_complete_device_flow_without_pending_raises(store, monkeypatch):
    monkeypatch.setattr(accounts, "_pending_device_flow", None)
    with pytest.raises(accounts.AccountError, match="No pending device flow"):
        accounts.complete_device_flow()


def test_scopes_exclude_msal_reserved():
    # MSAL errors if openid/profile/offline_access are passed explicitly.
    assert not {"openid", "profile", "offline_access"} & set(accounts.SCOPES)


def test_scopes_least_privilege_onedrive():
    # OneDrive tools hit only /me/drive/* → Files.ReadWrite (own drive) is
    # sufficient. Files.ReadWrite.All reaches every file the user can access
    # (org/SharePoint) and is over-broad; it must never creep back in. Keeping
    # the requested scope minimal is a SEPARATE invariant from "scopes match
    # across surfaces" — this test locks it (MYC-2578).
    assert "Files.ReadWrite" in accounts.SCOPES
    assert "Files.ReadWrite.All" not in accounts.SCOPES


def test_windows_defaults_to_file_backend(monkeypatch, tmp_path):
    # Windows Credential Manager can't hold the MSAL blob → must use the file.
    monkeypatch.delenv(accounts.TOKEN_FILE_ENV, raising=False)
    monkeypatch.setattr(accounts, "DEFAULT_TOKEN_FILE", tmp_path / "token_cache.json")
    monkeypatch.setattr(accounts.os, "name", "nt")
    assert accounts._token_file() == tmp_path / "token_cache.json"


def test_posix_defaults_to_keyring(monkeypatch, tmp_path):
    monkeypatch.delenv(accounts.TOKEN_FILE_ENV, raising=False)
    monkeypatch.setattr(accounts, "DEFAULT_TOKEN_FILE", tmp_path / "nope.json")  # absent
    monkeypatch.setattr(accounts.os, "name", "posix")
    assert accounts._token_file() is None  # → keyring backend
