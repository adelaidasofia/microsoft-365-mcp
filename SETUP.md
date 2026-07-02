# Setup — Microsoft 365 MCP

Two ways in:

- **Shared app (team / cohort)** — someone already registered the Entra app and sent you its **Application (client) ID**. You skip the Microsoft Entra portal entirely: jump to [Shared-app mode](#shared-app-mode-the-2-minute-path). ~2 minutes.
- **Your own app (solo)** — you register your own (free) Entra app. Follow §1–§6 below. ~15 minutes once; every account you add after takes ~90 seconds.

Works with **work/school accounts** (Microsoft 365 / Entra ID) and **personal accounts** (outlook.com, hotmail.com, live.com).

## Shared-app mode (the 2-minute path)

Your team or program already registered one Entra app and will send you its **Application (client) ID** (a UUID like `1a2b3c4d-...`). You do **not** touch the Entra portal. Three steps:

**What you need**

- The Client ID your program/admin sends you. It is an app *identifier*, not a secret — but keep it to your team's private channel anyway.
- macOS, Windows, or Linux. Tokens live in the OS keychain on macOS; on Windows/Linux the MCP falls back to a chmod-600 file automatically (see [Token storage](#token-storage)).

**1. Drop the ID in place**

```bash
mkdir -p ~/.claude/microsoft-365-mcp   # or wherever you cloned this repo
cat > /path/to/microsoft-365-mcp/client_config.json <<'EOF'
{"client_id": "PASTE-CLIENT-ID-HERE", "tenant": "common"}
EOF
```

(Or set `M365_CLIENT_ID` as an environment variable instead — both work.)

**2. Authorize your own account**

In Claude Code, run `/mcp` and confirm `microsoft-365` is listed. Then, in a message:

> Call m365_account_add

A browser opens. Sign in with **your** Microsoft account and accept the permissions.

**3. Corporate account blocked? That's the IT step**

If your work account shows **"Need admin approval"**, your company's tenant requires an admin to consent to the app. This is expected on locked-down corporate tenants and it is a one-time, standard request. Send IT:

> Please grant admin consent (or user-consent allowlisting) for the app with **Application (client) ID `<the ID you received>`** for these **delegated** Microsoft Graph permissions: `Mail.ReadWrite`, `Mail.Send`, `Calendars.ReadWrite`, `Files.ReadWrite`, `User.Read`, `offline_access`. The app is a public desktop client; users each sign in with their own account and tokens stay on their own machine.

Once IT approves (Enterprise applications → the app → Permissions → Grant admin consent), re-run `m365_account_add`.

Done — skip to [§6](#6-set-a-default-account-optional) and [§7](#7-verify). You never needed §1–§4.

**Why sharing one Client ID is safe.** This is a *public client* app: there is no app secret at all. The Client ID only names the app; every member signs in with their own Microsoft account, gets their own refresh token, and that token never leaves their machine. The shared app grants no cross-account access.

### Running the shared app for a team (admin)

One person registers the app once, then everyone uses shared-app mode above:

1. Do §1–§4 below (register app, redirect URI, public-client flag, permissions).
2. In §1 step 3, pick **"Accounts in any organizational directory and personal Microsoft accounts"** so any member tenant + personal accounts can sign in.
3. Distribute the **Application (client) ID** to members over a private channel.
4. Members on locked corporate tenants forward the IT request in shared-app step 3 to their admin. There is nothing to approve on YOUR side — consent happens in *their* tenant.

---

## 1. Register an Entra app (solo path)

1. Open https://entra.microsoft.com (or https://portal.azure.com) and sign in with any Microsoft account. No paid subscription is needed for app registration.
2. Go to **Entra ID → App registrations → New registration** (on some tenants: Identity → Applications → App registrations).
3. Fill in:
   - **Name:** `Microsoft 365 MCP` (anything works).
   - **Supported account types:** **"Accounts in any organizational directory and personal Microsoft accounts"** (the most permissive option). Personal-only? Pick "Personal Microsoft accounts only" and use `"tenant": "consumers"` later.
   - **Redirect URI:** platform **"Public client/native (mobile & desktop)"**, value `http://localhost`.
4. Click **Register**.
5. On the app's **Overview** page, copy the **Application (client) ID**.

## 2. Allow public client flows

Still in the app: **Authentication → Advanced settings → Allow public client flows → Yes → Save.**

(This is what makes the browser + device-code sign-in work for a desktop app. Missing this is the #1 setup failure.)

## 3. Add the Graph permissions

**API permissions → Add a permission → Microsoft Graph → Delegated permissions**, add:

- `Mail.ReadWrite`
- `Mail.Send`
- `Calendars.ReadWrite`
- `Files.ReadWrite`
- `User.Read` (usually present already)
- `offline_access` (under "OpenId permissions" — this is what yields refresh tokens)

No admin consent is needed for your own personal account. For a work account in a locked tenant, see the IT note in shared-app step 3.

## 4. Drop the Client ID in place

```bash
cat > /path/to/microsoft-365-mcp/client_config.json <<'EOF'
{"client_id": "PASTE-CLIENT-ID-HERE", "tenant": "common"}
EOF
```

`tenant` values: `common` (work + personal, default) · `consumers` (personal only) · `organizations` (any work account) · your tenant GUID (single locked tenant). Environment variables `M365_CLIENT_ID` / `M365_TENANT_ID` override the file.

`client_config.json` is gitignored.

## 5. Authorize each account

In Claude Code, run `/mcp` and confirm `microsoft-365` appears. Then:

> Call m365_account_add

A browser window opens. Sign in, accept the consent screen, wait for the "authentication complete" page, come back to Claude. Repeat for each account (work + personal) you want connected.

**Headless / remote machine?** Use the device-code flow instead:

> Call m365_account_add with method "device_code"

It returns a URL + code; open the URL on any device, enter the code, sign in, then:

> Call m365_account_complete

## 6. Set a default account (optional)

The first authorized account becomes default automatically. To override, add to `~/.zshrc` (or your shell profile):

```bash
export M365_DEFAULT_ACCOUNT="you@yourcompany.com"
```

Every tool takes an optional `account` param that overrides this per call.

## 7. Verify

```
Call m365_account_list
```

Should return `{"accounts": ["you@yourcompany.com"], "default": "you@yourcompany.com", "count": 1}`.

Then:

```
Call outlook_search with query "" and limit 5
```

Should return up to 5 compact message summaries from your inbox.

---

## Token storage

| Platform | Default | Notes |
|---|---|---|
| macOS | Apple Keychain (service `microsoft-365-mcp`) | Encrypted at rest. |
| Windows | File fallback (`token_cache.json`, chmod-600 equivalent) | Windows Credential Manager caps blobs at ~2.5 KB — smaller than one account's tokens — so the MCP falls back to the file automatically on the first write failure. |
| Linux | Keyring if available, else file fallback | Headless boxes usually land on the file. |

Force the file backend anywhere by setting `M365_TOKEN_FILE=/path/to/token_cache.json` (or just creating `token_cache.json` next to `accounts.py`). The file is gitignored; keep it chmod 600.

**macOS keychain prompts on every call?** That happens on ad-hoc-signed Python interpreters (e.g. uv-managed) whose signature can't anchor a persistent "Always Allow". The file backend avoids the Keychain entirely — set `M365_TOKEN_FILE` and re-add accounts.

## Troubleshooting

**"Need admin approval"** (work account) — your tenant requires admin consent for new apps. Send IT the request in shared-app step 3. This is tenant policy, not an app defect.

**"AADSTS7000218 / public client" errors** — §2 was skipped: set **Allow public client flows = Yes** on the app.

**"AADSTS50011: redirect URI mismatch"** — the app registration is missing the **Public client/native** platform with `http://localhost` (§1 step 3). Add it and retry.

**"AADSTS65001: user or administrator has not consented"** — the permissions in §3 weren't added, or (work account) admin consent is pending.

**Personal account fails on tenant `common`** — some older personal accounts misroute; set `"tenant": "consumers"` in `client_config.json` and re-add.

**"No cached credentials for X"** — that account's tokens were removed or never added on this machine. Run `m365_account_add` and sign in as that account.

**Token refresh stopped working after ~90 days of no use** — Microsoft expires inactive refresh tokens. Re-run `m365_account_add` for that account; active use auto-renews.

**Corporate tenant blocks `anonymous` sharing links** — expected; use `scope="organization"` in `onedrive_share_link` (the default) or share to a specific person with `onedrive_share`.
