# Mail Alias Manager

A multi-tenant PWA for Microsoft 365 users who want a small rotating set of real SMTP aliases on their existing Exchange Online mailbox.

The app is customer-agnostic. It does **not** hard-code a customer tenant, domain, mailbox, or `.onmicrosoft.com` name. New aliases use short human-looking local parts such as `jpeterson@domain`; ownership is tracked separately in Exchange mailbox metadata so normal-looking aliases can still be managed safely.

## What it does

- Signs users in with Microsoft Entra ID.
- Binds the target mailbox to the signed-in user's immutable Microsoft Entra object ID; clients cannot submit an arbitrary mailbox.
- Discovers the signed-in tenant's primary `.onmicrosoft.com` organization automatically; users do not type or store it in the browser.
- Reads accepted Exchange domains and the current mailbox alias set together after tenant admin onboarding.
- Generates short person-style aliases such as `jpeterson@domain`, `mcarter@domain`, and `abrooks@domain`. Each generation advances a durable sequence so a deleted name is not reused.
- Keeps at most 30 app-managed aliases per mailbox across all selected domains.
- On alias 31, removes the oldest managed alias and adds the new one (FIFO).
- Stores app ownership and sequence markers in Exchange `ExtensionCustomAttribute5` using `m365am:v2:*` values; unrelated values in that multivalued attribute are preserved.
- Keeps state in Exchange itself. No customer database is required.
- Never removes the mailbox primary SMTP address or an unmarked normal-looking alias. Legacy `m365am-*` aliases created by older releases remain recognized and removable.

## Architecture

```text
Cloudflare Pages PWA
        |
        | Microsoft access token
        v
Render API (Fastify/Node)
        |
        | verified tenant + signed-in mailbox only
        v
PowerShell 7 + ExchangeOnlineManagement
        |
        v
Customer Exchange Online tenant
```

The browser never receives the Exchange app certificate.

## Requirements

- Node.js 22+
- PowerShell 7 for local Exchange integration
- A Microsoft Entra multi-tenant app registration
- Exchange Online
- An X.509 certificate uploaded to the Entra app registration

## Entra app registration

Create one app registration with **Accounts in any organizational directory**.

### Authentication

Add the deployed PWA URLs as **Single-page application** redirect URIs, for example:

- `http://localhost:5173`
- `https://your-project.pages.dev`

### Expose an API

Expose a delegated scope:

- Scope name: `access_as_user`
- Who can consent: admins and users
- Suggested App ID URI: `api://<client-id>`

The frontend requests:

```text
api://<client-id>/access_as_user
```

Set the app manifest's `api.requestedAccessTokenVersion` to `2`. The backend validates the v2 access token audience (the API client-ID GUID), tenant-specific issuer, tenant ID, immutable user object ID, the `access_as_user` scope, and the `azp` authorized-party claim so tokens acquired by other client applications are rejected.

### Application permissions

Add these application permissions to the app registration:

- **Office 365 Exchange Online → Exchange.ManageAsApp** — used for mailbox alias and accepted-domain operations.
- **Microsoft Graph → Domain.Read.All** — used only to discover a tenant's initial `.onmicrosoft.com` domain when it cannot be derived from the signed-in username.

Every customer tenant must grant admin consent before the app can connect to that tenant. The browser never receives either application permission or the Exchange certificate.

Upload the public half of the server certificate under **Certificates & secrets → Certificates**. Keep the private PFX only in the backend secret store.

## Tenant onboarding

The signed-in tenant is discovered automatically. If the user's Microsoft 365 username already ends in `.onmicrosoft.com`, that domain is used directly. Otherwise the backend uses the tenant ID from the validated access token plus Microsoft Graph `Domain.Read.All` to discover the tenant's initial `.onmicrosoft.com` domain.

The PWA then loads the accepted domains, mailbox, and existing Mail Alias Manager aliases in one bootstrap request. The selected alias domain is a dropdown populated from Exchange; organization identity is never taken from browser local storage.

If Exchange access is not ready, the PWA shows:

1. A tenant-specific **Grant admin consent** link.
2. A generated PowerShell setup script.

The setup script creates custom Exchange roles instead of granting the application the full Exchange Administrator role:

- `Get-Mailbox -Filter -ResultSize`
- `Set-Mailbox -Identity -EmailAddresses -ExtensionCustomAttribute5`
- `Get-AcceptedDomain`

The customer administrator can inspect the script before running it. The setup account needs Exchange Organization Management rights and must be able to consent to the script's Microsoft Graph `Application.Read.All` delegated lookup so it can resolve the tenant-local enterprise application object ID.

> Exchange permission assignments can take time to propagate after initial setup.

## Local development

Install dependencies:

```bash
npm install
```

Copy environment files:

```bash
cp apps/web/.env.example apps/web/.env
cp apps/api/.env.example apps/api/.env
```

The API does not automatically load `.env`. For local development, export the variables in your shell or use your preferred environment runner.

Start the API:

```bash
npm run dev:api
```

Start the PWA in another terminal:

```bash
npm run dev:web
```

The PWA runs on `http://localhost:5173` and the API defaults to `http://localhost:8787`.

## Backend certificate secret

The repository includes a local helper that exports the most recent certificate named `CN=Mail Alias Manager Exchange App`, prompts you for a PFX password, and copies the PFX base64 value directly to your clipboard:

```powershell
pwsh ./scripts/export-render-certificate.ps1
```

Paste the clipboard value directly into Render as `EXCHANGE_CERTIFICATE_BASE64`, and use the password you entered as `EXCHANGE_CERTIFICATE_PASSWORD`.

The generated PFX is stored under `.local-secrets/`, which is git-ignored. Do not paste either secret into chat or commit the private key.

## Render deployment

This repository includes `render.yaml` and `apps/api/Dockerfile`.

The Docker image contains:

- Node.js 22
- PowerShell 7
- `ExchangeOnlineManagement`

Create a Render Blueprint from the repository and populate the secret environment variables.

Required API values:

| Variable | Example |
| --- | --- |
| `ENTRA_CLIENT_ID` | Entra application client ID |
| `API_AUDIENCE` | `<client-id GUID>` |
| `API_SCOPE` | `access_as_user` |
| `ALLOWED_ORIGINS` | `https://your-project.pages.dev` |
| `PUBLIC_APP_URL` | `https://your-project.pages.dev` |
| `EXCHANGE_CERTIFICATE_BASE64` | base64 PFX |
| `EXCHANGE_CERTIFICATE_PASSWORD` | PFX password, if any |

Health check:

```text
GET /health
```

The PWA calls `/health` as soon as it starts so a sleeping Render service can begin waking while the user signs in. That health request also starts a reusable PowerShell worker so `ExchangeOnlineManagement` can load before the first authenticated request. While the service is awake, the worker reuses its Exchange connection instead of reconnecting for each domains/aliases call. The API still does not depend on browser session state; Exchange remains the source of truth.

## Cloudflare Pages deployment

Build command:

```bash
npm run build -w @alias-manager/web
```

Output directory:

```text
apps/web/dist
```

Build variables:

| Variable | Value |
| --- | --- |
| `VITE_ENTRA_CLIENT_ID` | Entra application client ID |
| `VITE_API_SCOPE` | `api://<client-id>/access_as_user` |
| `VITE_API_BASE_URL` | Render API URL |

Add the final Pages URL to both the Entra SPA redirect URI list and the API `ALLOWED_ORIGINS`. The PWA also ships a Cloudflare Pages `_headers` file with clickjacking, MIME-sniffing, referrer, permissions-policy, and safe baseline CSP protections.

## Security model

The API deliberately does not accept a mailbox address from the browser.

For every request it:

1. Validates the Microsoft JWT signature using Microsoft's current signing keys.
2. Validates the configured API audience and delegated scope.
3. Validates the tenant-specific v2 issuer.
4. Uses the immutable `oid` claim to find exactly one Exchange mailbox by `ExternalDirectoryObjectId`.
5. Verifies the supplied primary `.onmicrosoft.com` organization maps to the same tenant ID.
6. Re-checks the requested alias domain against Exchange accepted domains.
7. Only removes aliases recorded in the app's `m365am:v2:*` Exchange metadata, plus legacy `m365am-*` aliases created by older releases.

Mutations are serialized per tenant/user inside each API process so concurrent requests using different prefixes or domains cannot race the mailbox-wide FIFO counter.

## Important production notes

- This app depends on the Exchange Online PowerShell management surface because Microsoft Graph does not expose supported mailbox proxy-address mutation. Microsoft's newer Exchange Admin REST API is currently preview and its Mailbox endpoint does not support changing `EmailAddresses`.
- A Render free instance may sleep. The app now warms the backend and PowerShell worker as early as possible, but a true Render free-tier cold start can still add noticeable latency. The app remains stateless with Exchange as the source of truth, so sleeping does not risk customer data.
- The in-process mutation lock protects one API instance. If the service is later scaled to multiple instances, replace it with a distributed lock before enabling concurrent replicas.
- The 30-alias cap is mailbox-wide for aliases recorded in the app metadata plus legacy `m365am-*` addresses. Unrelated SMTP aliases are never counted or deleted.
- Customer organizations remain responsible for their Microsoft 365 licensing, policies, and acceptable-use requirements.

## Checks

```bash
npm run check
```
