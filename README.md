# Mail Alias Manager

A multi-tenant PWA for Microsoft 365 users who want a small rotating set of real SMTP aliases on their existing Exchange Online mailbox.

The app is customer-agnostic. It does **not** hard-code a customer tenant, domain, mailbox, or `.onmicrosoft.com` name. The only reserved naming convention is the `m365am-` prefix used to distinguish aliases this app owns from unrelated mailbox aliases.

## What it does

- Signs users in with Microsoft Entra ID.
- Binds the target mailbox to the signed-in user's immutable Microsoft Entra object ID; clients cannot submit an arbitrary mailbox.
- Verifies the supplied primary `.onmicrosoft.com` organization domain resolves to the same Entra tenant as the token.
- Reads accepted Exchange domains after tenant admin onboarding.
- Generates aliases in the reserved form `m365am-<prefix>-000001-<random>@domain`: the number auto-increments while the random suffix prevents practical address reuse after manual deletion.
- Keeps at most 30 app-managed aliases per mailbox, across all selected prefixes and domains.
- On alias 31, removes the oldest managed alias and adds the new one (FIFO).
- Keeps state in Exchange itself. No customer database is required.
- Never removes the mailbox primary SMTP address or aliases outside the reserved app-managed namespace.

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

### Exchange application permission

Add **Office 365 Exchange Online → Application permissions → Exchange.ManageAsApp**.

Every customer tenant must grant admin consent before the app can connect to that tenant.

Upload the public half of the server certificate under **Certificates & secrets → Certificates**. Keep the private PFX only in the backend secret store.

## Tenant onboarding

A signed-in user enters the tenant's primary domain, such as:

```text
contoso.onmicrosoft.com
```

The backend resolves Microsoft's OpenID configuration for that domain and verifies the resulting tenant GUID matches the user's signed-in tenant.

If Exchange access is not ready, the PWA shows:

1. A tenant-specific **Grant admin consent** link.
2. A generated PowerShell setup script.

The setup script creates custom Exchange roles instead of granting the application the full Exchange Administrator role:

- `Get-Mailbox -Filter -ResultSize`
- `Set-Mailbox -Identity -EmailAddresses`
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

The PWA calls `/health` as soon as it starts so a sleeping Render service can begin waking while the user signs in. The API does not depend on server-side session state.

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
7. Only removes aliases in the reserved `m365am-<prefix>-######-<12-hex>@domain` namespace.

Mutations are serialized per tenant/user inside each API process so concurrent requests using different prefixes or domains cannot race the mailbox-wide FIFO counter.

## Important production notes

- This app depends on the Exchange Online PowerShell management surface because Microsoft Graph does not expose supported mailbox proxy-address mutation. Microsoft's newer Exchange Admin REST API is currently preview and its Mailbox endpoint does not support changing `EmailAddresses`.
- A Render free instance may sleep. The app remains stateless, so sleeping does not risk customer data.
- The in-process mutation lock protects one API instance. If the service is later scaled to multiple instances, replace it with a distributed lock before enabling concurrent replicas.
- The 30-alias cap is mailbox-wide for aliases in the reserved `m365am-` namespace. Unrelated SMTP aliases are never counted or deleted.
- Customer organizations remain responsible for their Microsoft 365 licensing, policies, and acceptable-use requirements.

## Checks

```bash
npm run check
```
