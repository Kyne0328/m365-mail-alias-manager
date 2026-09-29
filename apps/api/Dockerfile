FROM node:22-bookworm-slim AS builder

WORKDIR /app

COPY package.json package-lock.json ./
COPY apps/api/package.json apps/api/package.json
COPY apps/web/package.json apps/web/package.json
RUN npm ci --workspace @alias-manager/api --include-workspace-root=false

COPY apps/api ./apps/api
RUN npm run build -w @alias-manager/api

FROM mcr.microsoft.com/powershell:7.5-ubuntu-24.04

ENV DEBIAN_FRONTEND=noninteractive

RUN apt-get update \
    && apt-get install -y --no-install-recommends ca-certificates curl \
    && curl -fsSL https://deb.nodesource.com/setup_22.x | bash - \
    && apt-get install -y --no-install-recommends nodejs \
    && rm -rf /var/lib/apt/lists/*

RUN pwsh -NoLogo -NoProfile -Command "Set-PSRepository -Name PSGallery -InstallationPolicy Trusted; Install-Module ExchangeOnlineManagement -RequiredVersion 3.10.1 -Scope AllUsers -Force"

WORKDIR /app

COPY package.json package-lock.json ./
COPY apps/api/package.json apps/api/package.json
COPY apps/web/package.json apps/web/package.json
RUN npm ci --omit=dev --workspace @alias-manager/api --include-workspace-root=false

COPY --from=builder /app/apps/api/dist ./apps/api/dist
COPY apps/api/scripts ./apps/api/scripts

WORKDIR /app/apps/api

EXPOSE 8787
CMD ["node", "dist/server.js"]
