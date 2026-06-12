# Multi-stage build: Vite SPA -> static nginx image for Azure Container Apps.
#
# The web service is a PURE STATIC SPA. MSAL.js (in the bundle) talks to
# Microsoft Graph directly, so there is no backend in this image.

# ---- Stage 1: build the Vite bundle ----
FROM node:22-alpine AS build

WORKDIR /app

# Public MSAL identifiers (client ID + tenant ID) — NOT secrets. They ship in
# every JWT and are baked into the bundle at build time via import.meta.env.
# Existing "Team Pulse" registration is the Graph identity (NOT the EasyAuth
# gate identity, which is the AO-created ao-m365-pull-app registration).
ARG VITE_MSAL_CLIENT_ID="ccef276a-d864-4017-acee-9c7294b401e9"
ARG VITE_MSAL_TENANT_ID="72f988bf-86f1-41af-91ab-2d7cd011db47"
ENV VITE_MSAL_CLIENT_ID=${VITE_MSAL_CLIENT_ID}
ENV VITE_MSAL_TENANT_ID=${VITE_MSAL_TENANT_ID}

# Install deps against the lockfile first (better layer caching).
COPY package.json package-lock.json ./
RUN npm ci

# Build the app (tsc && vite build -> dist/).
COPY . .
RUN npm run build

# ---- Stage 2: serve dist/ with nginx ----
FROM nginx:1.27-alpine AS runtime

# Our nginx config: listen :80, /health, SPA fallback, CSP + security headers.
COPY nginx.conf /etc/nginx/conf.d/default.conf

# Static assets from the build stage.
COPY --from=build /app/dist /usr/share/nginx/html

EXPOSE 80

# nginx:alpine's default CMD already runs nginx in the foreground.
