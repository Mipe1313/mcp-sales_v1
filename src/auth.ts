import axios from "axios";
import type { DynamicsConfig, TokenInfo } from "./types.js";

// In-memory token cache keyed by "tenantId:clientId:resourceOrigin"
const tokenCache = new Map<string, TokenInfo>();

function cacheKey(tenantId: string, clientId: string, resourceOrigin: string): string {
  return `${tenantId}:${clientId}:${resourceOrigin}`;
}

/**
 * Returns a valid Bearer token for the given Dynamics 365 config.
 * Uses OAuth2 client_credentials against the Microsoft identity platform.
 * Tokens are cached and refreshed 5 minutes before expiry.
 */
export async function getAccessToken(config: DynamicsConfig): Promise<string> {
  const resourceOrigin = new URL(config.environmentUrl).origin;
  const key = cacheKey(config.tenantId, config.clientId, resourceOrigin);

  const cached = tokenCache.get(key);
  // Keep 5-minute safety buffer before expiry
  if (cached && cached.expiresAt > Date.now() + 5 * 60 * 1000) {
    return cached.accessToken;
  }

  const tokenUrl =
    `https://login.microsoftonline.com/${encodeURIComponent(config.tenantId)}/oauth2/v2.0/token`;

  const body = new URLSearchParams({
    grant_type: "client_credentials",
    client_id: config.clientId,
    client_secret: config.clientSecret,
    scope: `${resourceOrigin}/.default`,
  });

  const response = await axios.post<{ access_token: string; expires_in: number }>(
    tokenUrl,
    body.toString(),
    { headers: { "Content-Type": "application/x-www-form-urlencoded" } },
  );

  const { access_token, expires_in } = response.data;
  tokenCache.set(key, {
    accessToken: access_token,
    expiresAt: Date.now() + expires_in * 1000,
  });

  return access_token;
}
