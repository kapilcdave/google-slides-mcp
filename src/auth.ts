import fs from "node:fs";
import { google } from "googleapis";
import { SCOPES } from "./config.js";

export type GoogleAuthClient = ReturnType<typeof createAuth>;

function serviceAccountCredentials(): Record<string, unknown> {
  const raw = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
  if (!raw) throw new Error("GOOGLE_SERVICE_ACCOUNT_JSON is not set");

  const json = raw.trim().startsWith("{")
    ? raw
    : fs.readFileSync(raw, "utf8");

  return JSON.parse(json);
}

/**
 * Builds a Google auth client from whichever credentials are available.
 *
 * Supported, in order of precedence:
 *   1. GOOGLE_ACCESS_TOKEN                                  (short-lived token)
 *   2. GOOGLE_CLIENT_ID + GOOGLE_CLIENT_SECRET + GOOGLE_REFRESH_TOKEN
 *   3. GOOGLE_SERVICE_ACCOUNT_JSON (+ optional GOOGLE_IMPERSONATE_USER)
 *   4. GOOGLE_APPLICATION_CREDENTIALS
 */
export function createAuth() {
  const accessToken = process.env.GOOGLE_ACCESS_TOKEN;
  if (accessToken) {
    const client = new google.auth.OAuth2();
    client.setCredentials({ access_token: accessToken });
    return client;
  }

  const refreshToken = process.env.GOOGLE_REFRESH_TOKEN;
  if (refreshToken) {
    const clientId = process.env.GOOGLE_CLIENT_ID;
    const clientSecret = process.env.GOOGLE_CLIENT_SECRET;
    if (!clientId || !clientSecret) {
      throw new Error(
        "GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET are required when using GOOGLE_REFRESH_TOKEN.",
      );
    }
    const client = new google.auth.OAuth2(clientId, clientSecret);
    client.setCredentials({ refresh_token: refreshToken });
    return client;
  }

  if (process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
    const credentials = serviceAccountCredentials();
    const subject = process.env.GOOGLE_IMPERSONATE_USER;
    return new google.auth.GoogleAuth({
      credentials,
      scopes: SCOPES,
      ...(subject ? { clientOptions: { subject } } : {}),
    });
  }

  if (process.env.GOOGLE_APPLICATION_CREDENTIALS) {
    return new google.auth.GoogleAuth({ scopes: SCOPES });
  }

  throw new Error(
    [
      "No Google credentials found.",
      "Set one of:",
      "  GOOGLE_ACCESS_TOKEN",
      "  GOOGLE_CLIENT_ID + GOOGLE_CLIENT_SECRET + GOOGLE_REFRESH_TOKEN",
      "  GOOGLE_SERVICE_ACCOUNT_JSON (+ optional GOOGLE_IMPERSONATE_USER)",
      "  GOOGLE_APPLICATION_CREDENTIALS",
      "Run `npm run auth` to create a refresh token.",
    ].join("\n"),
  );
}
