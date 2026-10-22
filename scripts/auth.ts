import fs from "node:fs";
import http from "node:http";
import os from "node:os";
import path from "node:path";
import { exec } from "node:child_process";
import { google } from "googleapis";
import "dotenv/config";
import { SCOPES } from "../src/config.js";

const PORT = 53682;
const REDIRECT_URI = `http://localhost:${PORT}/oauth2callback`;

const clientId = process.env.GOOGLE_CLIENT_ID;
const clientSecret = process.env.GOOGLE_CLIENT_SECRET;

if (!clientId || !clientSecret) {
  console.error(
    [
      "Missing OAuth client credentials.",
      "",
      "Create an OAuth client (type: Desktop app) in Google Cloud Console,",
      "then run:",
      "",
      "  GOOGLE_CLIENT_ID=... GOOGLE_CLIENT_SECRET=... npm run auth",
    ].join("\n"),
  );
  process.exit(1);
}

const oauth2 = new google.auth.OAuth2(clientId, clientSecret, REDIRECT_URI);

const authUrl = oauth2.generateAuthUrl({
  access_type: "offline",
  prompt: "consent",
  scope: SCOPES,
});

console.log("Open this URL in your browser and approve access:\n");
console.log(authUrl);
console.log("");

if (process.platform === "darwin") {
  exec(`open "${authUrl}"`);
}

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url ?? "/", REDIRECT_URI);
  if (url.pathname !== "/oauth2callback") {
    res.writeHead(404).end("Not found");
    return;
  }

  const code = url.searchParams.get("code");
  if (!code) {
    res.writeHead(400).end("Missing authorization code.");
    return;
  }

  try {
    const { tokens } = await oauth2.getToken(code);
    const envPath =
      process.env.GOOGLE_SLIDES_MCP_ENV ??
      path.join(os.homedir(), ".config", "google-slides-mcp", ".env");

    fs.mkdirSync(path.dirname(envPath), { recursive: true });
    fs.writeFileSync(
      envPath,
      [
        `GOOGLE_CLIENT_ID=${clientId}`,
        `GOOGLE_CLIENT_SECRET=${clientSecret}`,
        `GOOGLE_REFRESH_TOKEN=${tokens.refresh_token ?? ""}`,
        "",
      ].join("\n"),
      { mode: 0o600 },
    );

    res
      .writeHead(200, { "content-type": "text/html" })
      .end(
        "<h1>Connected.</h1><p>You can close this tab and return to your terminal.</p>",
      );

    console.log(`Saved credentials to ${envPath}`);
    if (!tokens.refresh_token) {
      console.warn(
        "Google did not return a refresh token. Revoke the app at myaccount.google.com/permissions and run `npm run auth` again.",
      );
    }

    setTimeout(() => {
      server.close();
      process.exit(0);
    }, 300);
  } catch (error) {
    console.error("Token exchange failed:", error);
    res.writeHead(500).end("Token exchange failed. Check the terminal.");
    server.close();
    process.exit(1);
  }
});

server.listen(PORT, () => {
  console.log(`Waiting for the OAuth callback on ${REDIRECT_URI} ...`);
});
