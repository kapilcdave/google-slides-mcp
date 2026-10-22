import dotenv from "dotenv";
import os from "node:os";
import path from "node:path";

dotenv.config({ quiet: true });

dotenv.config({
  path: path.join(os.homedir(), ".config", "google-slides-mcp", ".env"),
  quiet: true,
});

export const SCOPES = [
  "https://www.googleapis.com/auth/presentations",
  "https://www.googleapis.com/auth/drive.file",
  "https://www.googleapis.com/auth/drive.readonly",
];
