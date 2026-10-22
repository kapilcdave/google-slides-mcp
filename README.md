# Google Slides MCP

An open-source [Model Context Protocol](https://modelcontextprotocol.io) server for **Google Slides**. Connect it to any MCP client — Claude, Cursor, Windsurf, ChatGPT desktop, or a local model — and the AI can create decks, add slides, write text, insert images/tables/shapes, edit speaker notes, restyle content, and export to PDF/PPTX.

- **Any model.** MCP is model-agnostic; the server exposes tools, your client's model decides how to use them.
- **Any Google token.** OAuth refresh token, short-lived access token, service account, or Application Default Credentials.
- **No lock-in.** The AI drives the official Google Slides API directly. Your decks stay in your Google account.

> Not affiliated with Google or Anthropic. Google Slides is a trademark of Google LLC.

## Quick start

```bash
git clone https://github.com/kapilcdave/google-slides-mcp
cd google-slides-mcp
npm install
npm run build
```

### 1. Enable the APIs

In [Google Cloud Console](https://console.cloud.google.com/):

1. Create (or pick) a project.
2. Enable **Google Slides API** and **Google Drive API**.
3. Configure the OAuth consent screen. Add yourself as a test user while the app is in testing.
4. Create credentials → **OAuth client ID** → **Desktop app**.
5. Copy the client ID and client secret.

### 2. Get a refresh token

```bash
GOOGLE_CLIENT_ID=your-id \
GOOGLE_CLIENT_SECRET=your-secret \
npm run auth
```

This opens a browser, asks for the `presentations`, `drive.file`, and `drive.readonly` scopes, and writes:

```
~/.config/google-slides-mcp/.env
```

The server loads that file automatically, so you do not need to repeat the credentials in every client config.

### 3. Connect your MCP client

**Claude Code**

```bash
claude mcp add google-slides -- node /absolute/path/to/google-slides-mcp/dist/index.js
```

**Claude Desktop** (`~/Library/Application Support/Claude/claude_desktop_config.json` on macOS)

```json
{
  "mcpServers": {
    "google-slides": {
      "command": "node",
      "args": ["/absolute/path/to/google-slides-mcp/dist/index.js"]
    }
  }
}
```

**Cursor / Windsurf / generic MCP config**

```json
{
  "mcpServers": {
    "google-slides": {
      "command": "node",
      "args": ["/absolute/path/to/google-slides-mcp/dist/index.js"],
      "env": {
        "GOOGLE_CLIENT_ID": "optional-if-not-in-env-file",
        "GOOGLE_CLIENT_SECRET": "optional-if-not-in-env-file",
        "GOOGLE_REFRESH_TOKEN": "optional-if-not-in-env-file"
      }
    }
  }
}
```

Run without cloning (uses the GitHub repo):

```json
{
  "mcpServers": {
    "google-slides": {
      "command": "npx",
      "args": ["-y", "github:kapilcdave/google-slides-mcp"]
    }
  }
}
```

Restart the client and ask: **"Make me a 10-slide deck about the future of urban farming."**

## Authentication options

The server picks the first available option:

| Credentials | Use case |
| --- | --- |
| `GOOGLE_CLIENT_ID` + `GOOGLE_CLIENT_SECRET` + `GOOGLE_REFRESH_TOKEN` | Recommended. Long-lived, auto-refreshing. |
| `GOOGLE_ACCESS_TOKEN` | Bring any short-lived OAuth access token (e.g. from `gcloud auth print-access-token` with the right scopes). Expires in ~1 hour. |
| `GOOGLE_SERVICE_ACCOUNT_JSON` | Service account key file path or inline JSON. Share the target deck with the service account email. |
| `GOOGLE_SERVICE_ACCOUNT_JSON` + `GOOGLE_IMPERSONATE_USER` | Domain-wide delegation: act as a Workspace user. |
| `GOOGLE_APPLICATION_CREDENTIALS` | Application Default Credentials. |

Required scopes:

```
https://www.googleapis.com/auth/presentations
https://www.googleapis.com/auth/drive.file
https://www.googleapis.com/auth/drive.readonly
```

`drive.readonly` is only needed for exporting decks you did not create with this app. Drop it if you only need `export_deck` on decks created by the server.

## Tools

| Tool | What it does |
| --- | --- |
| `create_presentation` | Create an empty deck |
| `create_deck_from_outline` | Build a full themed deck from structured slides + bullets + notes |
| `get_presentation` | Compact summary of slides, element IDs, text, notes (or raw JSON) |
| `list_slides` | List slides, indexes, IDs, titles |
| `get_slide` | Full element detail for one slide |
| `add_slide` | Add a slide with any predefined layout, title, body, bullets, notes, theme |
| `delete_slide` | Delete a slide |
| `duplicate_slide` | Duplicate a slide |
| `move_slide` | Move one slide to an index |
| `reorder_slides` | Set the complete slide order |
| `add_text_box` | Add styled text (font, size, color, alignment, bullets) |
| `add_bullets` | Add a bullet list |
| `update_text` | Replace text in an element |
| `replace_all_text` | Deck-wide find and replace |
| `format_text` | Style a character range (bold, italic, size, color, highlight, link) |
| `add_image` | Insert a public image URL |
| `add_shape` | Insert a shape with fill, outline and text |
| `add_table` | Insert a table filled from a 2D array |
| `set_slide_background` | Set a solid background color |
| `set_speaker_notes` | Replace speaker notes |
| `delete_element` | Delete a shape/image/table |
| `export_deck` | Export to PDF or PPTX via Drive |
| `batch_update` | Raw Slides API `batchUpdate` escape hatch |

There is also a `build_deck` prompt that gives the model a good outline-first workflow.

### Themes

`midnight`, `light`, `sunset`, `forest`. Themes set the background, title/body colors, accent bar and font. Coordinates in tools are in **inches** on a 13.333 × 7.5 in widescreen slide.

## Example prompts

```
Build a 12-slide investor update for a climate tech startup, midnight theme,
with speaker notes on every slide.

Turn this outline into a deck: ...

Add a two-column comparison slide after slide 4 and use the forest theme.

Replace every "Q3" with "Q4" and export the deck to PDF.
```

## How it works

```
MCP client (Claude / Cursor / any model)
        │  stdio (JSON-RPC)
        ▼
google-slides-mcp  ──►  Google Slides API
                   └──►  Google Drive API (export only)
```

The server does not call an LLM. The client's model chooses tools and arguments; the server validates them with Zod and translates them into Google Slides `batchUpdate` requests.

## Development

```bash
npm run dev        # run from source with tsx
npm run typecheck
npm run build
```

`export_deck` writes files relative to the server process's working directory unless you pass an absolute `outputPath`.

## Security

- Credentials are read from env vars or `~/.config/google-slides-mcp/.env` (mode `600`). Never commit them.
- The server only talks to Google APIs. `batch_update` can do anything the Slides API allows with your token, so only connect MCP clients you trust.
- Revoke access at [myaccount.google.com/permissions](https://myaccount.google.com/permissions).

## License

MIT
