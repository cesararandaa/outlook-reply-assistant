# Reply Assistant

An Outlook web add-in that drafts email replies **in your own writing voice**, using Claude. Reads the open email (and the surrounding thread when EWS is available), generates a draft, and inserts it into the reply window.

Designed for fast iteration: streaming output, quick-action chips for common intents (Confirm / Acknowledge / Decline / etc.), bilingual ES/EN, and a "learn from my sent items" feature that pulls real samples of your own writing from your Sent folder so drafts genuinely sound like you.

License: MIT.

## Architecture

- **Task pane** (Office.js + TypeScript, served by Vite on `https://localhost:3000`) — runs inside Outlook on the web.
- **Backend** (Express + Anthropic SDK on `http://localhost:3001`) — holds the API key, calls Claude with style samples baked into a cached system prompt.
- **`style-samples.json`** — your real reply samples. The model treats these as the source of voice and tone. Add more for better mimicry.

## Setup

```sh
cd reply-assistant
npm install
```

The backend reads `ANTHROPIC_API_KEY` from the environment. If you've already got it exported in your shell (e.g. for gitdash), nothing else to do:

```sh
# In ~/.bashrc / ~/.zshrc — same key gitdash uses
export ANTHROPIC_API_KEY="sk-ant-..."
```

Optional: a project-local `.env` works too (`cp .env.example .env`), but a shell-exported var takes priority and is the recommended path so the key isn't duplicated on disk.

## Run dev

```sh
npm run dev
```

Two processes start: Vite on `https://localhost:3000` (auto-trusted cert via `vite-plugin-mkcert`), Express on `http://localhost:3001`. Vite proxies `/api/*` to the backend, so the task pane only ever sees `https://localhost:3000`.

Sanity check the API: `curl http://localhost:3001/api/health`.

## Sideload into Outlook on the web

1. Open Outlook on the web (`outlook.office.com`).
2. Settings → **Mail** → **Customize actions** → **Get Add-ins** → **My add-ins** → **Custom add-ins** → **Add a custom add-in** → **Add from file**.
3. Pick `manifest.xml` from this repo.
4. Open any email → click **Reply Assistant** in the ribbon → **Draft reply**.

The first time, Outlook will ask you to trust the localhost cert. If the task pane is blank, open the browser dev tools, find the iframe, and check that `https://localhost:3000/taskpane.html` loads in a regular browser tab too.

## Adding your own style samples

Edit `style-samples.json`. Each entry is `{ context, reply, language? }`:

```json
{
  "language": "es",
  "context": "<the email someone sent you>",
  "reply": "<the actual reply you wrote>"
}
```

5–15 real samples gives a much better impression than 3 placeholders. The samples are sent in the system prompt with `cache_control: ephemeral`, so adding more doesn't slow down repeat calls (cache hits show up in `usage.cache_read_input_tokens`).

## Configuration

- `ANTHROPIC_API_KEY` (required) — your Anthropic API key.
- `ANTHROPIC_MODEL` (optional, defaults to `claude-opus-4-7`) — set to `claude-sonnet-4-6` for ~3x faster, ~5x cheaper drafts at slightly lower style fidelity.
- `PORT` (optional, defaults to `3001`) — backend port.

## Production deploy (later)

For real use, host the static `dist/` (after `npm run build`) and the Express server behind HTTPS, update the URLs in `manifest.xml`, and submit it to your tenant's add-in catalog (or AppSource for general distribution).
