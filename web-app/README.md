# Work IQ web-app

Tiny Express server + single-page HTML interface for asking Microsoft Work IQ questions in a browser.

Demonstrates the **embed pattern** for adding Work IQ to any internal web app — chat with persistent conversation context, no CLI required for the end user.

## What it does

- Authenticates once via **MSAL device-code flow** (browser sign-in, one-time, then cached)
- Calls Work IQ **A2A REST endpoint** directly (no CLI, no MCP subprocess)
- Express POST `/api/ask {question, conversationId?}` -> `{answer, account, conversationId, latency_ms}`
- Multi-turn conversation support via Work IQ's `contextId` (per-conversation state held server-side)
- Single-page HTML/CSS/JS frontend, ~200 lines, no framework, no build step

## Setup

```powershell
# 1. Install
cd web-app
npm install

# 2. Set your tenant ID
$env:WORKIQ_TENANT_ID = "00000000-0000-0000-0000-000000000000"

# 3. Run
npm start
```

Open **http://localhost:3001** in your browser.

**First run only:** server prints a device code + URL in the terminal. Open URL, paste code, sign in. Token cached at `./.cache/token.json` (gitignored). After that the server starts silently.

## File map

| File | Lines | Purpose |
|---|---|---|
| `server.js` | ~150 | Express + MSAL + A2A REST + per-conversation context |
| `public/index.html` | ~200 | Single-page UI — input, answer cards, example chips |
| `package.json` | 20 | Two deps: `express` + `@azure/msal-node` |

## Multi-turn conversations

The server holds a `Map<conversationId, workIqContextId>` and passes `contextId` on follow-up questions, giving you real conversational continuity:

```js
let conversationId = null;
async function ask(question) {
  const res = await fetch('/api/ask', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ question, conversationId })
  });
  const data = await res.json();
  conversationId = data.conversationId; // reuse on follow-ups
  return data.answer;
}
```

The bundled frontend uses one conversation per page load. Refresh = fresh conversation.

## Production hardening to add

This is intentionally a starting point. Before you put it in front of real users:

- **Multi-user auth** — current impl is single-account. For multi-user, OAuth2 federation per session + per-user MSAL state.
- **Streaming responses** — A2A supports `SendStreamingMessage` (SSE). Wire that for snappier UX on long answers.
- **Rate limiting** — Copilot Credits cost money. Add `express-rate-limit` per session.
- **Error categorisation** — auth errors, timeouts, quota exceeded each deserve different UX.
- **Markdown rendering** — Work IQ answers are markdown; pipe through `marked` for proper rendering of headers, lists, links.
- **Citation expansion** — A2A responses include `uncite-*` artifacts with source URLs + sensitivity labels. The current frontend ignores these; surface as expandable chips for trust.

## Sample interactions

| Ask | Work IQ returns |
|---|---|
| *"What's on my calendar today?"* | Today's meetings with attendees + times |
| *"Summarise everything about Project Adventure in the last 7 days."* | Cross-source brief (mail + meetings + chat + docs) with citations |
| *"What commitments have I made that I haven't followed up on?"* | Action-item list across email + Teams chats |
| *"Find emails from Mario about the auth spike."* | Filtered email list with sender + date + snippet |

## Honest take

~350 lines of code total. Small enough to read end-to-end in 10 minutes. The MSAL device-code + A2A REST pattern is the same one you'd use in production; we've stripped the surrounding scaffolding.

Lift it, adapt it, ship something good. [Send me a link](https://www.aguidetocloud.com/feedback/) if you do.
