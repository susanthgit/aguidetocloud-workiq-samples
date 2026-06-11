# aguidetocloud-workiq-samples

Two minimal, ready-to-run samples that show what you can actually *build* on the Microsoft Work IQ API beyond asking a chatbot questions.

Companion code for [the Day-1 Work IQ guide on aguidetocloud.com](https://www.aguidetocloud.com/blog/microsoft-work-iq-api-day-1-ga/).

## What's inside

| Folder | What it is | Run it |
|---|---|---|
| [`morning-brief/`](./morning-brief) | 90-line Node.js script that asks Work IQ 3 questions (project status · today's calendar · open commitments), stitches the answers into a markdown brief, and saves it to disk. Schedule it daily with Task Scheduler or cron and you've got an automatic morning standup. | `cd morning-brief && node morning-brief.js` |
| [`web-app/`](./web-app) | Tiny Express server + single-page HTML interface that lets you ask Work IQ questions in a browser. Demonstrates the embed pattern for adding Work IQ to any internal app. | `cd web-app && npm start` then open `http://localhost:3001` |

Both samples talk to Work IQ through the **Model Context Protocol (MCP)** — the same wire protocol Copilot CLI and VS Code use. So everything you see here is exactly how a production integration would talk to Work IQ.

## Before you run either sample

Three quick things:

1. **Node.js 18 or later** — check with `node --version`. Get it from [nodejs.org](https://nodejs.org) if missing.
2. **Microsoft 365 Copilot licence** on the account you'll authenticate as — Work IQ rides on Copilot.
3. **Accept the Work IQ EULA once** for the account you'll use — from a real windowed terminal (Windows Terminal / pwsh in a normal window — *not* an embedded terminal in an IDE if you hit MSAL WAM errors):
   ```bash
   npx -y @microsoft/workiq --account <your-email> accept-eula
   ```
   A browser pops once, you sign in, you accept. Both samples reuse the cached token after that.

## Set your account once

Both samples read `WORKIQ_ACCOUNT` from the environment so the auth target is unambiguous:

```powershell
# PowerShell (persists for this terminal session)
$env:WORKIQ_ACCOUNT = "your-email@your-tenant.onmicrosoft.com"
```

```bash
# bash / zsh
export WORKIQ_ACCOUNT="your-email@your-tenant.onmicrosoft.com"
```

Or pass it inline:
```powershell
$env:WORKIQ_ACCOUNT="..."; node morning-brief.js
```

## Honest take

These samples deliberately stay small and unopinionated so you can lift and adapt them. Real production code adds: retries, structured logging, rate-limit handling, per-tenant config, secret stores, observability. None of that is here — these are starting points, not reference architectures.

If you ship something interesting built on top of these, [send me a link](https://www.aguidetocloud.com/feedback/) — I'd love to see what the community builds first.

## Licence

MIT — see [LICENSE](./LICENSE).
