# Morning brief generator

~190-line Node.js script that asks Microsoft Work IQ three questions every morning and writes a markdown standup brief to disk.

## What it does

When you run `node morning-brief.js`, the script:
1. Authenticates against Microsoft via **MSAL device-code flow** (browser sign-in, one-time, then cached)
2. Acquires a bearer token for `api://workiq.svc.cloud.microsoft/.default`
3. Calls the Work IQ **A2A REST endpoint** three times with orchestrated questions:
   - *"Summarise everything that's happened with Project X in the last 7 days"*
   - *"What's on my calendar today and tomorrow?"*
   - *"What commitments have I made that I haven't followed up on?"*
4. Stitches the answers into a markdown file at `./briefs/brief-<account>-<date>.md`

Schedule it with **Windows Task Scheduler** or **cron** for a fresh brief every morning.

## Why direct A2A REST (not MCP or CLI)

Three reasons:
1. **No WAM dependency** — works in any terminal (CI/CD, Cloud Shell, headless server). The `workiq` CLI requires a Windows window handle for MSAL WAM auth; A2A REST with device-code doesn't.
2. **Account routing is explicit** — the MSAL token's UPN claim IS the auth identity. No risk of silently using a different cached account.
3. **Future-proof** — when Microsoft ships their public REST API (post-GA), it'll be the same shape with a different endpoint URL.

## Setup

```powershell
# 1. Install deps
cd morning-brief
npm install

# 2. Set your tenant ID (find it: entra.microsoft.com -> Overview -> Tenant ID)
$env:WORKIQ_TENANT_ID = "00000000-0000-0000-0000-000000000000"

# 3. Optional: pick a project (default "Project Adventure")
$env:PROJECT_NAME = "Project Phoenix"

# 4. Optional: timezone for "today/tomorrow" calculations
$env:WORKIQ_TIMEZONE = "Pacific/Auckland"
$env:WORKIQ_TZ_OFFSET = "720"

# 5. Run it
node morning-brief.js
```

**First run only:** the script prints a device code + URL. Open the URL in any browser, paste the code, sign in as your work account. Token cached at `./.cache/token.json` (gitignored).

**Subsequent runs:** silent. No browser, no prompt.

## Schedule it daily

### Windows (Task Scheduler)

```powershell
$action = New-ScheduledTaskAction `
  -Execute 'node.exe' `
  -Argument 'morning-brief.js' `
  -WorkingDirectory (Resolve-Path .).Path
$trigger = New-ScheduledTaskTrigger -Daily -At 8:00am
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType S4U
Register-ScheduledTask -TaskName "WorkIQ Morning Brief" `
  -Action $action -Trigger $trigger -Principal $principal
```

### Linux / macOS (cron)

```bash
crontab -e
# Add:
0 8 * * * cd /path/to/morning-brief && WORKIQ_TENANT_ID="..." node morning-brief.js
```

## Extend it

The brief is just markdown -- easy to wire into anything:

| Want to... | Add |
|---|---|
| Email the brief | `nodemailer` -> `await sendMail({ html: marked(briefMd) })` |
| Post to Teams | POST to a Teams Incoming Webhook with `{ "text": briefMd }` |
| Post to Slack | Same pattern, Slack webhook |
| Render as HTML | Pipe through `marked` or `markdown-it`, save as `.html` |
| Add more questions | Add to the `PROMPTS` array in `morning-brief.js` |
| Different project | `$env:PROJECT_NAME = "your project"` |

## Sample output (excerpt)

```markdown
# Morning brief -- 2026-06-11

_Account: you@your-tenant.onmicrosoft.com_

## What's new on Project Adventure

# 1) Kickoff, Scope & Planning
**What happened**
- Kickoff initiated via [Project Adventure -- Kickoff](...) with:
  - Customer: Trailhead Outdoor (new portal engagement)
  - Topics: scope, team roles, cadence, key dates
- Initial scope captured...

**Unresolved follow-ups**
- Hosting decision: Azure vs AWS
- Identity approach (high-level) still open
- Launch timeline misalignment...

# 2) Pricing & Commercial Strategy
[... etc, 6 themed sections with citations ...]

## Today's calendar
[meetings list or honest "nothing scheduled today"]

## Open commitments
[ranked overdue commitments parsed from email + Teams chats]
```

## Honest take

This script is intentionally small and unopinionated. Production runs would add: retries on transient errors, structured logging, observability, per-tenant config, secrets management. None of that is here -- it's a starting point for you to lift and adapt.

If you ship something interesting on top of this, [send me a link](https://www.aguidetocloud.com/feedback/).
