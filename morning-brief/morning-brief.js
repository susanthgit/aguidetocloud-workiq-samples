#!/usr/bin/env node
/**
 * Morning brief generator -- Microsoft Work IQ
 * https://github.com/susanthgit/aguidetocloud-workiq-samples
 *
 * Asks Work IQ three orchestrated questions via A2A REST, stitches the answers
 * into a markdown standup brief, and saves it to ./briefs/.
 *
 * Auth: MSAL device-code (one-time browser sign-in) -> cached refresh token ->
 *       silent token acquisition for all subsequent runs.
 *
 * Usage:
 *   $env:WORKIQ_TENANT_ID = "<your-tenant-id-guid>"
 *   node morning-brief.js
 *
 * First run: opens the device-code flow in your terminal. Enter the code on
 * any browser, sign in, done. Token cached at ./.cache/token.json.
 */

const msal = require('@azure/msal-node');
const fs = require('node:fs/promises');
const path = require('node:path');
const crypto = require('node:crypto');

// === Config ===
const TENANT_ID = process.env.WORKIQ_TENANT_ID;
const PROJECT = process.env.PROJECT_NAME || 'Project Adventure';
const TIMEZONE = process.env.WORKIQ_TIMEZONE || 'Pacific/Auckland';
const TIMEZONE_OFFSET_MIN = Number(process.env.WORKIQ_TZ_OFFSET || 720); // NZST default

// Microsoft Work IQ CLI public client app -- already registered in your tenant
// after you ran the Work IQ admin-consent URL.
const CLIENT_ID = 'ba081686-5d24-4bc6-a0d6-d034ecffed87';
const SCOPE = 'api://workiq.svc.cloud.microsoft/.default';
const A2A_ENDPOINT = 'https://workiq.svc.cloud.microsoft/a2a/';
const CACHE_DIR = path.join(__dirname, '.cache');
const CACHE_PATH = path.join(CACHE_DIR, 'token.json');
const OUT_DIR = path.join(__dirname, 'briefs');

if (!TENANT_ID) {
  console.error('ERROR: set WORKIQ_TENANT_ID environment variable to your Microsoft 365 tenant ID (a GUID).');
  console.error('  PowerShell:  $env:WORKIQ_TENANT_ID="00000000-0000-0000-0000-000000000000"');
  console.error('  bash/zsh:    export WORKIQ_TENANT_ID="00000000-..."');
  console.error('');
  console.error('  Find your tenant ID: entra.microsoft.com -> Overview -> Tenant ID');
  process.exit(1);
}

// The 3 questions that make up the brief. Tweak freely.
const PROMPTS = [
  {
    section: `What's new on ${PROJECT}`,
    question: `Summarise everything that's happened with ${PROJECT} in the last 7 days -- meetings, emails, files, Teams chats. Group by theme. Highlight any unresolved follow-ups.`,
  },
  {
    section: `Today's calendar`,
    question: `What's on my calendar today and tomorrow? For each meeting show the title, time, and who else is attending. Keep it concise.`,
  },
  {
    section: `Open commitments`,
    question: `Look at my last 14 days of emails and Teams chats. Find any commitments I made ("I'll send", "I'll follow up", "I'll get back to you") that I haven't acted on yet. List them in order of how overdue they are.`,
  },
];

// === Auth: MSAL device code with persistent cache ===
async function getAccessToken() {
  await fs.mkdir(CACHE_DIR, { recursive: true });
  const cachePlugin = {
    beforeCacheAccess: async (ctx) => {
      try {
        ctx.tokenCache.deserialize(await fs.readFile(CACHE_PATH, 'utf8'));
      } catch {} // fresh on first run
    },
    afterCacheAccess: async (ctx) => {
      if (ctx.cacheHasChanged) {
        await fs.writeFile(CACHE_PATH, ctx.tokenCache.serialize(), 'utf8');
      }
    },
  };

  const pca = new msal.PublicClientApplication({
    auth: { clientId: CLIENT_ID, authority: `https://login.microsoftonline.com/${TENANT_ID}` },
    cache: { cachePlugin },
  });

  // Try cached refresh token first (silent -- no browser)
  const accounts = await pca.getTokenCache().getAllAccounts();
  if (accounts.length > 0) {
    try {
      const result = await pca.acquireTokenSilent({ account: accounts[0], scopes: [SCOPE] });
      return { token: result.accessToken, account: result.account.username };
    } catch {} // fall through to device code
  }

  // Device code fallback (first run, or cache expired)
  console.log('First run -- you need to sign in once to cache a token.');
  const result = await pca.acquireTokenByDeviceCode({
    scopes: [SCOPE],
    deviceCodeCallback: (r) => {
      console.log('');
      console.log('============================================================');
      console.log(`OPEN: ${r.verificationUri}`);
      console.log(`CODE: ${r.userCode}`);
      console.log('============================================================');
      console.log(`(code expires in ${r.expiresIn}s)`);
      console.log('');
    },
  });
  return { token: result.accessToken, account: result.account.username };
}

// === Work IQ A2A: send one message, get the answer text ===
async function askWorkIq(token, question) {
  const body = {
    jsonrpc: '2.0',
    id: crypto.randomUUID(),
    method: 'SendMessage',
    params: {
      message: {
        role: 'ROLE_USER',
        messageId: crypto.randomUUID(),
        parts: [{ text: question }],
        metadata: { Location: { timeZoneOffset: TIMEZONE_OFFSET_MIN, timeZone: TIMEZONE } },
      },
    },
  };
  // One retry on transient fetch failures (Work IQ A2A occasionally drops on first call).
  let lastErr = null;
  for (let attempt = 0; attempt < 2; attempt++) {
    try {
      const res = await fetch(A2A_ENDPOINT, {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/json',
          'A2A-Version': '1.0',
        },
        body: JSON.stringify(body),
      });
      if (!res.ok) {
        throw new Error(`Work IQ A2A returned HTTP ${res.status}: ${(await res.text()).slice(0, 500)}`);
      }
      const json = await res.json();
      if (json.error) throw new Error(`Work IQ error: ${json.error.message || JSON.stringify(json.error)}`);
      const artifacts = json.result?.task?.artifacts || [];
      const textParts = artifacts.flatMap((a) => (a.parts || []).filter((p) => p.text)).map((p) => p.text);
      return textParts.join('\n\n') || '(Work IQ returned no text.)';
    } catch (err) {
      lastErr = err;
      if (attempt === 0 && /fetch failed|ECONNRESET|ETIMEDOUT/i.test(err.message)) {
        // Transient -- brief pause + retry once
        await new Promise((r) => setTimeout(r, 1500));
        continue;
      }
      throw err;
    }
  }
  throw lastErr;
}

// === Clean Work IQ markdown for the brief file ===
// Work IQ inlines:
//   - long base64 OWA / Teams / SharePoint URLs in [label](url) form
//   - Unicode citation markers like 【1-aaea3e】
// Neither helps in a markdown file -- strip both, keep the readable label.
function cleanForBrief(text) {
  return text
    // [label](http...) -> **label** (handles nested-bracket case Work IQ emits)
    .replace(/\[([^\[\]]*(?:\[[^\]]*\][^\[\]]*)*)\]\(https?:\/\/[^)]+\)/g, '**$1**')
    // strip citation markers 【...】
    .replace(/\u3010[^\u3011]+\u3011/g, '')
    // tidy double-spaces / double-blank-lines left over
    .replace(/ {2,}/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

// === Main ===
async function main() {
  const divider = '━'.repeat(58);
  console.log('');
  console.log(divider);
  console.log('  📋 Morning brief generator -- Microsoft Work IQ');
  console.log(divider);
  console.log(`  tenant:  ${TENANT_ID}`);
  console.log(`  project: ${PROJECT}`);
  console.log('');

  const totalStart = Date.now();
  const { token, account } = await getAccessToken();
  console.log(`  ✓ Authenticated as ${account}`);
  console.log(`  ⏳ Asking ${PROMPTS.length} questions via Work IQ A2A...`);
  console.log('');

  const sections = [];
  const timings = [];
  for (const { section, question } of PROMPTS) {
    process.stdout.write(`     → ${section.padEnd(34)} `);
    const start = Date.now();
    try {
      const answer = await askWorkIq(token, question);
      const elapsed = ((Date.now() - start) / 1000).toFixed(1);
      console.log(`✓ ${elapsed}s`);
      sections.push({ section, answer });
      timings.push(Number(elapsed));
    } catch (err) {
      console.log(`✗ ${err.message}`);
      sections.push({ section, answer: `_(Work IQ error: ${err.message})_` });
    }
  }

  // Stitch into markdown brief
  const today = new Date().toISOString().slice(0, 10);
  const lines = [
    `# Morning brief -- ${today}`,
    ``,
    `_Account: ${account}_`,
    `_Generated by [aguidetocloud-workiq-samples](https://github.com/susanthgit/aguidetocloud-workiq-samples) at ${new Date().toLocaleString()}_`,
    ``,
  ];
  for (const { section, answer } of sections) {
    lines.push(`## ${section}`, ``, cleanForBrief(answer), ``);
  }

  await fs.mkdir(OUT_DIR, { recursive: true });
  const accountSafe = account.replace(/[^a-zA-Z0-9]/g, '_');
  const outPath = path.join(OUT_DIR, `brief-${accountSafe}-${today}.md`);
  await fs.writeFile(outPath, lines.join('\n'), 'utf8');

  const totalSec = ((Date.now() - totalStart) / 1000).toFixed(1);
  const briefSize = (await fs.stat(outPath)).size;

  console.log('');
  console.log(divider);
  console.log(`  ✅ Brief ready in ${totalSec}s  (${(briefSize / 1024).toFixed(1)} KB markdown)`);
  console.log(divider);
  console.log(`     ${outPath}`);
  console.log('');
  console.log(`     Open it · email it · post to Teams · render as HTML.`);
  console.log(`     It's just markdown.`);
  console.log('');
}

main().catch((err) => {
  console.error('');
  console.error('[FAIL] Brief generation failed:', err.message);
  process.exit(1);
});
