#!/usr/bin/env node
/**
 * Work IQ web-app -- Express server + browser frontend
 * https://github.com/susanthgit/aguidetocloud-workiq-samples
 *
 * - Authenticates once via MSAL device code (browser sign-in), caches refresh token
 * - Serves /public for the HTML/CSS/JS frontend
 * - POST /api/ask {question, conversationId?} -> {answer, account, contextId, latency_ms}
 * - Calls Work IQ A2A REST endpoint directly (no CLI, no MCP subprocess)
 *
 * Run:  npm start
 * Open: http://localhost:3001
 */

const express = require('express');
const path = require('node:path');
const fs = require('node:fs/promises');
const crypto = require('node:crypto');
const msal = require('@azure/msal-node');

// === Config ===
const TENANT_ID = process.env.WORKIQ_TENANT_ID;
const PORT = Number(process.env.PORT) || 3001;
const TIMEZONE = process.env.WORKIQ_TIMEZONE || 'Pacific/Auckland';
const TIMEZONE_OFFSET_MIN = Number(process.env.WORKIQ_TZ_OFFSET || 720);

const CLIENT_ID = 'ba081686-5d24-4bc6-a0d6-d034ecffed87';
const SCOPE = 'api://workiq.svc.cloud.microsoft/.default';
const A2A_ENDPOINT = 'https://workiq.svc.cloud.microsoft/a2a/';
const CACHE_DIR = path.join(__dirname, '.cache');
const CACHE_PATH = path.join(CACHE_DIR, 'token.json');

if (!TENANT_ID) {
  console.error('ERROR: set WORKIQ_TENANT_ID environment variable to your Microsoft 365 tenant ID.');
  console.error('  PowerShell:  $env:WORKIQ_TENANT_ID="00000000-0000-0000-0000-000000000000"; npm start');
  console.error('  bash/zsh:    WORKIQ_TENANT_ID="..." npm start');
  process.exit(1);
}

// === Auth: MSAL with persistent cache ===
let pca = null;
let cachedAccount = null;

async function initAuth() {
  await fs.mkdir(CACHE_DIR, { recursive: true });
  const cachePlugin = {
    beforeCacheAccess: async (ctx) => {
      try { ctx.tokenCache.deserialize(await fs.readFile(CACHE_PATH, 'utf8')); } catch {}
    },
    afterCacheAccess: async (ctx) => {
      if (ctx.cacheHasChanged) await fs.writeFile(CACHE_PATH, ctx.tokenCache.serialize(), 'utf8');
    },
  };
  pca = new msal.PublicClientApplication({
    auth: { clientId: CLIENT_ID, authority: `https://login.microsoftonline.com/${TENANT_ID}` },
    cache: { cachePlugin },
  });

  const accounts = await pca.getTokenCache().getAllAccounts();
  if (accounts.length > 0) {
    try {
      const r = await pca.acquireTokenSilent({ account: accounts[0], scopes: [SCOPE] });
      cachedAccount = r.account;
      return;
    } catch {}
  }

  console.log('First run -- sign in once to cache a token.');
  const r = await pca.acquireTokenByDeviceCode({
    scopes: [SCOPE],
    deviceCodeCallback: (resp) => {
      console.log('');
      console.log('============================================================');
      console.log(`OPEN: ${resp.verificationUri}`);
      console.log(`CODE: ${resp.userCode}`);
      console.log('============================================================');
      console.log(`(code expires in ${resp.expiresIn}s)`);
      console.log('');
    },
  });
  cachedAccount = r.account;
}

async function getToken() {
  const r = await pca.acquireTokenSilent({ account: cachedAccount, scopes: [SCOPE] });
  return r.accessToken;
}

// === Work IQ A2A ===
async function askWorkIq(token, question, contextId = null) {
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
        ...(contextId ? { contextId } : {}),
      },
    },
  };
  const res = await fetch(A2A_ENDPOINT, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
      'A2A-Version': '1.0',
    },
    body: JSON.stringify(body),
  });
  if (!res.ok) throw new Error(`HTTP ${res.status}: ${(await res.text()).slice(0, 300)}`);
  const json = await res.json();
  if (json.error) throw new Error(json.error.message || JSON.stringify(json.error));

  const task = json.result?.task || {};
  const artifacts = task.artifacts || [];
  const textParts = artifacts.flatMap((a) => (a.parts || []).filter((p) => p.text)).map((p) => p.text);
  return { answer: textParts.join('\n\n') || '(no text)', contextId: task.contextId };
}

// === Express server ===
const app = express();
app.use(express.json({ limit: '64kb' }));
app.use(express.static(path.join(__dirname, 'public')));

app.get('/api/health', (req, res) => {
  res.json({ ok: !!pca, account: cachedAccount?.username || null });
});

const conversations = new Map(); // conversationId -> Work IQ contextId

app.post('/api/ask', async (req, res) => {
  const question = String(req.body?.question || '').trim();
  const conversationId = String(req.body?.conversationId || '').trim() || crypto.randomUUID();
  if (!question) return res.status(400).json({ error: 'question required' });

  const start = Date.now();
  try {
    const token = await getToken();
    const { answer, contextId } = await askWorkIq(token, question, conversations.get(conversationId));
    if (contextId) conversations.set(conversationId, contextId);
    res.json({
      answer,
      account: cachedAccount?.username,
      conversationId,
      latency_ms: Date.now() - start,
    });
  } catch (err) {
    console.error('  [FAIL] ask:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// === Start ===
(async () => {
  await initAuth();
  app.listen(PORT, () => {
    console.log('');
    console.log(`Work IQ web-app live at http://localhost:${PORT}`);
    console.log(`  Account: ${cachedAccount?.username}`);
    console.log(`  Press Ctrl+C to stop.`);
    console.log('');
  });
})();

process.on('SIGINT', () => {
  console.log('\nShutting down...');
  process.exit(0);
});
