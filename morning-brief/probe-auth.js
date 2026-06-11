#!/usr/bin/env node
/**
 * Probe: can we use Work IQ CLI's app id (ba081686-...) as a public client
 * via MSAL device-code flow, then POST to Work IQ A2A?
 *
 * If YES → that's the architecture for our blog samples.
 * If NO → register our own client app, document setup.
 */
const msal = require('@azure/msal-node');
const fs = require('node:fs/promises');
const path = require('node:path');

const TENANT_ID = '00b98149-2e3e-468c-b063-fb0cfa35fe44'; // CDX
const CLIENT_ID = 'ba081686-5d24-4bc6-a0d6-d034ecffed87'; // Work IQ CLI app
const SCOPE = 'api://workiq.svc.cloud.microsoft/.default';
const A2A_ENDPOINT = 'https://workiq.svc.cloud.microsoft/a2a/';
const CACHE_PATH = path.join(__dirname, '.probe-token-cache.json');

// Token cache plugin
const cachePlugin = {
  beforeCacheAccess: async (ctx) => {
    try {
      const data = await fs.readFile(CACHE_PATH, 'utf8');
      ctx.tokenCache.deserialize(data);
    } catch {} // fresh
  },
  afterCacheAccess: async (ctx) => {
    if (ctx.cacheHasChanged) {
      await fs.writeFile(CACHE_PATH, ctx.tokenCache.serialize(), 'utf8');
    }
  },
};

async function main() {
  console.log('=== Probe: MSAL device-code + Work IQ A2A ===');
  console.log(`tenant: ${TENANT_ID}`);
  console.log(`clientId: ${CLIENT_ID} (Work IQ CLI)`);
  console.log(`scope: ${SCOPE}`);
  console.log('');

  const pca = new msal.PublicClientApplication({
    auth: {
      clientId: CLIENT_ID,
      authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    },
    cache: { cachePlugin },
  });

  // Try silent first (uses cached refresh token if present)
  let result = null;
  const accounts = await pca.getTokenCache().getAllAccounts();
  if (accounts.length > 0) {
    console.log(`Cached account found: ${accounts[0].username} — trying silent token acquisition...`);
    try {
      result = await pca.acquireTokenSilent({
        account: accounts[0],
        scopes: [SCOPE],
      });
      console.log(`[OK] Silent token acquired. Expires: ${result.expiresOn}`);
    } catch (err) {
      console.log(`Silent failed (${err.errorCode || err.message}), falling back to device code.`);
    }
  }

  // Device code if needed
  if (!result) {
    console.log('Starting device code flow...');
    result = await pca.acquireTokenByDeviceCode({
      scopes: [SCOPE],
      deviceCodeCallback: (response) => {
        console.log('');
        console.log('============================================================');
        console.log(`OPEN: ${response.verificationUri}`);
        console.log(`CODE: ${response.userCode}`);
        console.log('============================================================');
        console.log('Sign in as: admin@M365CPI52224224.onmicrosoft.com');
        console.log(`(code expires in ${response.expiresIn}s)`);
        console.log('');
      },
    });
    console.log('[OK] Device code flow succeeded.');
  }

  // Decode token to verify audience + scopes
  const payload = result.accessToken.split('.')[1];
  let pad = payload;
  while (pad.length % 4) pad += '=';
  const claims = JSON.parse(
    Buffer.from(pad.replace(/-/g, '+').replace(/_/g, '/'), 'base64').toString('utf8')
  );
  console.log('');
  console.log('=== Token claims ===');
  console.log(`audience: ${claims.aud}`);
  console.log(`upn: ${claims.upn || claims.preferred_username}`);
  console.log(`appid: ${claims.appid}`);
  console.log(`scp: ${claims.scp}`);
  console.log('');

  // Test A2A POST
  console.log(`=== Testing A2A POST to ${A2A_ENDPOINT} ===`);
  const body = {
    jsonrpc: '2.0',
    id: crypto.randomUUID(),
    method: 'SendMessage',
    params: {
      message: {
        role: 'ROLE_USER',
        messageId: crypto.randomUUID(),
        parts: [{ text: 'Just say "hello" and nothing else.' }],
        metadata: {
          Location: {
            timeZoneOffset: 720, // NZST +12h
            timeZone: 'Pacific/Auckland',
          },
        },
      },
    },
  };
  const start = Date.now();
  const res = await fetch(A2A_ENDPOINT, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${result.accessToken}`,
      'Content-Type': 'application/json',
      'A2A-Version': '1.0',
    },
    body: JSON.stringify(body),
  });
  const elapsed = ((Date.now() - start) / 1000).toFixed(1);
  console.log(`HTTP ${res.status} ${res.statusText} (${elapsed}s)`);
  const text = await res.text();
  if (res.ok) {
    console.log('[OK] A2A endpoint responded.');
    try {
      const json = JSON.parse(text);
      console.log('Response (truncated):');
      console.log(JSON.stringify(json, null, 2).slice(0, 1500));
    } catch {
      console.log('Response (non-JSON):', text.slice(0, 1000));
    }
  } else {
    console.log('[FAIL] Response body:');
    console.log(text.slice(0, 2000));
  }
}

main().catch((err) => {
  console.error('Probe failed:', err.message);
  if (err.errorCode) console.error('  errorCode:', err.errorCode);
  process.exit(1);
});
