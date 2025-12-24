/********************************************************************************************
 * InnsynAI Teams Bridge – BASELINE (GLOBAL BOT, CROSS-TENANT SAFE)
 *
 * PURPOSE:
 * - Prove Teams → Bot → Teams roundtrip works
 * - Works in HOME tenant AND EXTERNAL tenants
 * - Marketplace-correct auth model
 ********************************************************************************************/

import { serve } from "https://deno.land/std@0.224.0/http/server.ts";
import * as jose from "https://deno.land/x/jose@v5.4.0/index.ts";

/********************************************************************************************
 * ENV
 ********************************************************************************************/
const env = Deno.env.toObject();

const INTERNAL_LOOKUP_SECRET = env.INTERNAL_LOOKUP_SECRET;
const TEAMS_TENANT_LOOKUP_URL = env.TEAMS_TENANT_LOOKUP_URL;
const SUPABASE_ANON_KEY = env.SUPABASE_ANON_KEY;
const TEAMS_BOT_APP_ID = env.TEAMS_BOT_APP_ID;
const TEAMS_BOT_APP_PASSWORD = env.TEAMS_BOT_APP_PASSWORD;

if (
  !INTERNAL_LOOKUP_SECRET ||
  !TEAMS_TENANT_LOOKUP_URL ||
  !SUPABASE_ANON_KEY ||
  !TEAMS_BOT_APP_ID ||
  !TEAMS_BOT_APP_PASSWORD
) {
  console.error("❌ Missing required env vars");
  Deno.exit(1);
}

/********************************************************************************************
 * BOTFRAMEWORK JWKS
 ********************************************************************************************/
const OPENID_CONFIG_URL =
  "https://login.botframework.com/v1/.well-known/openidconfiguration";

let jwks: jose.JSONWebKeySet | null = null;

async function getJwks() {
  if (jwks) return jwks;
  const meta = await fetch(OPENID_CONFIG_URL).then(r => r.json());
  jwks = await fetch(meta.jwks_uri).then(r => r.json());
  return jwks!;
}

/********************************************************************************************
 * JWT VERIFY (INBOUND)
 * IMPORTANT: validate against the *actual aud* sent by Teams
 ********************************************************************************************/
async function verifyJwt(authHeader: string): Promise<string> {
  const token = authHeader.slice(7);
  const decoded = jose.decodeJwt(token);

  const botAppId = decoded.aud as string;

  console.log("🔐 Incoming JWT", {
    aud: botAppId,
    iss: decoded.iss,
    tid: decoded.tid,
  });

  const keyStore = jose.createLocalJWKSet(await getJwks());

  await jose.jwtVerify(token, keyStore, {
    issuer: "https://api.botframework.com",
    audience: botAppId,
  });

  return botAppId;
}

/********************************************************************************************
 * TENANT LOOKUP (AUTO-PROVISION)
 ********************************************************************************************/
async function resolveTenant(aadTenantId: string): Promise<string | null> {
  const res = await fetch(TEAMS_TENANT_LOOKUP_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      apikey: SUPABASE_ANON_KEY,
      "x-internal-token": INTERNAL_LOOKUP_SECRET,
    },
    body: JSON.stringify({
      teams_tenant_id: aadTenantId,
      auto_provision: true,
    }),
  });

  if (!res.ok) {
    console.error("❌ Tenant lookup failed", await res.text());
    return null;
  }

  const json = await res.json();
  return json?.tenant_id ?? null;
}

/********************************************************************************************
 * BOT TOKEN (GLOBAL BOTFRAMEWORK AUTHORITY – REQUIRED)
 ********************************************************************************************/
async function getBotToken(): Promise<string> {
  console.log("🔑 Minting bot token (botframework.com)");

  const res = await fetch(
    "https://login.microsoftonline.com/botframework.com/oauth2/v2.0/token",
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type: "client_credentials",
        client_id: TEAMS_BOT_APP_ID,
        client_secret: TEAMS_BOT_APP_PASSWORD,
        scope: "https://api.botframework.com/.default",
      }),
    },
  );

  const json = await res.json();

  if (!json.access_token) {
    console.error("❌ Bot token mint failed", json);
    throw new Error("bot token failure");
  }

  return json.access_token;
}

/********************************************************************************************
 * MAIN HANDLER
 ********************************************************************************************/
async function handleTeams(req: Request): Promise<Response> {
  if (req.method !== "POST") return new Response("ok");

  let activity: any;
  try {
    activity = JSON.parse(await req.text());
  } catch {
    return new Response("ok");
  }

  const auth = req.headers.get("Authorization");
  if (!auth?.startsWith("Bearer ")) return new Response("ok");

  const botAppId = await verifyJwt(auth);

  if (botAppId !== TEAMS_BOT_APP_ID) {
    console.warn("⚠️ JWT aud does not match configured bot id");
    return new Response("ok");
  }

  const aadTenantId =
    activity.channelData?.tenant?.id || activity.conversation?.tenantId;

  console.log("📨 Activity received", {
    aadTenantId,
    conversationType: activity.conversation?.conversationType,
    text: activity.text,
  });

  if (!aadTenantId || !activity.serviceUrl || !activity.conversation?.id) {
    return new Response("ok");
  }

  const tenantId = await resolveTenant(aadTenantId);
  if (!tenantId) return new Response("ok");

  const serviceUrl = String(activity.serviceUrl).replace(/\/$/, "");
  const token = await getBotToken();

  const postUrl =
    serviceUrl +
    "/v3/conversations/" +
    encodeURIComponent(activity.conversation.id) +
    "/activities";

  const res = await fetch(postUrl, {
    method: "POST",
    headers: {
      Authorization: "Bearer " + token,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      type: "message",
      text: "Hello from InnsynAI 👋",
      from: activity.recipient,
      recipient: activity.from,
      conversation: { id: activity.conversation.id },
    }),
  });

  if (!res.ok) {
    console.error("❌ Teams send failed", res.status, await res.text());
  } else {
    console.log("✅ Message sent to Teams");
  }

  return new Response("ok");
}

/********************************************************************************************
 * SERVER
 ********************************************************************************************/
serve((req) => {
  const path = new URL(req.url).pathname;
  console.log("➡️ Request", req.method, path);
  if (path === "/teams") return handleTeams(req);
  return new Response("ok");
});
