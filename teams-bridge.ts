/********************************************************************************************
 * InnsynAI Teams Bridge – BASELINE (MATCHES WORKING TOKEN FLOW)
 *
 * PURPOSE:
 * - Prove Teams → Bot → Teams roundtrip works
 * - No PATCH
 * - No RAG
 * - No Adaptive Cards
 *
 * Key change vs your failing baseline:
 * - Mint bot token from customer tenant authority (WORKED in your per-tenant bot version)
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
  const meta = await fetch(OPENID_CONFIG_URL).then((r) => r.json());
  jwks = await fetch(meta.jwks_uri).then((r) => r.json());
  return jwks!;
}

/********************************************************************************************
 * JWT VERIFY (INBOUND)
 ********************************************************************************************/
async function verifyJwt(authHeader: string) {
  const token = authHeader.slice(7);
  const decoded = jose.decodeJwt(token);

  console.log("🔐 Incoming JWT", {
    aud: decoded.aud,
    iss: decoded.iss,
    tid: decoded.tid,
  });

  const keyStore = jose.createLocalJWKSet(await getJwks());

  await jose.jwtVerify(token, keyStore, {
    issuer: "https://api.botframework.com",
    audience: TEAMS_BOT_APP_ID,
  });
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
 * BOT TOKEN (CUSTOMER TENANT AUTHORITY — this matches your working version)
 ********************************************************************************************/
async function getBotToken(aadTenantId: string): Promise<string> {
  console.log("🔑 Minting bot token", {
    authority: aadTenantId,
    client_id: TEAMS_BOT_APP_ID,
  });

  const res = await fetch(
    `https://login.microsoftonline.com/${aadTenantId}/oauth2/v2.0/token`,
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

  console.log("✅ Bot token minted");
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
  if (!auth || !auth.startsWith("Bearer ")) return new Response("ok");

  await verifyJwt(auth);

  const aadTenantId =
    activity.channelData?.tenant?.id || activity.conversation?.tenantId;

  console.log("📨 Activity received", {
    aadTenantId,
    conversationType: activity.conversation?.conversationType,
    serviceUrl: activity.serviceUrl,
    conversationId: activity.conversation?.id,
    activityId: activity.id,
    replyToId: activity.replyToId,
    text: activity.text,
  });

  if (!aadTenantId || !activity.serviceUrl || !activity.conversation?.id) {
    console.warn("⚠️ Missing required activity fields");
    return new Response("ok");
  }

  const tenantId = await resolveTenant(aadTenantId);
  console.log("🧭 Tenant resolved", tenantId);
  if (!tenantId) return new Response("ok");

  // Preserve exact serviceUrl, but avoid double slashes
  const serviceUrl = String(activity.serviceUrl).replace(/\/$/, "");

  const token = await getBotToken(aadTenantId);

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
      // This mirrors the version that worked
      replyToId: activity.replyToId ?? activity.id,
    }),
  });

  const bodyText = await res.text();
  if (!res.ok) {
    console.error("❌ Teams send failed", res.status, bodyText);

    // Helpful: many times the real clue is in WWW-Authenticate
    const wwwAuth = res.headers.get("www-authenticate");
    if (wwwAuth) console.error("🔎 WWW-Authenticate:", wwwAuth);
  } else {
    console.log("✅ Message sent to Teams", bodyText);
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
