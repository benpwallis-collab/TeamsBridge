/********************************************************************************************
 * InnsynAI Teams Bridge – FINAL (STORE / ADD-TO-TEAMS SAFE)
 *
 * ✅ Single global bot identity
 * ✅ Multi-tenant routing via AAD tenant id
 * ✅ Auto-provision tenant mapping
 * ✅ Uses EXACT serviceUrl from Teams
 * ✅ Tenant-specific OAuth authority (CRITICAL FIX)
 ********************************************************************************************/

import { serve } from "https://deno.land/std@0.224.0/http/server.ts";
import * as jose from "https://deno.land/x/jose@v5.4.0/index.ts";

/********************************************************************************************
 * ENV
 ********************************************************************************************/
const {
  INTERNAL_LOOKUP_SECRET,
  TEAMS_TENANT_LOOKUP_URL,
  RAG_QUERY_URL,
  SUPABASE_ANON_KEY,
  TEAMS_BOT_APP_ID,
  TEAMS_BOT_APP_PASSWORD,
} = Deno.env.toObject();

if (
  !INTERNAL_LOOKUP_SECRET ||
  !TEAMS_TENANT_LOOKUP_URL ||
  !RAG_QUERY_URL ||
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
    console.error("❌ tenant lookup failed", await res.text());
    return null;
  }

  const json = await res.json();
  return json?.tenant_id ?? null;
}

/********************************************************************************************
 * JWT VERIFY
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
 * BOT TOKEN (TENANT-SPECIFIC AUTHORITY — REQUIRED)
 ********************************************************************************************/
async function getBotToken(aadTenantId: string) {
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
    console.error("❌ Token mint failed", json);
    throw new Error("bot token failure");
  }

  console.log("✅ Bot token minted for tenant", aadTenantId);
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

  await verifyJwt(auth);

  const aadTenantId =
    activity.channelData?.tenant?.id || activity.conversation?.tenantId;

  console.log("📨 Activity received", {
    botAppId: TEAMS_BOT_APP_ID,
    aadTenantId,
    serviceUrl: activity.serviceUrl,
    conversationId: activity.conversation?.id,
  });

  if (!aadTenantId || !activity.serviceUrl || !activity.conversation?.id) {
    console.warn("⚠️ Missing required activity fields");
    return new Response("ok");
  }

  const tenantId = await resolveTenant(aadTenantId);
  console.log("🧭 Tenant resolved", tenantId);

  if (!tenantId) return new Response("ok");

  // 🔒 DO NOT REWRITE serviceUrl
  const serviceUrl = activity.serviceUrl.replace(/\/$/, "");

  const token = await getBotToken(aadTenantId);

  /****************************
   * SEND PLACEHOLDER
   ****************************/
  const postUrl =
    `${serviceUrl}/v3/conversations/${encodeURIComponent(
      activity.conversation.id,
    )}/activities`;

  const placeholderRes = await fetch(postUrl, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      type: "message",
      text: "⏳ Working on it…",
      replyToId: activity.replyToId ?? activity.id,
    }),
  );

  if (!placeholderRes.ok) {
    console.error("❌ TEAMS API ERROR", {
      status: placeholderRes.status,
      body: await placeholderRes.text(),
    });
    return new Response("ok");
  }

  const placeholder = await placeholderRes.json().catch(() => ({}));
  const activityId = placeholder?.id;

  /****************************
   * RAG QUERY
   ****************************/
  const ragRes = await fetch(RAG_QUERY_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      apikey: SUPABASE_ANON_KEY,
      "x-tenant-id": tenantId,
    },
    body: JSON.stringify({
      question: activity.text,
      source: "teams",
    }),
  });

  const rag = await ragRes.json().catch(() => null);
  if (!activityId || !rag) return new Response("ok");

  /****************************
   * PATCH FINAL MESSAGE
   ****************************/
  await fetch(
    `${serviceUrl}/v3/conversations/${encodeURIComponent(
      activity.conversation.id,
    )}/activities/${activityId}`,
    {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        type: "message",
        text: rag.answer ?? "No answer found.",
      }),
    },
  );

  return new Response("ok");
}

/********************************************************************************************
 * SERVER
 ********************************************************************************************/
serve(req => {
  const path = new URL(req.url).pathname;
  console.log("➡️ Request", req.method, path);
  if (path === "/teams") return handleTeams(req);
  return new Response("ok");
});
