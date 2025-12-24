/********************************************************************************************
 * InnsynAI Teams Bridge – MULTI-TENANT / STORE READY
 * - Logs AAD tenant ID
 * - Auto-provisions tenant on first message
 * - Preserves existing bot + RAG behavior
 ********************************************************************************************/

import { serve } from "https://deno.land/std@0.224.0/http/server.ts";
import * as jose from "https://deno.land/x/jose@v5.4.0/index.ts";

/********************************************************************************************
 * ENV VARS
 ********************************************************************************************/
const {
  INTERNAL_LOOKUP_SECRET,
  TEAMS_TENANT_LOOKUP_URL,
  RAG_QUERY_URL,
  SUPABASE_ANON_KEY,
} = Deno.env.toObject();

if (!INTERNAL_LOOKUP_SECRET || !TEAMS_TENANT_LOOKUP_URL || !RAG_QUERY_URL || !SUPABASE_ANON_KEY) {
  console.error("❌ Missing env vars");
  Deno.exit(1);
}

/********************************************************************************************
 * BOTFRAMEWORK OPENID CONFIG
 ********************************************************************************************/
const OPENID_CONFIG_URL =
  "https://login.botframework.com/v1/.well-known/openidconfiguration";

let jwks: jose.JSONWebKeySet | null = null;

async function getJwks(): Promise<jose.JSONWebKeySet> {
  if (jwks) return jwks;
  const meta = await fetch(OPENID_CONFIG_URL).then(r => r.json());
  jwks = await fetch(meta.jwks_uri).then(r => r.json());
  return jwks!;
}

/********************************************************************************************
 * TENANT LOOKUP / AUTO-PROVISION
 ********************************************************************************************/
async function resolveOrCreateTenant(aadTenantId: string) {
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
    console.error("❌ Tenant lookup/provision failed", await res.text());
    return null;
  }

  const json = await res.json();
  return json.tenant_id ?? null;
}

/********************************************************************************************
 * BOT LOOKUP BY APP ID (JWT VALIDATION)
 ********************************************************************************************/
async function resolveByBotAppId(botAppId: string) {
  const res = await fetch(TEAMS_TENANT_LOOKUP_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      apikey: SUPABASE_ANON_KEY,
      "x-internal-token": INTERNAL_LOOKUP_SECRET,
    },
    body: JSON.stringify({ bot_app_id: botAppId }),
  });
  return res.ok ? res.json() : null;
}

/********************************************************************************************
 * JWT VERIFICATION
 ********************************************************************************************/
async function verifyJwt(authHeader: string) {
  const token = authHeader.slice(7);
  const decoded = jose.decodeJwt(token);
  const botAppId = decoded.aud as string;

  const tenantInfo = await resolveByBotAppId(botAppId);
  if (!tenantInfo) throw new Error("Unknown bot");

  const keyStore = jose.createLocalJWKSet(await getJwks());
  await jose.jwtVerify(token, keyStore, {
    issuer: "https://api.botframework.com",
    audience: botAppId,
  });

  return tenantInfo;
}

/********************************************************************************************
 * BOT TOKEN
 ********************************************************************************************/
async function getBotAccessToken(aadTenantId: string, creds: any) {
  const res = await fetch(
    `https://login.microsoftonline.com/${aadTenantId}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type: "client_credentials",
        client_id: creds.bot_app_id,
        client_secret: creds.bot_app_password,
        scope: "https://api.botframework.com/.default",
      }),
    },
  );

  const json = await res.json();
  if (!json.access_token) {
    console.error("❌ Failed to get bot token", json);
    throw new Error("Bot token failure");
  }

  return json.access_token;
}

/********************************************************************************************
 * MAIN HANDLER
 ********************************************************************************************/
async function handleTeams(req: Request): Promise<Response> {
  console.log("🔥 Incoming Teams request");

  if (req.method !== "POST") return new Response("ok");

  const activity = await req.json();

  const auth = req.headers.get("Authorization");
  if (!auth) return new Response("ok");

  let creds;
  try {
    creds = await verifyJwt(auth);
  } catch (err) {
    console.error("❌ JWT verification failed", err);
    return new Response("unauthorized", { status: 401 });
  }

  const aadTenantId =
    activity.channelData?.tenant?.id || activity.conversation?.tenantId;

  console.log("🏷️ AAD TENANT ID:", aadTenantId);

  if (!aadTenantId) {
    console.warn("⚠️ Missing AAD tenant id");
    return new Response("ok");
  }

  const tenantId = await resolveOrCreateTenant(aadTenantId);

  if (!tenantId) {
    console.error("❌ Unable to resolve or create tenant for", aadTenantId);
    return new Response("ok");
  }

  console.log("🔐 InnsynAI tenant resolved:", tenantId);

  if (!activity.text) return new Response("ok");

  /****************************
   * SEND PLACEHOLDER
   ****************************/
  const accessToken = await getBotAccessToken(aadTenantId, creds);

  const placeholderRes = await fetch(
    `${activity.serviceUrl}/v3/conversations/${encodeURIComponent(
      activity.conversation.id,
    )}/activities`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        type: "message",
        text: "⏳ Working on it…",
        replyToId: activity.replyToId ?? activity.id,
      }),
    },
  );

  const placeholderJson = await placeholderRes.json();
  const placeholderActivityId = placeholderJson.id;

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
      question: activity.text.trim(),
      source: "teams",
    }),
  });

  if (!ragRes.ok) {
    console.error("❌ RAG failed", await ragRes.text());
    return new Response("ok");
  }

  const rag = await ragRes.json();

  /****************************
   * UPDATE MESSAGE
   ****************************/
  const patchToken = await getBotAccessToken(aadTenantId, creds);

  await fetch(
    `${activity.serviceUrl}/v3/conversations/${encodeURIComponent(
      activity.conversation.id,
    )}/activities/${placeholderActivityId}`,
    {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${patchToken}`,
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
  if (path === "/teams") return handleTeams(req);
  return new Response("ok");
});
