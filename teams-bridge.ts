/********************************************************************************************
 * InnsynAI Teams Bridge – Fully Patched Version
 * 
 * Includes:
 *  - Full activity logging
 *  - Valid BotFramework JWT validation
 *  - Tenant resolution by bot_app_id + AAD tenant
 *  - RAG query forwarding
 *  - Correct Teams reply handling
 *  - FIX: serviceUrl normalization (removes invalid tenant-id suffix)
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

console.log("🔧 TEAMS BRIDGE STARTUP");
console.log("  TEAMS_TENANT_LOOKUP_URL:", TEAMS_TENANT_LOOKUP_URL);
console.log("  RAG_QUERY_URL:", RAG_QUERY_URL);

if (!INTERNAL_LOOKUP_SECRET || !TEAMS_TENANT_LOOKUP_URL ||
    !RAG_QUERY_URL || !SUPABASE_ANON_KEY) {
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
  console.log("🔍 Fetching BotFramework OpenID configuration");
  const meta = await fetch(OPENID_CONFIG_URL).then(r => r.json());
  jwks = await fetch(meta.jwks_uri).then(r => r.json());
  console.log("✅ JWKS loaded");
  return jwks!;
}

/********************************************************************************************
 * RESOLVERS
 ********************************************************************************************/
async function resolveByBotAppId(botAppId: string) {
  console.log("🔍 Resolving Innsyn tenant for bot_app_id:", botAppId);

  const res = await fetch(TEAMS_TENANT_LOOKUP_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      apikey: SUPABASE_ANON_KEY,
      "x-internal-token": INTERNAL_LOOKUP_SECRET,
    },
    body: JSON.stringify({ bot_app_id: botAppId }),
  });

  if (!res.ok) {
    console.error("❌ BotAppId resolver failed:", await res.text());
    return null;
  }

  const json = await res.json();
  console.log("✅ BotAppId resolver result:", json);
  return json;
}

async function resolveInnsynTenantId(aadTenantId: string) {
  console.log("🔍 Resolving InnsynAI tenant for AAD tenant:", aadTenantId);

  const res = await fetch(TEAMS_TENANT_LOOKUP_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      apikey: SUPABASE_ANON_KEY,
      "x-internal-token": INTERNAL_LOOKUP_SECRET,
    },
    body: JSON.stringify({ teams_tenant_id: aadTenantId }),
  });

  if (!res.ok) {
    console.error("❌ Tenant resolver failed:", await res.text());
    return null;
  }

  const json = await res.json();
  console.log("✅ Tenant resolver result:", json);
  return json.tenant_id ?? null;
}

/********************************************************************************************
 * MULTI-TENANT BF JWT VALIDATION
 ********************************************************************************************/
async function verifyBotFrameworkJwt(authHeader: string | null) {
  if (!authHeader?.startsWith("Bearer "))
    throw new Error("Missing auth header");

  const token = authHeader.slice("Bearer ".length);
  const decoded = jose.decodeJwt(token);

  const botAppId = decoded.aud;
  if (!botAppId || typeof botAppId !== "string")
    throw new Error("Missing or invalid aud");

  console.log("🔍 Incoming bot App ID (aud):", botAppId);

  const tenantInfo = await resolveByBotAppId(botAppId);
  if (!tenantInfo) throw new Error("Unknown bot App ID");

  const keyStore = jose.createLocalJWKSet(await getJwks());
  await jose.jwtVerify(token, keyStore, {
    issuer: "https://api.botframework.com",
    audience: botAppId,
  });

  console.log("✅ BF JWT verified for bot:", botAppId);
  return tenantInfo;
}

/********************************************************************************************
 * TYPES
 ********************************************************************************************/
interface TeamsActivity {
  type: string;
  id?: string;
  text?: string;
  serviceUrl?: string;
  replyToId?: string;
  conversation?: { id: string };
  channelData?: { tenant?: { id?: string } };
  [key: string]: unknown;
}

/********************************************************************************************
 * RAG QUERY
 ********************************************************************************************/
async function callRagQuery(tenantId: string, q: string) {
  console.log("🔍 Calling RAG for tenant:", tenantId, "question:", q);

  const res = await fetch(RAG_QUERY_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      apikey: SUPABASE_ANON_KEY,
      "x-tenant-id": tenantId,
    },
    body: JSON.stringify({ question: q, source: "teams" }),
  });

  if (!res.ok) {
    console.error("❌ RAG error:", await res.text());
    throw new Error("rag");
  }

  console.log("✅ RAG response received");
  return res.json();
}

/********************************************************************************************
 * FIX: Correct serviceUrl normalization
 ********************************************************************************************/
function normalizeServiceUrl(raw: string): string {
  if (!raw) return "";

  let url = raw.trim();
  url = url.replace(/\/+$/, "");      // strip trailing slash
  url = url.split("?")[0];            // remove any query params

  const emeaPrefix = "https://smba.trafficmanager.net/emea";

  if (url.startsWith(emeaPrefix + "/")) {
    const suffix = url.slice(emeaPrefix.length + 1);

    // if last path segment is a GUID => it's invalid Teams tenantId segment
    if (/^[0-9a-fA-F-]{36}$/.test(suffix)) {
      console.log("⚠️ Removing invalid tenant segment from serviceUrl:", suffix);
      url = emeaPrefix;
    }
  }

  return url;
}

/********************************************************************************************
 * BOTFRAMEWORK TOKEN
 ********************************************************************************************/
async function getBotFrameworkToken(id: string, secret: string) {
  console.log("🔍 BF TOKEN REQUEST for bot:", id);
  const body = new URLSearchParams({
    grant_type: "client_credentials",
    client_id: id,
    client_secret: secret,
    scope: "https://api.botframework.com/.default",
  });

  const res = await fetch(
    "https://login.microsoftonline.com/botframework.com/oauth2/v2.0/token",
    { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body }
  );

  if (!res.ok) {
    console.error("❌ BF TOKEN ERROR:", await res.text());
    throw new Error("bf-token");
  }

  console.log("✅ BF TOKEN ACQUIRED");
  const json = await res.json();
  return json.access_token;
}

/********************************************************************************************
 * SEND TEAMS REPLY (PATChed)
 ********************************************************************************************/
async function sendTeamsReply(activity: TeamsActivity, text: string, creds: any) {
  console.log("📝 FULL ACTIVITY:", JSON.stringify(activity, null, 2));

  if (!activity.serviceUrl) {
    console.error("❌ No serviceUrl");
    return;
  }

  if (!activity.conversation?.id) {
    console.error("❌ No conversation.id");
    return;
  }

  // FIXED serviceUrl
  const serviceUrl = normalizeServiceUrl(activity.serviceUrl);
  console.log("🔍 Normalized serviceUrl:", serviceUrl);

  // Token for reply
  const token = await getBotFrameworkToken(creds.bot_app_id, creds.bot_app_password);

  const url = `${serviceUrl}/v3/conversations/${encodeURIComponent(activity.conversation.id)}/activities`;
  console.log("🔍 Teams reply URL:", url);

  const payload = {
    type: "message",
    text,
    replyToId: activity.replyToId ?? activity.id,
  };

  console.log("📤 Reply payload:", payload);

  const res = await fetch(url, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });

  const body = await res.text();
  if (!res.ok) {
    console.error("❌ REPLY ERROR", res.status, body);
  } else {
    console.log("✅ Reply sent:", body);
  }
}

/********************************************************************************************
 * MAIN HANDLER
 ********************************************************************************************/
async function handleTeams(req: Request): Promise<Response> {
  if (req.method === "GET") return new Response("ok");
  if (req.method !== "POST") return new Response("method not allowed", { status: 405 });

  console.log("🔔 Incoming Teams POST");

  // 1) JWT validation
  let creds;
  try {
    creds = await verifyBotFrameworkJwt(req.headers.get("Authorization"));
  } catch (e) {
    console.error("❌ JWT failed:", e);
    return new Response("unauthorized", { status: 401 });
  }

  // 2) Parse body
  let activity: TeamsActivity;
  try {
    activity = await req.json();
  } catch (e) {
    console.error("❌ Bad JSON:", e);
    return new Response("bad request", { status: 400 });
  }

  console.log("📝 FULL RAW ACTIVITY:", JSON.stringify(activity, null, 2));

  if (activity.type !== "message" || !activity.text) {
    console.log("ℹ️ Ignored non-message activity");
    return new Response("ignored");
  }

  const aadTenantId = activity.channelData?.tenant?.id;
  if (!aadTenantId) return new Response("bad request", { status: 400 });

  // 3) Resolve Innsyn tenant
  const tenantId = await resolveInnsynTenantId(aadTenantId);
  if (!tenantId) {
    await sendTeamsReply(activity, "⚠️ InnsynAI is not configured for your Microsoft 365 tenant.", creds);
    return new Response("no tenant");
  }

  // 4) RAG
  let rag;
  try {
    rag = await callRagQuery(tenantId, activity.text.trim());
  } catch {
    await sendTeamsReply(activity, "❌ Something went wrong.", creds);
    return new Response("rag error");
  }

  // 5) Reply
  await sendTeamsReply(activity, rag?.answer ?? "No answer found.", creds);
  return new Response("ok");
}

/********************************************************************************************
 * SERVER
 ********************************************************************************************/
serve((req) => {
  const url = new URL(req.url);
  if (url.pathname === "/health") return new Response("ok");
  if (url.pathname === "/teams") return handleTeams(req);
  return new Response("not found", { status: 404 });
});
