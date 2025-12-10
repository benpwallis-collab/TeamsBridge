/********************************************************************************************
 * InnsynAI Teams Bridge (Shared, Multi-tenant, Slack-aligned security model)
 *
 * - Multi-tenant BotFramework JWT validation (audience = customer's Bot App ID)
 * - Resolves Bot App ID → Innsyn tenant (new)
 * - Resolves AAD tenant → Innsyn tenant (existing)
 * - Calls RAG via anon key + x-tenant-id
 * - Replies to Teams using per-tenant bot credentials
 *
 * No service-role keys. All lookups go through Lovable’s internal resolver.
 ********************************************************************************************/

import { serve } from "https://deno.land/std@0.224.0/http/server.ts";
import * as jose from "https://deno.land/x/jose@v5.4.0/index.ts";

/********************************************************************************************
 * ENV VARS
 ********************************************************************************************/
const {
  INTERNAL_LOOKUP_SECRET,
  TEAMS_TENANT_LOOKUP_URL, // Used for BOTH AAD tenant and bot_app_id lookup
  RAG_QUERY_URL,
  SUPABASE_ANON_KEY,
} = Deno.env.toObject();

console.log("🔧 TEAMS BRIDGE STARTUP");
console.log("  TEAMS_TENANT_LOOKUP_URL:", TEAMS_TENANT_LOOKUP_URL);
console.log("  RAG_QUERY_URL:", RAG_QUERY_URL);

if (
  !INTERNAL_LOOKUP_SECRET ||
  !TEAMS_TENANT_LOOKUP_URL ||
  !RAG_QUERY_URL ||
  !SUPABASE_ANON_KEY
) {
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
  const open = await fetch(OPENID_CONFIG_URL).then((r) => r.json());
  jwks = await fetch(open.jwks_uri).then((r) => r.json());
  console.log("✅ JWKS loaded");
  return jwks!;
}

/********************************************************************************************
 * RESOLVERS (Lovable-managed microservice)
 ********************************************************************************************/

// NEW → Resolve using Bot App ID
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
  return json; // { tenant_id, bot_app_id, bot_app_password }
}

// Existing → Resolve using AAD tenant ID
async function resolveInnsynTenantId(aadTenantId: string): Promise<string | null> {
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
 * MULTI-TENANT JWT VALIDATION
 *
 * Steps:
 * 1. Decode BF JWT WITHOUT verifying → extract “aud” (customer bot App ID)
 * 2. Lookup bot_app_id → tenant & bot secret
 * 3. Verify JWT using that bot’s App ID
 ********************************************************************************************/

async function verifyBotFrameworkJwt(authHeader: string | null) {
  if (!authHeader?.startsWith("Bearer ")) {
    throw new Error("Missing or invalid Authorization header");
  }

  const token = authHeader.slice("Bearer ".length);
  const decoded = jose.decodeJwt(token);
  const incomingBotAppId = decoded.aud;

  if (typeof incomingBotAppId !== "string") {
    throw new Error("Invalid or missing aud claim in BF JWT");
  }

  console.log("🔍 Incoming bot App ID (aud):", incomingBotAppId);

  // 1) Resolve which Innsyn tenant owns this bot
  const tenantInfo = await resolveByBotAppId(incomingBotAppId);
  if (!tenantInfo) throw new Error("Unknown bot App ID");

  // 2) Validate JWT against *this tenant’s* bot App ID
  const keyStore = jose.createLocalJWKSet(await getJwks());
  await jose.jwtVerify(token, keyStore, {
    issuer: "https://api.botframework.com",
    audience: incomingBotAppId,
  });

  console.log("✅ BF JWT verified for bot:", incomingBotAppId);
  return tenantInfo; // { tenant_id, bot_app_id, bot_app_password }
}

/********************************************************************************************
 * TYPES
 ********************************************************************************************/
interface TeamsActivity {
  type: string;
  id?: string;
  serviceUrl?: string;
  text?: string;
  replyToId?: string;
  conversation?: { id: string };
  channelData?: {
    tenant?: { id?: string }; // AAD tenant ID
  };
}

/********************************************************************************************
 * RAG QUERY
 ********************************************************************************************/
async function callRagQuery(tenantId: string, question: string) {
  console.log("🔍 Calling RAG for tenant:", tenantId, "question:", question);

  const res = await fetch(RAG_QUERY_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      apikey: SUPABASE_ANON_KEY,
      "x-tenant-id": tenantId,
    },
    body: JSON.stringify({ question, source: "teams" }),
  });

  if (!res.ok) {
    const text = await res.text();
    console.error("❌ RAG error:", text);
    throw new Error("RAG failed");
  }

  const json = await res.json();
  console.log("✅ RAG response received");
  return json;
}

/********************************************************************************************
 * REPLY TO TEAMS (USING PER-TENANT BOT CREDENTIALS)
 ********************************************************************************************/
async function getBotFrameworkToken(botAppId: string, botAppPassword: string): Promise<string> {
  console.log("🔍 BF TOKEN REQUEST for bot:", botAppId);

  const body = new URLSearchParams({
    grant_type: "client_credentials",
    client_id: botAppId,
    client_secret: botAppPassword,
    scope: "https://api.botframework.com/.default",
  });

  const res = await fetch(
    "https://login.microsoftonline.com/botframework.com/oauth2/v2.0/token",
    { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body }
  );

  if (!res.ok) {
    console.error("❌ BF TOKEN ERROR:", await res.text());
    throw new Error("botframework token error");
  }

  const json = await res.json();
  console.log("✅ BF TOKEN ACQUIRED");
  return json.access_token;
}

async function sendTeamsReply(activity: TeamsActivity, text: string, creds: any) {
  if (!activity.serviceUrl || !activity.conversation?.id) return;

  console.log("🔍 Sending reply using bot:", creds.bot_app_id);

  const token = await getBotFrameworkToken(creds.bot_app_id, creds.bot_app_password);
  const serviceUrl = activity.serviceUrl.replace(/\/+$/, "");
  const url = `${serviceUrl}/v3/conversations/${encodeURIComponent(activity.conversation.id)}/activities`;

  const payload = { type: "message", text, replyToId: activity.replyToId ?? activity.id };

  const res = await fetch(url, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });

  if (!res.ok) {
    console.error("❌ REPLY ERROR", await res.text());
  } else {
    console.log("✅ Reply sent");
  }
}

/********************************************************************************************
 * MAIN HANDLER
 ********************************************************************************************/
async function handleTeams(req: Request): Promise<Response> {
  if (req.method === "GET") return new Response("ok");
  if (req.method !== "POST") return new Response("method not allowed", { status: 405 });

  console.log("🔔 Incoming Teams POST");

  // 1) Multi-tenant JWT validation
  let creds;
  try {
    creds = await verifyBotFrameworkJwt(req.headers.get("Authorization"));
  } catch (err) {
    console.error("❌ JWT failed:", err);
    return new Response("unauthorized", { status: 401 });
  }

  // 2) Parse activity
  const activity = (await req.json()) as TeamsActivity;
  if (activity.type !== "message" || !activity.text) return new Response("ignored");

  const aadTenantId = activity.channelData?.tenant?.id;
  if (!aadTenantId) return new Response("bad request", { status: 400 });

  // 3) Resolve Innsyn tenant by AAD tenant
  const tenantId = await resolveInnsynTenantId(aadTenantId);
  if (!tenantId) {
    await sendTeamsReply(activity, "⚠️ InnsynAI is not configured for this Microsoft 365 tenant.", creds);
    return new Response("no tenant mapping");
  }

  // 4) Call RAG
  let rag;
  try {
    rag = await callRagQuery(tenantId, activity.text.trim());
  } catch {
    await sendTeamsReply(activity, "❌ Something went wrong.", creds);
    return new Response("rag error");
  }

  // 5) Reply
  await sendTeamsReply(activity, rag?.answer ?? "No answer.", creds);
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
