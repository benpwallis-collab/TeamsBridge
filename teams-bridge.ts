/********************************************************************************************
 * InnsynAI Teams Bridge – Tenant-Aware, Fully Patched Version
 *
 * - Multi-tenant BotFramework JWT validation (aud = customer's Bot App ID)
 * - Resolve bot_app_id → Innsyn tenant (Lovable resolver)
 * - Resolve AAD tenant → Innsyn tenant (Lovable resolver)
 * - Call RAG via anon key + x-tenant-id
 * - Reply to Teams using per-tenant bot credentials
 *
 * Key fixes:
 *  - Correct serviceUrl normalization (remove invalid tenant-id suffix)
 *  - Use AAD tenant-specific token endpoint (not botframework.com) for BF access token
 *  - Use official scope: https://api.botframework.com/.default
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
 * BOTFRAMEWORK OPENID CONFIG FOR INBOUND JWT
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
 * RESOLVERS (Lovable microservice)
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
  return json; // { tenant_id, bot_app_id, bot_app_password }
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
 * MULTI-TENANT BF JWT VALIDATION (INBOUND)
 ********************************************************************************************/
async function verifyBotFrameworkJwt(authHeader: string | null) {
  if (!authHeader?.startsWith("Bearer "))
    throw new Error("Missing auth header");

  const token = authHeader.slice("Bearer ".length);
  const decoded = jose.decodeJwt(token);

  const botAppId = decoded.aud;
  if (!botAppId || typeof botAppId !== "string")
    throw new Error("Missing or invalid aud in BF JWT");

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
  conversation?: {
    id: string;
    tenantId?: string;
    conversationType?: string;
  };
  channelData?: {
    tenant?: { id?: string };
  };
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
 * serviceUrl NORMALIZATION
 ********************************************************************************************/
function normalizeServiceUrl(raw: string): string {
  if (!raw) return "";

  let url = raw.trim();
  url = url.replace(/\/+$/, ""); // strip trailing slash
  url = url.split("?")[0];       // remove params

  const emeaPrefix = "https://smba.trafficmanager.net/emea";

  // Some tenants append tenantId: https://smba.trafficmanager.net/emea/<tenant-id>
  if (url.startsWith(emeaPrefix + "/")) {
    const suffix = url.slice(emeaPrefix.length + 1);
    if (/^[0-9a-fA-F-]{36}$/.test(suffix)) {
      console.log("⚠️ Removing invalid tenant segment from serviceUrl:", suffix);
      url = emeaPrefix;
    }
  }

  return url;
}

/********************************************************************************************
 * BOTFRAMEWORK TOKEN (OUTBOUND) – TENANT-SPECIFIC
 *
 * Updated per MS guidance: use your own AAD tenant, not botframework.com,
 * with scope = https://api.botframework.com/.default
 ********************************************************************************************/
async function getBotFrameworkToken(
  botAppId: string,
  botAppPassword: string,
  aadTenantId: string,
): Promise<string> {
  const tokenUrl =
    `https://login.microsoftonline.com/${aadTenantId}/oauth2/v2.0/token`;

  console.log("🔍 BF TOKEN REQUEST for bot:", botAppId);
  console.log("🔍 Token URL:", tokenUrl);

  const body = new URLSearchParams({
    grant_type: "client_credentials",
    client_id: botAppId,
    client_secret: botAppPassword,
    scope: "https://api.botframework.com/.default",
  });

  const res = await fetch(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });

  const text = await res.text();
  if (!res.ok) {
    console.error("❌ BF TOKEN ERROR", res.status, text);
    throw new Error("bf-token-acquisition-failed");
  }

  let json: any;
  try {
    json = JSON.parse(text);
  } catch {
    console.error("❌ Failed to parse BF token JSON:", text);
    throw new Error("bf-token-parse-failed");
  }

  const accessToken = json.access_token;
  if (!accessToken) {
    console.error("❌ No access_token in BF token response:", json);
    throw new Error("bf-token-missing-access-token");
  }

  // Log key claims (no secret)
  try {
    const decoded = jose.decodeJwt(accessToken);
    console.log("🔍 BF ACCESS TOKEN PAYLOAD (truncated):", {
      aud: decoded.aud,
      iss: decoded.iss,
      appid: (decoded as any).appid,
      azp: (decoded as any).azp,
      tid: (decoded as any).tid,
      exp: (decoded as any).exp,
    });
  } catch (e) {
    console.warn("⚠️ Could not decode BF access token:", e);
  }

  console.log("✅ BF TOKEN ACQUIRED");
  return accessToken;
}

/********************************************************************************************
 * SEND TEAMS REPLY
 ********************************************************************************************/
async function sendTeamsReply(
  activity: TeamsActivity,
  text: string,
  creds: any,
  aadTenantId: string,
) {
  console.log("📝 FULL ACTIVITY:", JSON.stringify(activity, null, 2));

  if (!activity.serviceUrl) {
    console.error("❌ No serviceUrl");
    return;
  }

  if (!activity.conversation?.id) {
    console.error("❌ No conversation.id");
    return;
  }

  const serviceUrl = normalizeServiceUrl(activity.serviceUrl);
  console.log("🔍 Normalized serviceUrl:", serviceUrl);

  const bfToken = await getBotFrameworkToken(
    creds.bot_app_id,
    creds.bot_app_password,
    aadTenantId,
  );

  const replyUrl =
    `${serviceUrl}/v3/conversations/${encodeURIComponent(activity.conversation.id)}/activities`;

  console.log("🔍 Teams reply URL:", replyUrl);

  const payload = {
    type: "message",
    text,
    replyToId: activity.replyToId ?? activity.id,
  };

  console.log("📤 Reply payload:", payload);

  const res = await fetch(replyUrl, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${bfToken}`,
      "Content-Type": "application/json",
    },
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
  if (req.method !== "POST") {
    return new Response("method not allowed", { status: 405 });
  }

  console.log("🔔 Incoming Teams POST");

  // 1) Inbound JWT validation
  let creds;
  try {
    creds = await verifyBotFrameworkJwt(req.headers.get("Authorization"));
  } catch (e) {
    console.error("❌ JWT failed:", e);
    return new Response("unauthorized", { status: 401 });
  }

  // 2) Parse activity
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

  const aadTenantId =
    activity.channelData?.tenant?.id ||
    activity.conversation?.tenantId;

  if (!aadTenantId) {
    console.error("❌ Missing AAD tenant ID in activity");
    return new Response("bad request", { status: 400 });
  }

  // 3) Resolve Innsyn tenant
  const tenantId = await resolveInnsynTenantId(aadTenantId);
  if (!tenantId) {
    await sendTeamsReply(
      activity,
      "⚠️ InnsynAI is not configured for your Microsoft 365 tenant.",
      creds,
      aadTenantId,
    );
    return new Response("no tenant");
  }

  // 4) Call RAG
  let rag;
  try {
    rag = await callRagQuery(tenantId, activity.text.trim());
  } catch (e) {
    console.error("❌ RAG failed:", e);
    await sendTeamsReply(
      activity,
      "❌ Something went wrong.",
      creds,
      aadTenantId,
    );
    return new Response("rag error");
  }

  // 5) Reply with answer
  await sendTeamsReply(
    activity,
    rag?.answer ?? "No answer found.",
    creds,
    aadTenantId,
  );

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
