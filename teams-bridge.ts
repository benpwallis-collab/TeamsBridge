/********************************************************************************************
 * InnsynAI Teams Bridge (Shared, Multi-tenant, Slack-aligned security model)
 *
 * - Verifies BotFramework JWT
 * - Resolves AAD Tenant -> InnsynAI tenant via Lovable internal resolver
 * - Calls RAG via anon key + x-tenant-id
 * - Replies to Teams via BotFramework API
 ********************************************************************************************/

import { serve } from "https://deno.land/std@0.224.0/http/server.ts";
import * as jose from "https://deno.land/x/jose@v5.4.0/index.ts";

/********************************************************************************************
 * ENV VARS
 ********************************************************************************************/
const {
  TEAMS_APP_ID,
  TEAMS_APP_PASSWORD,
  INTERNAL_LOOKUP_SECRET,
  TEAMS_TENANT_LOOKUP_URL,
  RAG_QUERY_URL,
  SUPABASE_ANON_KEY,
} = Deno.env.toObject();

console.log("🔧 TEAMS BRIDGE STARTUP");
console.log("  TEAMS_APP_ID:", TEAMS_APP_ID);
console.log(
  "  TEAMS_APP_PASSWORD length:",
  TEAMS_APP_PASSWORD ? TEAMS_APP_PASSWORD.length : 0,
);
console.log("  TEAMS_TENANT_LOOKUP_URL:", TEAMS_TENANT_LOOKUP_URL);
console.log("  RAG_QUERY_URL:", RAG_QUERY_URL);

if (
  !TEAMS_APP_ID ||
  !TEAMS_APP_PASSWORD ||
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

async function verifyBotFrameworkJwt(authHeader: string | null) {
  if (!authHeader?.startsWith("Bearer ")) {
    throw new Error("Missing or invalid Authorization header");
  }

  const token = authHeader.slice("Bearer ".length);
  const keyStore = jose.createLocalJWKSet(await getJwks());

  console.log("🔍 Verifying incoming BF JWT with audience:", TEAMS_APP_ID);

  const result = await jose.jwtVerify(token, keyStore, {
    issuer: "https://api.botframework.com",
    audience: TEAMS_APP_ID,
  });

  console.log("✅ Incoming JWT verified");
  return result;
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
 * INTERNAL TENANT LOOKUP (Slack-aligned)
 ********************************************************************************************/
async function resolveInnsynTenantId(
  aadTenantId: string,
): Promise<string | null> {
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
 * RAG QUERY (Slack-aligned security model)
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
    body: JSON.stringify({
      question,
      source: "teams",
    }),
  });

  if (!res.ok) {
    const text = await res.text();
    console.error("❌ RAG error:", text);
    throw new Error("RAG failed");
  }

  const json = await res.json();
  console.log("✅ RAG response received (keys):", Object.keys(json));
  return json;
}

/********************************************************************************************
 * BOTFRAMEWORK REPLY
 ********************************************************************************************/
async function getBotFrameworkToken(): Promise<string> {
  console.log("🔍 BF TOKEN REQUEST");
  console.log("  TEAMS_APP_ID:", TEAMS_APP_ID);
  console.log(
    "  TEAMS_APP_PASSWORD length:",
    TEAMS_APP_PASSWORD ? TEAMS_APP_PASSWORD.length : 0,
  );

  const body = new URLSearchParams({
    grant_type: "client_credentials",
    client_id: TEAMS_APP_ID,
    client_secret: TEAMS_APP_PASSWORD,
    scope: "https://api.botframework.com/.default",
  });

  const res = await fetch(
    "https://login.microsoftonline.com/botframework.com/oauth2/v2.0/token",
    {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
      body,
    },
  );

  if (!res.ok) {
    const text = await res.text();
    console.error("❌ BF TOKEN ERROR");
    console.error("  Status:", res.status);
    console.error("  Body:", text);
    throw new Error("botframework token error");
  }

  const json = await res.json();
  const accessToken: string = json.access_token;

  console.log("✅ BF TOKEN ACQUIRED");
  console.log("  Access token first 40 chars:", accessToken.slice(0, 40));

  // Best-effort decode of token to inspect appid/aud
  try {
    const parts = accessToken.split(".");
    if (parts.length === 3) {
      const payloadB64 = parts[1]
        .replace(/-/g, "+")
        .replace(/_/g, "/")
        .padEnd(parts[1].length + (4 - (parts[1].length % 4)) % 4, "=");
      const payloadJson = JSON.parse(atob(payloadB64));
      console.log("🔍 DECODED ACCESS TOKEN PAYLOAD:", payloadJson);
    } else {
      console.error("⚠️ Unexpected access token format, cannot decode payload");
    }
  } catch (err) {
    console.error("⚠️ Failed to decode access token payload:", err);
  }

  return accessToken;
}

async function sendTeamsReply(activity: TeamsActivity, text: string) {
  if (!activity.serviceUrl || !activity.conversation?.id) {
    console.error("❌ Missing serviceUrl or conversation.id");
    console.error("Activity:", JSON.stringify(activity, null, 2));
    return;
  }

  console.log("🔍 REPLY PREP");
  console.log("  serviceUrl:", activity.serviceUrl);
  console.log("  conversationId:", activity.conversation.id);
  console.log("  replyToId:", activity.replyToId || activity.id);
  console.log("  text:", text);

  const token = await getBotFrameworkToken();

  const url =
    `${activity.serviceUrl}/v3/conversations/${encodeURIComponent(
      activity.conversation.id,
    )}/activities`;

  console.log("🔍 REPLY POST →", url);

  const payload = {
    type: "message",
    text,
    replyToId: activity.replyToId ?? activity.id,
  };

  const res = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(payload),
  });

  if (!res.ok) {
    const body = await res.text();
    console.error("❌ REPLY ERROR");
    console.error("  Status:", res.status);
    console.error("  Body:", body);
    return;
  }

  console.log("✅ REPLY SENT SUCCESSFULLY");
}

/********************************************************************************************
 * MAIN HANDLER
 ********************************************************************************************/
async function handleTeams(req: Request): Promise<Response> {
  if (req.method === "GET") {
    return new Response("ok", { status: 200 });
  }
  if (req.method !== "POST") {
    return new Response("method not allowed", { status: 405 });
  }

  console.log("🔔 Incoming Teams request:", req.method, req.url);

  // 1) Verify JWT
  try {
    await verifyBotFrameworkJwt(req.headers.get("Authorization"));
  } catch (err) {
    console.error("❌ JWT verification failed:", err);
    return new Response("unauthorized", { status: 401 });
  }

  const activity = (await req.json()) as TeamsActivity;

  console.log("🔍 Parsed activity:");
  console.log("  type:", activity.type);
  console.log("  id:", activity.id);
  console.log("  serviceUrl:", activity.serviceUrl);
  console.log("  conversationId:", activity.conversation?.id);
  console.log("  text:", activity.text);
  console.log("  channelData.tenant.id:", activity.channelData?.tenant?.id);

  if (activity.type !== "message" || !activity.text) {
    console.log("ℹ️ Non-message or empty-text activity, ignoring");
    return new Response("ignored", { status: 200 });
  }

  // 2) Extract AAD Tenant ID
  const aadTenantId = activity.channelData?.tenant?.id;
  if (!aadTenantId) {
    console.error("❌ No tenant id in activity");
    return new Response("bad request", { status: 400 });
  }

  // 3) Resolve InnsynAI tenant (Slack-aligned)
  const tenantId = await resolveInnsynTenantId(aadTenantId);
  if (!tenantId) {
    console.error("❌ No InnsynAI tenant mapping for AAD tenant:", aadTenantId);
    await sendTeamsReply(
      activity,
      "⚠️ InnsynAI is not configured for this Microsoft 365 tenant. Admin must click 'Add to Teams' in InnsynAI.",
    );
    return new Response("no tenant mapping", { status: 200 });
  }

  // 4) Call RAG
  let rag;
  try {
    rag = await callRagQuery(tenantId, activity.text.trim());
  } catch (err) {
    console.error("❌ RAG call failed:", err);
    await sendTeamsReply(activity, "❌ Something went wrong.");
    return new Response("rag error", { status: 200 });
  }

  // 5) Reply to Teams
  const answer = rag?.answer ?? "No answer.";
  console.log("🔍 RAG answer:", answer);
  await sendTeamsReply(activity, answer);

  return new Response("ok", { status: 200 });
}

/********************************************************************************************
 * SERVER
 ********************************************************************************************/
serve((req) => {
  const url = new URL(req.url);

  if (url.pathname === "/health") {
    return new Response("ok", { status: 200 });
  }
  if (url.pathname === "/teams") {
    return handleTeams(req);
  }

  return new Response("not found", { status: 404 });
});
