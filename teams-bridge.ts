/********************************************************************************************
 * InnsynAI Teams Bridge – CORRECTED (Teams activity-safe, JWT-guarded)
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
  SUPABASE_URL,
} = Deno.env.toObject();

if (
  !INTERNAL_LOOKUP_SECRET ||
  !TEAMS_TENANT_LOOKUP_URL ||
  !RAG_QUERY_URL ||
  !SUPABASE_ANON_KEY ||
  !SUPABASE_URL
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
  const meta = await fetch(OPENID_CONFIG_URL).then(r => r.json());
  jwks = await fetch(meta.jwks_uri).then(r => r.json());
  return jwks!;
}

/********************************************************************************************
 * TENANT RESOLUTION
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

async function resolveInnsynTenantId(aadTenantId: string) {
  const res = await fetch(TEAMS_TENANT_LOOKUP_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      apikey: SUPABASE_ANON_KEY,
      "x-internal-token": INTERNAL_LOOKUP_SECRET,
    },
    body: JSON.stringify({ teams_tenant_id: aadTenantId }),
  });
  if (!res.ok) return null;
  const json = await res.json();
  return json.tenant_id ?? null;
}

/********************************************************************************************
 * JWT VALIDATION (GUARDED)
 ********************************************************************************************/
async function verifyBotFrameworkJwt(authHeader: string) {
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
 * TYPES
 ********************************************************************************************/
interface TeamsActivity {
  type?: string;
  id?: string;
  text?: string;
  from?: { id?: string };
  serviceUrl?: string;
  replyToId?: string;
  conversation?: { id: string; tenantId?: string };
  channelData?: { tenant?: { id?: string } };
  value?: any;
}

type RagResponse = {
  answer?: string;
  qa_log_id?: string;
};

/********************************************************************************************
 * RAG QUERY
 ********************************************************************************************/
async function callRagQuery(tenantId: string, q: string): Promise<RagResponse> {
  const res = await fetch(RAG_QUERY_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      apikey: SUPABASE_ANON_KEY,
      "x-tenant-id": tenantId,
    },
    body: JSON.stringify({ question: q, source: "teams" }),
  });

  if (!res.ok) throw new Error("RAG failed");
  return res.json();
}

/********************************************************************************************
 * SEND TEAMS REPLY
 ********************************************************************************************/
async function sendTeamsReply(
  activity: TeamsActivity,
  card: any,
  creds: any,
  aadTenantId: string,
) {
  const tokenRes = await fetch(
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

  const { access_token } = await tokenRes.json();

  await fetch(
    `${activity.serviceUrl}/v3/conversations/${encodeURIComponent(
      activity.conversation!.id,
    )}/activities`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${access_token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        type: "message",
        attachments: [
          {
            contentType: "application/vnd.microsoft.card.adaptive",
            content: card,
          },
        ],
        replyToId: activity.replyToId ?? activity.id,
      }),
    },
  );
}

/********************************************************************************************
 * MAIN HANDLER
 ********************************************************************************************/
async function handleTeams(req: Request): Promise<Response> {
  console.log("🔥 TEAMS HIT", req.method, req.url);

  if (req.method !== "POST") return new Response("ok");

  const activity = (await req.json()) as TeamsActivity;

  console.log(
    "📨 Activity:",
    activity.type,
    "| auth:",
    Boolean(req.headers.get("Authorization")),
  );

  // Ignore non-message noise
  if (!activity.text && !activity.value) {
    return new Response("ok");
  }

  // JWT guard
  const auth = req.headers.get("Authorization");
  if (!auth) {
    console.warn("⚠️ No auth header – ignoring activity");
    return new Response("ok");
  }

  let creds;
  try {
    creds = await verifyBotFrameworkJwt(auth);
    console.log("✅ JWT verified");
  } catch (err) {
    console.error("❌ JWT verification failed", err);
    return new Response("unauthorized", { status: 401 });
  }

  const aadTenantId =
    activity.channelData?.tenant?.id || activity.conversation?.tenantId;
  if (!aadTenantId) return new Response("bad request", { status: 400 });

  const tenantId = await resolveInnsynTenantId(aadTenantId);
  if (!tenantId) return new Response("no tenant");

  const rag = await callRagQuery(tenantId, activity.text.trim());

  await sendTeamsReply(
    activity,
    {
      type: "AdaptiveCard",
      version: "1.4",
      body: [{ type: "TextBlock", text: rag.answer ?? "No answer found.", wrap: true }],
    },
    creds,
    aadTenantId,
  );

  return new Response("ok");
}

/********************************************************************************************
 * SERVER
 ********************************************************************************************/
serve(req => {
  console.log("🔥 RAW REQUEST", req.method, new URL(req.url).pathname);

  if (new URL(req.url).pathname === "/teams") {
    return handleTeams(req);
  }

  return new Response("ok");
});
