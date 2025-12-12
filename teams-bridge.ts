/********************************************************************************************
 * InnsynAI Teams Bridge – FINAL CLEAN VERSION
 *
 * ✔ Adaptive Card ONLY (no text duplication)
 * ✔ Sources rendered
 * ✔ No repeated question
 * ✔ 👍 / 👎 feedback wired to analytics
 * ✔ Existing auth / tenant / BF logic preserved
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
  const meta = await fetch(OPENID_CONFIG_URL).then((r) => r.json());
  jwks = await fetch(meta.jwks_uri).then((r) => r.json());
  return jwks!;
}

/********************************************************************************************
 * RESOLVERS
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
 * JWT VALIDATION
 ********************************************************************************************/
async function verifyBotFrameworkJwt(authHeader: string | null) {
  if (!authHeader?.startsWith("Bearer ")) throw new Error("Missing auth");

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
  type: string;
  id?: string;
  text?: string;
  value?: any;
  serviceUrl?: string;
  replyToId?: string;
  conversation?: { id: string; tenantId?: string };
  channelData?: { tenant?: { id?: string } };
}

type RagResponse = {
  answer?: string;
  confidence?: number;
  reviewed?: boolean;
  sources?: { title?: string; url?: string }[];
};

/********************************************************************************************
 * RAG
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

  const json = await res.json();
  console.log("🧾 RAG FULL RESPONSE:", JSON.stringify(json, null, 2));
  return json;
}

/********************************************************************************************
 * ADAPTIVE CARD BUILDER
 ********************************************************************************************/
function buildAdaptiveCard(rag: RagResponse, tenantId: string) {
  const facts: any[] = [];
  if (typeof rag.confidence === "number") {
    facts.push({ title: "Confidence", value: `${Math.round(rag.confidence * 100)}%` });
  }
  if (rag.reviewed) {
    facts.push({ title: "Reviewed", value: "Yes" });
  }

  const sourceBlocks =
    rag.sources?.length
      ? [
          { type: "TextBlock", text: "Sources", weight: "Bolder", spacing: "Medium" },
          ...rag.sources.slice(0, 8).map((s) => ({
            type: "TextBlock",
            text: `• [${s.title ?? "Source"}](${s.url})`,
            wrap: true,
            spacing: "None",
          })),
        ]
      : [];

  return {
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.4",
    body: [
      { type: "TextBlock", text: rag.answer ?? "No answer found.", wrap: true },
      ...(facts.length ? [{ type: "FactSet", facts }] : []),
      ...(sourceBlocks.length ? [{ type: "Container", items: sourceBlocks }] : []),
    ],
    actions: [
      {
        type: "Action.Submit",
        title: "👍 Helpful",
        data: { action: "feedback", rating: "up", tenantId },
      },
      {
        type: "Action.Submit",
        title: "👎 Not helpful",
        data: { action: "feedback", rating: "down", tenantId },
      },
    ],
  };
}

/********************************************************************************************
 * SEND REPLY
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

  const replyUrl =
    `${activity.serviceUrl}/v3/conversations/${encodeURIComponent(
      activity.conversation!.id,
    )}/activities`;

  await fetch(replyUrl, {
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
  });
}

/********************************************************************************************
 * MAIN HANDLER
 ********************************************************************************************/
async function handleTeams(req: Request): Promise<Response> {
  if (req.method !== "POST") return new Response("ok");

  const creds = await verifyBotFrameworkJwt(req.headers.get("Authorization"));
  const activity = (await req.json()) as TeamsActivity;

  const aadTenantId =
    activity.channelData?.tenant?.id || activity.conversation?.tenantId;
  if (!aadTenantId) return new Response("bad request", { status: 400 });

  // 👍 / 👎 Feedback
  if (activity.value?.action === "feedback") {
    await fetch(`${RAG_QUERY_URL.replace("/rag-query", "/feedback")}`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        apikey: SUPABASE_ANON_KEY,
        "x-tenant-id": activity.value.tenantId,
      },
      body: JSON.stringify({
        source: "teams",
        rating: activity.value.rating,
      }),
    });
    return new Response("feedback ok");
  }

  if (!activity.text) return new Response("ignored");

  const tenantId = await resolveInnsynTenantId(aadTenantId);
  if (!tenantId) return new Response("no tenant");

  const rag = await callRagQuery(tenantId, activity.text.trim());
  const card = buildAdaptiveCard(rag, tenantId);

  await sendTeamsReply(activity, card, creds, aadTenantId);
  return new Response("ok");
}

/********************************************************************************************
 * SERVER
 ********************************************************************************************/
serve((req) => {
  const url = new URL(req.url);
  if (url.pathname === "/teams") return handleTeams(req);
  return new Response("ok");
});
