/********************************************************************************************
 * InnsynAI Teams Bridge – FINAL (with Human Answer Capture)
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
  from?: { id?: string };
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
 * HUMAN ANSWER CAPTURE (FIRE & FORGET)
 ********************************************************************************************/
function fireHumanAnswerCapture(
  tenantId: string,
  aadTenantId: string,
  activity: TeamsActivity,
) {
  try {
    if (!activity.text || !activity.conversation?.id || !activity.id) return;

    fetch(`${SUPABASE_URL}/functions/v1/capture-human-answers`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        apikey: SUPABASE_ANON_KEY,
      },
      body: JSON.stringify({
        tenant_id: tenantId,
        source_type: "teams",
        teams_tenant_id: aadTenantId,
        thread_messages: [
          {
            user_id: activity.from?.id ?? "unknown",
            text: activity.text,
            timestamp: new Date().toISOString(),
            is_bot: false,
          },
        ],
        source_reference: {
          conversation_id: activity.conversation.id,
          activity_id: activity.id,
        },
      }),
    });
  } catch {
    // HARD NO-OP: must never affect Teams flow
  }
}

/********************************************************************************************
 * ADAPTIVE CARD BUILDER
 ********************************************************************************************/
function buildAdaptiveCard(rag: RagResponse, tenantId: string, qaLogId?: string) {
  const facts: any[] = [];
  if (typeof rag.confidence === "number") {
    facts.push({ title: "Confidence", value: `${Math.round(rag.confidence * 100)}%` });
  }
  if (rag.reviewed) {
    facts.push({ title: "Reviewed", value: "Yes" });
  }

  const sources =
    rag.sources?.length
      ? [
          { type: "TextBlock", text: "Sources", weight: "Bolder", spacing: "Medium" },
          ...rag.sources.slice(0, 8).map(s => ({
            type: "TextBlock",
            text: `• [${s.title ?? "Source"}](${s.url})`,
            wrap: true,
            spacing: "None",
          })),
        ]
      : [];

  const feedback =
    qaLogId
      ? [
          {
            type: "Action.Submit",
            title: "👍 Helpful",
            data: { action: "feedback", feedback: "up", tenant_id: tenantId, qa_log_id: qaLogId },
          },
          {
            type: "Action.Submit",
            title: "👎 Not helpful",
            data: { action: "feedback", feedback: "down", tenant_id: tenantId, qa_log_id: qaLogId },
          },
        ]
      : [];

  return {
    type: "AdaptiveCard",
    version: "1.4",
    body: [
      { type: "TextBlock", text: rag.answer ?? "No answer found.", wrap: true },
      ...(facts.length ? [{ type: "FactSet", facts }] : []),
      ...(sources.length ? [{ type: "Container", items: sources }] : []),
    ],
    actions: feedback,
  };
}

/********************************************************************************************
 * SEND TEAMS REPLY
 ********************************************************************************************/
async function sendTeamsReply(activity: TeamsActivity, card: any, creds: any, aadTenantId: string) {
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
  if (req.method !== "POST") return new Response("ok");

  const creds = await verifyBotFrameworkJwt(req.headers.get("Authorization"));
  const activity = (await req.json()) as TeamsActivity;

  const aadTenantId =
    activity.channelData?.tenant?.id || activity.conversation?.tenantId;
  if (!aadTenantId) return new Response("bad request", { status: 400 });

  // Feedback
  if (activity.value?.action === "feedback") {
    await fetch(`${RAG_QUERY_URL.replace("/rag-query", "/feedback")}`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        apikey: SUPABASE_ANON_KEY,
        "x-internal-token": INTERNAL_LOOKUP_SECRET,
      },
      body: JSON.stringify(activity.value),
    });
    return new Response("ok");
  }

  if (!activity.text) return new Response("ignored");

  const tenantId = await resolveInnsynTenantId(aadTenantId);
  if (!tenantId) return new Response("no tenant");

  // 🔥 Fire-and-forget capture (NO await)
  fireHumanAnswerCapture(tenantId, aadTenantId, activity);

  const rag = await callRagQuery(tenantId, activity.text.trim());
  const card = buildAdaptiveCard(rag, tenantId, rag.qa_log_id);

  await sendTeamsReply(activity, card, creds, aadTenantId);
  return new Response("ok");
}

/********************************************************************************************
 * SERVER
 ********************************************************************************************/
serve(req => {
  const url = new URL(req.url);
  if (url.pathname === "/teams") return handleTeams(req);
  return new Response("ok");
});
