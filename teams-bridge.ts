/********************************************************************************************
 * InnsynAI Teams Bridge – FINAL, RENDER-SAFE
 * - Immediate "Working on it…" reply
 * - Inline RAG execution (no background async)
 * - PUT placeholder message with answer
 * - Feedback preserved
 * - Source citations with platform labels
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
  const meta = await fetch(OPENID_CONFIG_URL).then(r => r.json());
  jwks = await fetch(meta.jwks_uri).then(r => r.json());
  return jwks!;
}

/********************************************************************************************
 * HELPER FUNCTIONS
 ********************************************************************************************/
function getPlatformLabel(source: string): string {
  const labels: Record<string, string> = {
    notion: 'Notion',
    confluence: 'Confluence',
    gitlab: 'GitLab',
    google_drive: 'Google Drive',
    sharepoint: 'SharePoint',
    manual: 'Manual Upload',
    slack: 'Slack',
    teams: 'Teams',
  };
  return labels[source] || source || 'Unknown';
}

function getRelativeDate(dateStr: string): string {
  if (!dateStr) return 'recently';
  const date = new Date(dateStr);
  const now = new Date();
  const diffMs = now.getTime() - date.getTime();
  const diffHours = Math.floor(diffMs / (1000 * 60 * 60));
  const diffDays = Math.floor(diffHours / 24);
  
  if (diffHours < 1) return 'just now';
  if (diffHours < 24) return `${diffHours} hour${diffHours > 1 ? 's' : ''} ago`;
  if (diffDays < 7) return `${diffDays} day${diffDays > 1 ? 's' : ''} ago`;
  return date.toLocaleDateString();
}

function formatSourcesForCard(sources: any[]): any[] {
  if (!sources?.length) return [];
  
  return [
    {
      type: "TextBlock",
      text: "**Sources:**",
      wrap: true,
      spacing: "Medium",
    },
    ...sources.map((s) => ({
      type: "TextBlock",
      text: s.url 
        ? `• [${s.title}](${s.url}) — ${getPlatformLabel(s.source)} (Updated ${getRelativeDate(s.updated_at)})`
        : `• ${s.title} — ${getPlatformLabel(s.source)} (Updated ${getRelativeDate(s.updated_at)})`,
      wrap: true,
      spacing: "Small",
    })),
  ];
}

/********************************************************************************************
 * TENANT LOOKUP
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

async function resolveTenantId(aadTenantId: string) {
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

  let activity: TeamsActivity;
  try {
    activity = JSON.parse(await req.text());
  } catch {
    console.warn("⚠️ Invalid JSON body");
    return new Response("ok");
  }

  console.log(
    "📨 Activity",
    activity.type,
    "| text:",
    activity.text?.slice(0, 80),
  );

  const auth = req.headers.get("Authorization");
  if (!auth) {
    console.warn("⚠️ Missing Authorization header");
    return new Response("ok");
  }

  let creds;
  try {
    creds = await verifyJwt(auth);
  } catch (err) {
    console.error("❌ JWT verification failed", err);
    return new Response("unauthorized", { status: 401 });
  }

  const aadTenantId =
    activity.channelData?.tenant?.id || activity.conversation?.tenantId;
  if (!aadTenantId) {
    console.warn("⚠️ Missing AAD tenant id");
    return new Response("ok");
  }

  const tenantId = await resolveTenantId(aadTenantId);
  if (!tenantId) {
    console.warn("⚠️ Tenant not found");
    return new Response("ok");
  }

  /****************************
   * FEEDBACK
   ****************************/
  if (activity.value?.action === "feedback") {
    console.log("👍 Feedback received", activity.value);

    await fetch(RAG_QUERY_URL.replace("/rag-query", "/feedback"), {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        apikey: SUPABASE_ANON_KEY,
        "x-internal-token": INTERNAL_LOOKUP_SECRET,
      },
      body: JSON.stringify({
        qa_log_id: activity.value.qa_log_id,
        feedback: activity.value.feedback,
        tenant_id: tenantId,
        source: "teams",
        teams_user_id: activity.from?.id ?? null,
      }),
    });

    return new Response("ok");
  }

  if (!activity.text) return new Response("ok");

  /****************************
   * SEND PLACEHOLDER
   ****************************/
  const accessToken = await getBotAccessToken(aadTenantId, creds);

  const placeholderRes = await fetch(
    `${activity.serviceUrl}/v3/conversations/${encodeURIComponent(
      activity.conversation!.id,
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

  console.log("🕒 Placeholder sent", placeholderActivityId);

  /****************************
   * RAG QUERY (INLINE)
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
  console.log("🧠 RAG completed", rag.qa_log_id, "| sources:", rag.sources?.length ?? 0);

  /****************************
   * BUILD ACTIONS
   ****************************/
  const actions = rag.qa_log_id
    ? [
        {
          type: "Action.Submit",
          title: "👍 Helpful",
          data: { action: "feedback", feedback: "up", qa_log_id: rag.qa_log_id },
        },
        {
          type: "Action.Submit",
          title: "👎 Not helpful",
          data: {
            action: "feedback",
            feedback: "down",
            qa_log_id: rag.qa_log_id,
          },
        },
      ]
    : [];

  /****************************
   * BUILD CARD BODY WITH SOURCES
   ****************************/
  const cardBody = [
    {
      type: "TextBlock",
      text: rag.answer ?? "No answer found.",
      wrap: true,
    },
    ...formatSourcesForCard(rag.sources),
  ];

  /****************************
   * PUT PLACEHOLDER
   ****************************/
  const patchToken = await getBotAccessToken(aadTenantId, creds);

  const patchRes = await fetch(
    `${activity.serviceUrl}/v3/conversations/${encodeURIComponent(
      activity.conversation!.id,
    )}/activities/${placeholderActivityId}`,
    {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${patchToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        type: "message",
        attachments: [
          {
            contentType: "application/vnd.microsoft.card.adaptive",
            content: {
              type: "AdaptiveCard",
              version: "1.4",
              body: cardBody,
              actions,
            },
          },
        ],
      }),
    },
  );

  if (!patchRes.ok) {
    console.error("❌ PATCH failed", patchRes.status, await patchRes.text());
  } else {
    console.log("✅ Message updated with sources");
  }

  return new Response("ok");
}

/********************************************************************************************
 * SERVER
 ********************************************************************************************/
serve(req => {
  console.log("➡️ Request", req.method, new URL(req.url).pathname);
  if (new URL(req.url).pathname === "/teams") return handleTeams(req);
  return new Response("ok");
});
