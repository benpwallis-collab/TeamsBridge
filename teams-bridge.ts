/********************************************************************************************
 * InnsynAI Teams Bridge – FINAL (STORE / ADD-TO-TEAMS MODE, NO LEGACY)
 *
 * ✅ Single global bot identity (TEAMS_BOT_APP_ID / TEAMS_BOT_APP_PASSWORD)
 * ✅ Multi-tenant routing by AAD tenant id (activity.channelData.tenant.id)
 * ✅ Auto-provision tenant mapping via teams-tenant-lookup (auto_provision: true)
 * ✅ Inline RAG execution (no background async)
 * ✅ Immediate placeholder message, then PATCH with adaptive card + sources + feedback
 * ✅ Feedback preserved (/feedback) using x-internal-token
 * ✅ Strong, Render-friendly logging (aadTenantId + resolved tenantId + botAppId)
 *
 * REQUIRED ENVS (Render):
 *   INTERNAL_LOOKUP_SECRET
 *   TEAMS_TENANT_LOOKUP_URL
 *   RAG_QUERY_URL                  (Supabase edge: .../rag-query)
 *   SUPABASE_ANON_KEY
 *   TEAMS_BOT_APP_ID               (global bot app id)
 *   TEAMS_BOT_APP_PASSWORD         (global bot secret value)
 *
 * Optional:
 *   CONNECT_URL                    (if tenant created but no docs, show link; otherwise ignored)
 ********************************************************************************************/

import { serve } from "https://deno.land/std@0.224.0/http/server.ts";
import * as jose from "https://deno.land/x/jose@v5.4.0/index.ts";

/********************************************************************************************
 * ENV VARS
 ********************************************************************************************/
const env = Deno.env.toObject();

const INTERNAL_LOOKUP_SECRET = env.INTERNAL_LOOKUP_SECRET;
const TEAMS_TENANT_LOOKUP_URL = env.TEAMS_TENANT_LOOKUP_URL;
const RAG_QUERY_URL = env.RAG_QUERY_URL;
const SUPABASE_ANON_KEY = env.SUPABASE_ANON_KEY;

const TEAMS_BOT_APP_ID = env.TEAMS_BOT_APP_ID;
const TEAMS_BOT_APP_PASSWORD = env.TEAMS_BOT_APP_PASSWORD;

const CONNECT_URL = env.CONNECT_URL; // optional

if (
  !INTERNAL_LOOKUP_SECRET ||
  !TEAMS_TENANT_LOOKUP_URL ||
  !RAG_QUERY_URL ||
  !SUPABASE_ANON_KEY ||
  !TEAMS_BOT_APP_ID ||
  !TEAMS_BOT_APP_PASSWORD
) {
  console.error("❌ Missing env vars. Required: INTERNAL_LOOKUP_SECRET, TEAMS_TENANT_LOOKUP_URL, RAG_QUERY_URL, SUPABASE_ANON_KEY, TEAMS_BOT_APP_ID, TEAMS_BOT_APP_PASSWORD");
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
 * HELPERS
 ********************************************************************************************/
function safeStr(v: unknown, max = 120): string {
  const s = typeof v === "string" ? v : JSON.stringify(v);
  return s.length > max ? s.slice(0, max) + "…" : s;
}

function getPlatformLabel(source: string): string {
  const labels: Record<string, string> = {
    notion: "Notion",
    confluence: "Confluence",
    gitlab: "GitLab",
    google_drive: "Google Drive",
    sharepoint: "SharePoint",
    manual: "Manual Upload",
    slack: "Slack",
    teams: "Teams",
  };
  return labels[source] || source || "Unknown";
}

function getRelativeDate(dateStr: string): string {
  if (!dateStr) return "recently";
  const date = new Date(dateStr);
  const now = new Date();
  const diffMs = now.getTime() - date.getTime();
  const diffHours = Math.floor(diffMs / (1000 * 60 * 60));
  const diffDays = Math.floor(diffHours / 24);

  if (diffHours < 1) return "just now";
  if (diffHours < 24) return `${diffHours} hour${diffHours > 1 ? "s" : ""} ago`;
  if (diffDays < 7) return `${diffDays} day${diffDays > 1 ? "s" : ""} ago`;
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
 * TENANT LOOKUP (AUTO-PROVISION)
 ********************************************************************************************/
async function resolveOrCreateTenantId(aadTenantId: string): Promise<string | null> {
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
    console.error("❌ teams-tenant-lookup failed", res.status, await res.text().catch(() => ""));
    return null;
  }

  const json = await res.json().catch(() => null);
  return json?.tenant_id ?? null;
}

/********************************************************************************************
 * JWT VERIFICATION (SINGLE BOT)
 ********************************************************************************************/
async function verifyJwt(authHeader: string) {
  const token = authHeader.slice(7);

  const keyStore = jose.createLocalJWKSet(await getJwks());
  await jose.jwtVerify(token, keyStore, {
    issuer: "https://api.botframework.com",
    audience: TEAMS_BOT_APP_ID,
  });
}

/********************************************************************************************
 * BOT TOKEN (SINGLE BOT, TENANT-SPECIFIC ISSUER)
 ********************************************************************************************/
async function getBotAccessToken(aadTenantId: string) {
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

  const json = await res.json().catch(() => ({}));
  if (!json.access_token) {
    console.error("❌ Failed to get bot token", json);
    throw new Error("Bot token failure");
  }

  return json.access_token as string;
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
 * SEND MESSAGE HELPERS
 ********************************************************************************************/
async function sendPlaceholder(activity: TeamsActivity, accessToken: string): Promise<string | null> {
  if (!activity.serviceUrl || !activity.conversation?.id) return null;

  const res = await fetch(
    `${activity.serviceUrl}/v3/conversations/${encodeURIComponent(activity.conversation.id)}/activities`,
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

  const json = await res.json().catch(() => ({}));
  return json?.id ?? null;
}

async function patchWithCard(
  activity: TeamsActivity,
  accessToken: string,
  placeholderActivityId: string,
  rag: any,
) {
  if (!activity.serviceUrl || !activity.conversation?.id) return;

  const actions = rag?.qa_log_id
    ? [
        {
          type: "Action.Submit",
          title: "👍 Helpful",
          data: { action: "feedback", feedback: "up", qa_log_id: rag.qa_log_id },
        },
        {
          type: "Action.Submit",
          title: "👎 Not helpful",
          data: { action: "feedback", feedback: "down", qa_log_id: rag.qa_log_id },
        },
      ]
    : [];

  // If RAG returns setup prompt, optionally append a connect link
  let answerText = rag?.answer ?? "No answer found.";
  if (rag?.status === "no_documents" && CONNECT_URL) {
    answerText = `${answerText}\n\nSetup: ${CONNECT_URL}`;
  }

  const cardBody = [
    {
      type: "TextBlock",
      text: answerText,
      wrap: true,
    },
    ...formatSourcesForCard(rag?.sources ?? []),
  ];

  const res = await fetch(
    `${activity.serviceUrl}/v3/conversations/${encodeURIComponent(activity.conversation.id)}/activities/${placeholderActivityId}`,
    {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${accessToken}`,
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

  if (!res.ok) {
    console.error("❌ PATCH failed", res.status, await res.text().catch(() => ""));
  } else {
    console.log("✅ Message updated");
  }
}

/********************************************************************************************
 * MAIN HANDLER
 ********************************************************************************************/
async function handleTeams(req: Request): Promise<Response> {
  if (req.method !== "POST") return new Response("ok");

  let activity: TeamsActivity;
  try {
    activity = JSON.parse(await req.text());
  } catch {
    console.warn("⚠️ Invalid JSON body");
    return new Response("ok");
  }

  const auth = req.headers.get("Authorization");
  if (!auth?.startsWith("Bearer ")) {
    console.warn("⚠️ Missing/invalid Authorization header");
    return new Response("ok");
  }

  try {
    await verifyJwt(auth);
  } catch (e) {
    console.error("❌ JWT verification failed", e);
    return new Response("unauthorized", { status: 401 });
  }

  const aadTenantId =
    activity.channelData?.tenant?.id || activity.conversation?.tenantId;

  console.log(
    "📨 Activity",
    "botAppId=", TEAMS_BOT_APP_ID,
    "aadTenantId=", aadTenantId ?? "missing",
    "conversationId=", activity.conversation?.id ?? "missing",
    "activityId=", activity.id ?? "missing",
    "from=", activity.from?.id ?? "missing",
    "type=", activity.type ?? "missing",
    "text=", activity.text ? safeStr(activity.text, 80) : "none",
  );

  if (!aadTenantId) {
    console.warn("⚠️ Missing AAD tenant id");
    return new Response("ok");
  }

  const tenantId = await resolveOrCreateTenantId(aadTenantId);
  console.log("🧭 Tenant resolved", "aadTenantId=", aadTenantId, "-> tenantId=", tenantId ?? "NOT_FOUND");

  if (!tenantId) {
    console.error("❌ Unable to resolve or create tenant for AAD tenant:", aadTenantId);
    return new Response("ok");
  }

  /****************************
   * FEEDBACK
   ****************************/
  if (activity.value?.action === "feedback") {
    console.log("👍 Feedback received", safeStr(activity.value, 400));

    // Forward to feedback endpoint (internal token, no tenant spoof)
    const feedbackRes = await fetch(RAG_QUERY_URL.replace(/\/rag-query\/?$/, "/feedback"), {
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

    if (!feedbackRes.ok) {
      console.error("❌ Feedback forward failed", feedbackRes.status, await feedbackRes.text().catch(() => ""));
    }

    return new Response("ok");
  }

  if (!activity.text?.trim()) return new Response("ok");

  /****************************
   * SEND PLACEHOLDER
   ****************************/
  let accessToken: string;
  try {
    accessToken = await getBotAccessToken(aadTenantId);
  } catch (e) {
    console.error("❌ Bot token failure (cannot respond)", e);
    return new Response("ok");
  }

  const placeholderActivityId = await sendPlaceholder(activity, accessToken);
  console.log("🕒 Placeholder sent", "placeholderActivityId=", placeholderActivityId ?? "missing");

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
    console.error("❌ RAG failed", ragRes.status, await ragRes.text().catch(() => ""));
    // If we can't patch, at least don't crash
    return new Response("ok");
  }

  const rag = await ragRes.json().catch(() => null);
  console.log(
    "🧠 RAG completed",
    "tenantId=", tenantId,
    "qa_log_id=", rag?.qa_log_id ?? "none",
    "sources=", rag?.sources?.length ?? 0,
  );

  /****************************
   * PATCH PLACEHOLDER WITH ANSWER CARD
   ****************************/
  if (!placeholderActivityId) {
    // If placeholder id missing, nothing to patch; return ok
    console.warn("⚠️ Missing placeholderActivityId; cannot patch");
    return new Response("ok");
  }

  // Re-mint token for patch (safe, and avoids token expiry edge cases)
  const patchToken = await getBotAccessToken(aadTenantId);
  await patchWithCard(activity, patchToken, placeholderActivityId, rag);

  return new Response("ok");
}

/********************************************************************************************
 * SERVER
 ********************************************************************************************/
serve((req) => {
  const path = new URL(req.url).pathname;
  console.log("➡️ Request", req.method, path);
  if (path === "/teams") return handleTeams(req);
  return new Response("ok");
});
