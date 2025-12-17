/********************************************************************************************
 * InnsynAI Teams Bridge – FINAL (RAG + Feedback, strict & correct)
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
 * MAIN HANDLER
 ********************************************************************************************/
async function handleTeams(req: Request): Promise<Response> {
  if (req.method !== "POST") return new Response("ok");

  // ---- Safe parse ----
  let activity: TeamsActivity;
  try {
    const raw = await req.text();
    if (!raw) return new Response("ok");
    activity = JSON.parse(raw);
  } catch {
    return new Response("ok");
  }

  console.log(
    "📨 Activity:",
    activity.type,
    "| hasText:",
    Boolean(activity.text),
    "| hasValue:",
    Boolean(activity.value),
  );

  // ---- JWT guard ----
  const auth = req.headers.get("Authorization");
  if (!auth) return new Response("ok");

  let creds;
  try {
    creds = await verifyJwt(auth);
  } catch (err) {
    console.error("❌ JWT failed", err);
    return new Response("unauthorized", { status: 401 });
  }

  const aadTenantId =
    activity.channelData?.tenant?.id || activity.conversation?.tenantId;
  if (!aadTenantId) return new Response("ok");

  const tenantId = await resolveTenantId(aadTenantId);
  if (!tenantId) return new Response("ok");

  /****************************
   * FEEDBACK PATH
   ****************************/
  if (activity.value?.action === "feedback") {
    const payload = {
      qa_log_id: activity.value.qa_log_id,
      feedback: activity.value.feedback,
      tenant_id: tenantId,
      source: "teams",
      teams_user_id: activity.value.teams_user_id ?? activity.from?.id ?? null,
      comment: activity.value.comment ?? undefined,
    };

    console.log("➡️ Forwarding feedback", payload);

    const res = await fetch(
      RAG_QUERY_URL.replace("/rag-query", "/feedback"),
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          apikey: SUPABASE_ANON_KEY,
          "x-internal-token": INTERNAL_LOOKUP_SECRET,
        },
        body: JSON.stringify(payload),
      },
    );

    if (!res.ok) {
      console.error("❌ Feedback failed", res.status, await res.text());
    } else {
      console.log("✅ Feedback saved");
    }

    return new Response("ok");
  }

  /****************************
   * MESSAGE → RAG
   ****************************/
  if (!activity.text) return new Response("ok");

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

  const rag = await ragRes.json();

  console.log("🧠 RAG qa_log_id:", rag.qa_log_id);

  /****************************
   * BUILD ADAPTIVE CARD
   ****************************/
  const actions = rag.qa_log_id
    ? [
        {
          type: "Action.Submit",
          title: "👍 Helpful",
          data: {
            action: "feedback",
            feedback: "up",
            qa_log_id: rag.qa_log_id,
            tenant_id: tenantId,
            source: "teams",
            teams_user_id: activity.from?.id ?? null,
          },
        },
        {
          type: "Action.Submit",
          title: "👎 Not helpful",
          data: {
            action: "feedback",
            feedback: "down",
            qa_log_id: rag.qa_log_id,
            tenant_id: tenantId,
            source: "teams",
            teams_user_id: activity.from?.id ?? null,
          },
        },
      ]
    : [];

  /****************************
   * SEND REPLY
   ****************************/
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
            content: {
              type: "AdaptiveCard",
              version: "1.4",
              body: [
                {
                  type: "TextBlock",
                  text: rag.answer ?? "No answer found.",
                  wrap: true,
                },
              ],
              actions,
            },
          },
        ],
        replyToId: activity.replyToId ?? activity.id,
      }),
    },
  );

  return new Response("ok");
}

/********************************************************************************************
 * SERVER
 ********************************************************************************************/
serve(req => {
  console.log("🔥 RAW REQUEST", req.method, new URL(req.url).pathname);
  if (new URL(req.url).pathname === "/teams") return handleTeams(req);
  return new Response("ok");
});
