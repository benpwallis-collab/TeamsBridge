/********************************************************************************************
 * InnsynAI Teams Bridge – FINAL (STORE / ADD-TO-TEAMS MODE)
 *
 * CRITICAL FIX:
 *   ✅ Bot access tokens are minted against botframework.com
 *      NOT customer tenant IDs
 *
 * This resolves:
 *   - 401 "Authorization has been denied"
 *   - Missing service principal errors
 *   - Placeholder + PATCH failures
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
  TEAMS_BOT_APP_ID,
  TEAMS_BOT_APP_PASSWORD,
  CONNECT_URL,
} = Deno.env.toObject();

if (
  !INTERNAL_LOOKUP_SECRET ||
  !TEAMS_TENANT_LOOKUP_URL ||
  !RAG_QUERY_URL ||
  !SUPABASE_ANON_KEY ||
  !TEAMS_BOT_APP_ID ||
  !TEAMS_BOT_APP_PASSWORD
) {
  console.error("❌ Missing required env vars");
  Deno.exit(1);
}

/********************************************************************************************
 * BOTFRAMEWORK JWKS
 ********************************************************************************************/
const OPENID_CONFIG_URL =
  "https://login.botframework.com/v1/.well-known/openidconfiguration";

let jwks: jose.JSONWebKeySet | null = null;

async function getJwks() {
  if (jwks) return jwks;
  const meta = await fetch(OPENID_CONFIG_URL).then(r => r.json());
  jwks = await fetch(meta.jwks_uri).then(r => r.json());
  return jwks!;
}

/********************************************************************************************
 * HELPERS
 ********************************************************************************************/
const log = (...args: any[]) => console.log(...args);

function formatSourcesForCard(sources: any[]) {
  if (!sources?.length) return [];
  return [
    { type: "TextBlock", text: "**Sources:**", wrap: true },
    ...sources.map(s => ({
      type: "TextBlock",
      wrap: true,
      text: s.url
        ? `• [${s.title}](${s.url})`
        : `• ${s.title}`,
    })),
  ];
}

/********************************************************************************************
 * TENANT LOOKUP (AUTO-PROVISION)
 ********************************************************************************************/
async function resolveTenantId(aadTenantId: string): Promise<string | null> {
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
    log("❌ tenant lookup failed", res.status);
    return null;
  }

  const json = await res.json();
  return json?.tenant_id ?? null;
}

/********************************************************************************************
 * JWT VERIFY (BOTFRAMEWORK → YOUR BOT)
 ********************************************************************************************/
async function verifyJwt(authHeader: string) {
  const token = authHeader.slice(7);
  const decoded = jose.decodeJwt(token);

  log("🔐 Incoming JWT", {
    aud: decoded.aud,
    iss: decoded.iss,
    tid: decoded.tid,
  });

  const keyStore = jose.createLocalJWKSet(await getJwks());

  await jose.jwtVerify(token, keyStore, {
    issuer: "https://api.botframework.com",
    audience: TEAMS_BOT_APP_ID,
  });
}

/********************************************************************************************
 * 🔑 BOT TOKEN (CRITICAL FIX)
 * MUST USE botframework.com
 ********************************************************************************************/
async function getBotAccessToken() {
  log("🔑 Minting bot token", {
    tenant: "botframework.com",
    client_id: TEAMS_BOT_APP_ID,
  });

  const res = await fetch(
    "https://login.microsoftonline.com/botframework.com/oauth2/v2.0/token",
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

  const json = await res.json();
  if (!json.access_token) {
    log("❌ Token mint failed", json);
    throw new Error("Bot token failure");
  }

  log("✅ Bot token minted");
  return json.access_token as string;
}

/********************************************************************************************
 * MAIN HANDLER
 ********************************************************************************************/
async function handleTeams(req: Request): Promise<Response> {
  if (req.method !== "POST") return new Response("ok");

  const activity = await req.json();
  const auth = req.headers.get("Authorization");

  if (!auth?.startsWith("Bearer ")) {
    log("⚠️ Missing auth header");
    return new Response("ok");
  }

  await verifyJwt(auth);

  const aadTenantId =
    activity.channelData?.tenant?.id || activity.conversation?.tenantId;

  log("📨 Activity received", {
    botAppId: TEAMS_BOT_APP_ID,
    aadTenantId,
    serviceUrl: activity.serviceUrl,
  });

  const tenantId = await resolveTenantId(aadTenantId);
  log("🧭 Tenant resolved", tenantId);

  if (!tenantId) return new Response("ok");

  const token = await getBotAccessToken();

  /****************************
   * PLACEHOLDER
   ****************************/
  const postUrl =
    `${activity.serviceUrl}/v3/conversations/${encodeURIComponent(
      activity.conversation.id,
    )}/activities`;

  const placeholderRes = await fetch(postUrl, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      type: "message",
      text: "⏳ Working on it…",
      replyToId: activity.replyToId ?? activity.id,
    }),
  });

  let placeholderId: string | null = null;
  try {
    const json = await placeholderRes.json();
    placeholderId = json?.id ?? null;
  } catch {}

  /****************************
   * RAG
   ****************************/
  const ragRes = await fetch(RAG_QUERY_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      apikey: SUPABASE_ANON_KEY,
      "x-tenant-id": tenantId,
    },
    body: JSON.stringify({
      question: activity.text,
      source: "teams",
    }),
  });

  const rag = await ragRes.json();

  const card = {
    type: "message",
    attachments: [
      {
        contentType: "application/vnd.microsoft.card.adaptive",
        content: {
          type: "AdaptiveCard",
          version: "1.4",
          body: [
            { type: "TextBlock", wrap: true, text: rag.answer },
            ...formatSourcesForCard(rag.sources),
          ],
        },
      },
    ],
  };

  /****************************
   * PATCH OR FALLBACK POST
   ****************************/
  if (placeholderId) {
    await fetch(`${postUrl}/${placeholderId}`, {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(card),
    });
  } else {
    await fetch(postUrl, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(card),
    });
  }

  return new Response("ok");
}

/********************************************************************************************
 * SERVER
 ********************************************************************************************/
serve(req => {
  log("➡️ Request", req.method, new URL(req.url).pathname);
  if (new URL(req.url).pathname === "/teams") return handleTeams(req);
  return new Response("ok");
});
