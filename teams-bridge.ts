/********************************************************************************************
 * InnsynAI Teams Bridge – FINAL (STORE / ADD-TO-TEAMS MODE, NO LEGACY)
 * DIAGNOSTIC BUILD — DO NOT REMOVE LOGS UNTIL 401 ROOT CAUSE CONFIRMED
 ********************************************************************************************/

import { serve } from "https://deno.land/std@0.224.0/http/server.ts";
import * as jose from "https://deno.land/x/jose@v5.4.0/index.ts";

/********************************************************************************************
 * ENV VARS
 ********************************************************************************************/
const env = Deno.env.toObject();

const {
  INTERNAL_LOOKUP_SECRET,
  TEAMS_TENANT_LOOKUP_URL,
  RAG_QUERY_URL,
  SUPABASE_ANON_KEY,
  TEAMS_BOT_APP_ID,
  TEAMS_BOT_APP_PASSWORD,
  CONNECT_URL,
} = env;

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

/** Startup identity proof */
console.log("🤖 BOT IDENTITY LOADED", {
  TEAMS_BOT_APP_ID,
  botSecretLength: TEAMS_BOT_APP_PASSWORD.length,
});

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
 * JWT VERIFICATION
 ********************************************************************************************/
async function verifyJwt(authHeader: string) {
  const token = authHeader.slice(7);
  const decoded = jose.decodeJwt(token);

  console.log("🔐 Incoming JWT", {
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
 * TENANT LOOKUP
 ********************************************************************************************/
async function resolveOrCreateTenantId(aadTenantId: string) {
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

  const text = await res.text();
  if (!res.ok) {
    console.error("❌ Tenant lookup failed", res.status, text);
    return null;
  }

  const json = JSON.parse(text);
  return json?.tenant_id ?? null;
}

/********************************************************************************************
 * BOT TOKEN (CRITICAL LOGGING)
 ********************************************************************************************/
async function getBotAccessToken(aadTenantId: string) {
  console.log("🔑 Minting bot token", {
    tenant: aadTenantId,
    client_id: TEAMS_BOT_APP_ID,
  });

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

  if (!res.ok || !json.access_token) {
    console.error("❌ BOT TOKEN FAILED", {
      status: res.status,
      tenant: aadTenantId,
      response: json,
    });
    throw new Error("Bot token failure");
  }

  console.log("✅ Bot token minted");
  return json.access_token as string;
}

/********************************************************************************************
 * TEAMS SEND (DEEP 401 LOGGING)
 ********************************************************************************************/
async function sendTeams(
  url: string,
  token: string,
  payload: any,
  method: "POST" | "PUT",
) {
  const res = await fetch(url, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(payload),
  });

  const text = await res.text().catch(() => "");
  if (!res.ok) {
    console.error("❌ TEAMS API ERROR", {
      method,
      url,
      status: res.status,
      wwwAuth: res.headers.get("www-authenticate"),
      body: text,
    });
  }

  let id: string | null = null;
  try {
    const json = JSON.parse(text);
    id = json?.id ?? null;
  } catch {}

  return { ok: res.ok, status: res.status, id };
}

/********************************************************************************************
 * MAIN HANDLER
 ********************************************************************************************/
async function handleTeams(req: Request): Promise<Response> {
  const activity = await req.json();

  const auth = req.headers.get("Authorization");
  if (!auth?.startsWith("Bearer ")) return new Response("ok");

  await verifyJwt(auth);

  const aadTenantId =
    activity.channelData?.tenant?.id || activity.conversation?.tenantId;

  console.log("📨 Activity received", {
    botAppId: TEAMS_BOT_APP_ID,
    aadTenantId,
    serviceUrl: activity.serviceUrl,
  });

  const tenantId = await resolveOrCreateTenantId(aadTenantId);
  console.log("🧭 Tenant resolved", tenantId);

  const token = await getBotAccessToken(aadTenantId);

  const placeholderUrl =
    `${activity.serviceUrl}/v3/conversations/${encodeURIComponent(
      activity.conversation.id,
    )}/activities`;

  const placeholder = await sendTeams(
    placeholderUrl,
    token,
    { type: "message", text: "⏳ Working on it…" },
    "POST",
  );

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

  const cardPayload = {
    type: "message",
    text: rag.answer ?? "No answer found.",
  };

  if (placeholder.id) {
    const putUrl =
      `${activity.serviceUrl}/v3/conversations/${encodeURIComponent(
        activity.conversation.id,
      )}/activities/${placeholder.id}`;

    await sendTeams(putUrl, token, cardPayload, "PUT");
  } else {
    await sendTeams(placeholderUrl, token, cardPayload, "POST");
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
