import { serve } from "https://deno.land/std@0.224.0/http/server.ts";
import * as jose from "https://deno.land/x/jose@v5.4.0/index.ts";

/* ================= ENV ================= */

const env = Deno.env.toObject();

const INTERNAL_LOOKUP_SECRET = env.INTERNAL_LOOKUP_SECRET;
const TEAMS_TENANT_LOOKUP_URL = env.TEAMS_TENANT_LOOKUP_URL;
const RAG_QUERY_URL = env.RAG_QUERY_URL;
const SUPABASE_ANON_KEY = env.SUPABASE_ANON_KEY;
const TEAMS_BOT_APP_ID = env.TEAMS_BOT_APP_ID;
const TEAMS_BOT_APP_PASSWORD = env.TEAMS_BOT_APP_PASSWORD;

if (
  !INTERNAL_LOOKUP_SECRET ||
  !TEAMS_TENANT_LOOKUP_URL ||
  !RAG_QUERY_URL ||
  !SUPABASE_ANON_KEY ||
  !TEAMS_BOT_APP_ID ||
  !TEAMS_BOT_APP_PASSWORD
) {
  console.error("Missing env vars");
  Deno.exit(1);
}

/* ================= JWKS ================= */

const OPENID_CONFIG_URL =
  "https://login.botframework.com/v1/.well-known/openidconfiguration";

let jwks: jose.JSONWebKeySet | null = null;

async function getJwks() {
  if (jwks) return jwks;
  const meta = await fetch(OPENID_CONFIG_URL);
  const metaJson = await meta.json();
  const keys = await fetch(metaJson.jwks_uri);
  jwks = await keys.json();
  return jwks!;
}

/* ================= AUTH ================= */

async function verifyJwt(authHeader: string) {
  const token = authHeader.slice(7);
  const keyStore = jose.createLocalJWKSet(await getJwks());

  await jose.jwtVerify(token, keyStore, {
    issuer: "https://api.botframework.com",
    audience: TEAMS_BOT_APP_ID
  });
}

/* ================= BOT TOKEN ================= */

async function getBotToken(aadTenantId: string) {
  const res = await fetch(
    "https://login.microsoftonline.com/" +
      aadTenantId +
      "/oauth2/v2.0/token",
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type: "client_credentials",
        client_id: TEAMS_BOT_APP_ID,
        client_secret: TEAMS_BOT_APP_PASSWORD,
        scope: "https://api.botframework.com/.default"
      })
    }
  );

  const json = await res.json();

  if (!json.access_token) {
    console.error("Token failure", json);
    throw new Error("bot token failure");
  }

  return json.access_token;
}

/* ================= TENANT ================= */

async function resolveTenant(aadTenantId: string) {
  const res = await fetch(TEAMS_TENANT_LOOKUP_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      apikey: SUPABASE_ANON_KEY,
      "x-internal-token": INTERNAL_LOOKUP_SECRET
    },
    body: JSON.stringify({
      teams_tenant_id: aadTenantId,
      auto_provision: true
    })
  });

  if (!res.ok) return null;
  const json = await res.json();
  return json.tenant_id || null;
}

/* ================= HANDLER ================= */

async function handleTeams(req: Request) {
  if (req.method !== "POST") return new Response("ok");

  const auth = req.headers.get("Authorization");
  if (!auth || !auth.startsWith("Bearer ")) return new Response("ok");

  const text = await req.text();
  let activity: any;
  try {
    activity = JSON.parse(text);
  } catch {
    return new Response("ok");
  }

  await verifyJwt(auth);

  const aadTenantId =
    activity.channelData?.tenant?.id ||
    activity.conversation?.tenantId;

  if (!aadTenantId) return new Response("ok");

  const tenantId = await resolveTenant(aadTenantId);
  if (!tenantId) return new Response("ok");

  const serviceUrl = activity.serviceUrl.replace(/\/$/, "");
  const token = await getBotToken(aadTenantId);

  const postUrl =
    serviceUrl +
    "/v3/conversations/" +
    encodeURIComponent(activity.conversation.id) +
    "/activities";

  await fetch(postUrl, {
    method: "POST",
    headers: {
      Authorization: "Bearer " + token,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      type: "message",
      text: "Hello from InnsynAI 👋"
    })
  });

  return new Response("ok");
}

/* ================= SERVER ================= */

serve((req) => {
  if (new URL(req.url).pathname === "/teams") {
    return handleTeams(req);
  }
  return new Response("ok");
});
