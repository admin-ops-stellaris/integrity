import express from "express";
import cookieSession from "cookie-session";
import path from "path";
import { fileURLToPath } from "url";
import { Issuer, generators } from "openid-client";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const ALLOWED_GOOGLE_DOMAIN = process.env.ALLOWED_GOOGLE_DOMAIN;
const GOOGLE_CLIENT_ID = process.env.GOOGLE_CLIENT_ID;
const GOOGLE_CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET;
const SESSION_SECRET = process.env.SESSION_SECRET;

for (const [k, v] of Object.entries({
  ALLOWED_GOOGLE_DOMAIN,
  GOOGLE_CLIENT_ID,
  GOOGLE_CLIENT_SECRET,
  SESSION_SECRET,
})) {
  if (!v) throw new Error(`Missing ${k}`);
}

const CALLBACK_PATH = "/auth/google/callback";
const app = express();

app.use(express.json());
app.use(
  cookieSession({
    name: "integrity_session",
    keys: [SESSION_SECRET],
    httpOnly: true,
    sameSite: "lax",
    secure: process.env.NODE_ENV === "production",
    maxAge: 1000 * 60 * 60 * 10, // 10 hours
  }),
);

let clientCache = new Map();

function getBaseUrl(req) {
  const proto = req.headers["x-forwarded-proto"] || req.protocol;
  const host = req.headers["x-forwarded-host"] || req.get("host");
  return `${proto}://${host}`;
}

async function getClient(baseUrl) {
  if (clientCache.has(baseUrl)) return clientCache.get(baseUrl);

  const googleIssuer = await Issuer.discover("https://accounts.google.com");
  const client = new googleIssuer.Client({
    client_id: GOOGLE_CLIENT_ID,
    client_secret: GOOGLE_CLIENT_SECRET,
    redirect_uris: [`${baseUrl}${CALLBACK_PATH}`],
    response_types: ["code"],
  });

  clientCache.set(baseUrl, client);
  return client;
}

// ---- Auth routes ----
app.get("/auth/google", async (req, res, next) => {
  try {
    const baseUrl = getBaseUrl(req);
    const client = await getClient(baseUrl);

    const state = generators.state();
    const nonce = generators.nonce();
    req.session.oauth = { state, nonce, baseUrl };

    const authUrl = client.authorizationUrl({
      scope: "openid email profile",
      state,
      nonce,
      hd: ALLOWED_GOOGLE_DOMAIN, // hint only
      prompt: "select_account",
    });

    res.redirect(authUrl);
  } catch (e) {
    next(e);
  }
});

app.get(CALLBACK_PATH, async (req, res, next) => {
  try {
    const oauth = req.session?.oauth;
    if (!oauth?.state || !oauth?.nonce || !oauth?.baseUrl) {
      return res.status(400).send("Missing auth session. Try again.");
    }

    const client = await getClient(oauth.baseUrl);
    const params = client.callbackParams(req);

    const tokenSet = await client.callback(
      `${oauth.baseUrl}${CALLBACK_PATH}`,
      params,
      { state: oauth.state, nonce: oauth.nonce },
    );

    const claims = tokenSet.claims();

    // Enforce Workspace domain
    if (claims.hd !== ALLOWED_GOOGLE_DOMAIN) {
      req.session = null;
      return res.status(403).send("Unauthorized domain.");
    }

    req.session.user = {
      email: claims.email,
      name: claims.name,
      picture: claims.picture,
      hd: claims.hd,
    };
    delete req.session.oauth;

    res.redirect("/");
  } catch (e) {
    next(e);
  }
});

app.get("/logout", (req, res) => {
  req.session = null;
  res.redirect("/auth/google");
});

function requireAuth(req, res, next) {
  if (req.session?.user?.email) return next();
  return res.redirect("/auth/google");
}

// ---- Option A: gate everything ----
app.use(requireAuth);

// After auth, serve static files
app.get("/api/health", (req, res) => {
  res.json({
    ok: true,
    app: "Integrity",
    user: req.session?.user?.email || null,
    ts: new Date().toISOString()
  });
});
app.use(express.static(path.join(__dirname, "public")));

// SPA fallback
app.get("*", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

const port = Number(process.env.PORT) || 3000;
app.listen(port, () => console.log(`Server listening on ${port}`));
