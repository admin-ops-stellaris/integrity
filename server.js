import express from "express";
import cookieSession from "cookie-session";
import path from "path";
import { fileURLToPath } from "url";
import { Issuer, generators } from "openid-client";
import * as airtable from "./services/airtable.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const ALLOWED_GOOGLE_DOMAIN = process.env.ALLOWED_GOOGLE_DOMAIN;
const GOOGLE_CLIENT_ID = process.env.GOOGLE_CLIENT_ID;
const GOOGLE_CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET;
const SESSION_SECRET = process.env.SESSION_SECRET || "dev-secret";
const IS_REPLIT = !!process.env.REPL_ID;
const AUTH_DISABLED = process.env.AUTH_DISABLED === "true" || (IS_REPLIT && process.env.AUTH_DISABLED !== "false");

if (!AUTH_DISABLED) {
  for (const [k, v] of Object.entries({
    ALLOWED_GOOGLE_DOMAIN,
    GOOGLE_CLIENT_ID,
    GOOGLE_CLIENT_SECRET,
    SESSION_SECRET,
  })) {
    if (!v) throw new Error(`Missing ${k}`);
  }
}

const CALLBACK_PATH = "/auth/google/callback";
const app = express();

app.set("trust proxy", 1);
app.use(express.json());

app.use(
  cookieSession({
    name: "integrity_session",
    keys: [SESSION_SECRET],
    httpOnly: true,
    sameSite: "lax",
    secure: process.env.NODE_ENV === "production",
    maxAge: 1000 * 60 * 60 * 10,
  })
);

app.use((req, res, next) => {
  res.setHeader("Cache-Control", "no-cache, no-store, must-revalidate");
  res.setHeader("Pragma", "no-cache");
  res.setHeader("Expires", "0");
  next();
});

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

app.get("/auth/google", async (req, res, next) => {
  if (AUTH_DISABLED) return res.redirect("/");
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
      hd: ALLOWED_GOOGLE_DOMAIN,
      prompt: "select_account",
    });

    res.redirect(authUrl);
  } catch (e) {
    next(e);
  }
});

app.get(CALLBACK_PATH, async (req, res, next) => {
  if (AUTH_DISABLED) return res.redirect("/");
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
      { state: oauth.state, nonce: oauth.nonce }
    );

    const claims = tokenSet.claims();

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
  if (AUTH_DISABLED) return res.redirect("/");
  res.redirect("/auth/google");
});

function requireAuth(req, res, next) {
  if (AUTH_DISABLED) {
    if (!req.session) req.session = {};
    if (!req.session.user) {
      req.session.user = { email: "dev@example.com", name: "Dev User" };
    }
    return next();
  }
  if (req.session?.user?.email) return next();
  return res.redirect("/auth/google");
}

app.use(requireAuth);

app.get("/api/health", (req, res) => {
  res.json({
    ok: true,
    app: "Integrity",
    user: req.session?.user?.email || null,
    ts: new Date().toISOString()
  });
});

app.post("/api/getEffectiveUserEmail", (req, res) => {
  res.json(req.session?.user?.email || "unknown@example.com");
});

app.post("/api/getRecentContacts", async (req, res) => {
  try {
    const contacts = await airtable.getRecentContacts();
    res.json(contacts);
  } catch (err) {
    console.error("getRecentContacts error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/searchContacts", async (req, res) => {
  try {
    const query = req.body.args?.[0] || "";
    const contacts = await airtable.searchContacts(query);
    res.json(contacts);
  } catch (err) {
    console.error("searchContacts error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/getContactById", async (req, res) => {
  try {
    const id = req.body.args?.[0];
    const contact = await airtable.getContactById(id);
    res.json(contact);
  } catch (err) {
    console.error("getContactById error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/updateRecord", async (req, res) => {
  try {
    const [table, id, field, value] = req.body.args || [];
    if (table === "Contacts") {
      await airtable.updateContact(id, field, value);
    } else if (table === "Opportunities") {
      await airtable.updateOpportunity(id, field, value);
    }
    res.json({ success: true });
  } catch (err) {
    console.error("updateRecord error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/createRecord", async (req, res) => {
  try {
    const [table, fields] = req.body.args || [];
    if (table === "Contacts") {
      const contact = await airtable.createContact(fields);
      res.json(contact);
    } else {
      res.json({ success: false });
    }
  } catch (err) {
    console.error("createRecord error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/setSpouseStatus", async (req, res) => {
  try {
    const [contactId, spouseId, action] = req.body.args || [];
    await airtable.setSpouse(contactId, spouseId, action);
    res.json({ success: true });
  } catch (err) {
    console.error("setSpouseStatus error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/getLinkedOpportunities", async (req, res) => {
  try {
    const ids = req.body.args?.[0] || [];
    const opportunities = await airtable.getOpportunitiesById(ids);
    res.json(opportunities);
  } catch (err) {
    console.error("getLinkedOpportunities error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.use(express.static(path.join(__dirname, "public")));

app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

app.get("*", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

app.use((err, req, res, next) => {
  console.error("Error:", err);
  res.status(500).json({ error: err.message });
});

const port = Number(process.env.PORT) || 5000;
app.listen(port, "0.0.0.0", () => console.log(`Server listening on ${port}`));
