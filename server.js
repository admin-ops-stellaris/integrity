import express from "express";
import cookieSession from "cookie-session";
import path from "path";
import { fileURLToPath } from "url";
import * as data from "./data.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const SESSION_SECRET = process.env.SESSION_SECRET || "dev-secret";
const app = express();

app.set("trust proxy", 1);
app.use(express.json());

app.use(
  cookieSession({
    name: "integrity_session",
    keys: [SESSION_SECRET],
    httpOnly: true,
    sameSite: "lax",
    secure: false,
    maxAge: 1000 * 60 * 60 * 10,
  })
);

// Dev mode: inject mock user
app.use((req, res, next) => {
  if (!req.session) req.session = {};
  if (!req.session.user) {
    req.session.user = { email: "dev@example.com", name: "Dev User" };
  }
  // Disable caching for dev
  res.setHeader("Cache-Control", "no-cache, no-store, must-revalidate");
  res.setHeader("Pragma", "no-cache");
  res.setHeader("Expires", "0");
  next();
});

// === API ENDPOINTS ===

app.post("/api/getEffectiveUserEmail", (req, res) => {
  res.json(req.session?.user?.email || "unknown@example.com");
});

app.post("/api/getRecentContacts", (req, res) => {
  res.json(data.getRecentContacts());
});

app.post("/api/searchContacts", (req, res) => {
  const query = req.body.args?.[0] || "";
  res.json(data.searchContacts(query));
});

app.post("/api/getContactById", (req, res) => {
  const id = req.body.args?.[0];
  res.json(data.getContactById(id));
});

app.post("/api/updateRecord", (req, res) => {
  const [table, id, field, value] = req.body.args || [];
  if (table === "Contacts") {
    data.updateContact(id, field, value);
  }
  res.json({ success: true });
});

app.post("/api/createRecord", (req, res) => {
  const [table, fields] = req.body.args || [];
  if (table === "Contacts") {
    const contact = data.createContact(fields);
    res.json(contact);
  } else {
    res.json({ success: false });
  }
});

app.post("/api/setSpouseStatus", (req, res) => {
  const [contactId, spouseId, action] = req.body.args || [];
  data.setSpouse(contactId, spouseId, action);
  res.json({ success: true });
});

app.post("/api/getLinkedOpportunities", (req, res) => {
  const ids = req.body.args?.[0] || [];
  res.json(data.getOpportunitiesById(ids));
});

// Static files
app.use(express.static(path.join(__dirname, "public")));

app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

// SPA fallback
app.get("*", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

app.use((err, req, res, next) => {
  console.error("Error:", err);
  res.status(500).json({ error: err.message });
});

const port = Number(process.env.PORT) || 5000;
app.listen(port, "0.0.0.0", () => console.log(`Server listening on ${port}`));
