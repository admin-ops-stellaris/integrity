import express from "express";
import cookieSession from "cookie-session";
import path from "path";
import { fileURLToPath } from "url";
import { Issuer, generators } from "openid-client";
import * as airtable from "./services/airtable.js";
import * as gmail from "./services/gmail.js";

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
    maxAge: 1000 * 60 * 60 * 24 * 30, // 30 days
  })
);

app.use((req, res, next) => {
  res.setHeader("Cache-Control", "no-cache, no-store, must-revalidate");
  res.setHeader("Pragma", "no-cache");
  res.setHeader("Expires", "0");
  next();
});

let clientCache = new Map();

// Helper to parse Taco appointment date formats to ISO string (Australia/Sydney timezone)
// Supports formats like: "13/01/26 3:30 PM", "Mon 5 Jan 2026 3:30 PM", "Thursday 15/01/26 at 9:00am"
function parseTacoAppointmentTime(tacoDateStr) {
  if (!tacoDateStr || typeof tacoDateStr !== 'string') return null;
  
  const str = tacoDateStr.trim();
  if (!str) return null;
  
  try {
    // If already ISO format, return as-is
    if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(str)) {
      return str;
    }
    
    let day, month, year, hour, minute;
    
    // Try "DayName DD/MM/YY at H:MMam/pm" format (e.g., "Thursday 15/01/26 at 9:00am")
    const dayNameAtFormat = str.match(/^(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)\s+(\d{1,2})\/(\d{1,2})\/(\d{2,4})\s+at\s+(\d{1,2}):(\d{2})\s*(am|pm)?$/i);
    if (dayNameAtFormat) {
      day = parseInt(dayNameAtFormat[1]);
      month = parseInt(dayNameAtFormat[2]);
      year = parseInt(dayNameAtFormat[3]);
      if (year < 100) year += 2000;
      
      hour = parseInt(dayNameAtFormat[4]);
      minute = parseInt(dayNameAtFormat[5]);
      const ampm = (dayNameAtFormat[6] || '').toUpperCase();
      
      if (ampm === 'PM' && hour !== 12) hour += 12;
      if (ampm === 'AM' && hour === 12) hour = 0;
    }
    
    // Try DD/MM/YY HH:MM AM/PM format (e.g., "13/01/26 3:30 PM")
    if (!day) {
      const ddmmyyAmpm = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})\s+(\d{1,2}):(\d{2})\s*(AM|PM)?$/i);
      if (ddmmyyAmpm) {
        day = parseInt(ddmmyyAmpm[1]);
        month = parseInt(ddmmyyAmpm[2]);
        year = parseInt(ddmmyyAmpm[3]);
        if (year < 100) year += 2000;
        
        hour = parseInt(ddmmyyAmpm[4]);
        minute = parseInt(ddmmyyAmpm[5]);
        const ampm = (ddmmyyAmpm[6] || '').toUpperCase();
        
        if (ampm === 'PM' && hour !== 12) hour += 12;
        if (ampm === 'AM' && hour === 12) hour = 0;
      }
    }
    
    // Try "Day D Mon YYYY H:MM AM/PM" format (e.g., "Mon 5 Jan 2026 3:30 PM")
    if (!day) {
      const dayMonFormat = str.match(/^(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s+(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{4})\s+(\d{1,2}):(\d{2})\s*(AM|PM)?$/i);
      if (dayMonFormat) {
        day = parseInt(dayMonFormat[1]);
        const monthNames = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
        month = monthNames[dayMonFormat[2].toLowerCase()];
        year = parseInt(dayMonFormat[3]);
        
        hour = parseInt(dayMonFormat[4]);
        minute = parseInt(dayMonFormat[5]);
        const ampm = (dayMonFormat[6] || '').toUpperCase();
        
        if (ampm === 'PM' && hour !== 12) hour += 12;
        if (ampm === 'AM' && hour === 12) hour = 0;
      }
    }
    
    // If we parsed the components, build a datetime-local format string
    // This format is accepted by Airtable for datetime fields
    if (day && month && year && hour !== undefined && minute !== undefined) {
      const isoStr = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}T${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}`;
      console.log("Parsed Taco time:", str, "->", isoStr);
      return isoStr;
    }
    
    // Try native Date parsing as fallback
    const parsed = new Date(str);
    if (!isNaN(parsed.getTime())) {
      return parsed.toISOString();
    }
    
    console.warn("Could not parse Taco appointment time:", str);
    return null;
  } catch (err) {
    console.error("Error parsing Taco appointment time:", err.message);
    return null;
  }
}

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
      scope: "openid email profile https://www.googleapis.com/auth/gmail.send",
      state,
      nonce,
      hd: ALLOWED_GOOGLE_DOMAIN,
      access_type: "offline",
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
    
    // Store Gmail tokens for sending emails
    req.session.gmailTokens = {
      access_token: tokenSet.access_token,
      refresh_token: tokenSet.refresh_token,
      expires_at: tokenSet.expires_at ? tokenSet.expires_at * 1000 : Date.now() + 3600000,
    };
    
    const returnTo = req.session.returnTo || '/';
    delete req.session.oauth;
    delete req.session.returnTo;

    res.redirect(returnTo);
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
    req.session.user = { email: "admin.ops@stellaris.loans", name: "Dev User" };
    return next();
  }
  if (req.session?.user?.email) return next();
  
  if (req.path !== '/' && !req.path.startsWith('/auth/') && !req.path.startsWith('/api/')) {
    req.session.returnTo = req.originalUrl;
  }
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
  res.json(req.session?.user?.email || "admin.ops@stellaris.loans");
});

app.post("/api/getUserSignature", async (req, res) => {
  try {
    const email = req.session?.user?.email || "admin.ops@stellaris.loans";
    const result = await airtable.getUserSignature(email);
    res.json(result);
  } catch (err) {
    console.error("getUserSignature error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/updateUserSignature", async (req, res) => {
  try {
    const email = req.session?.user?.email || "admin.ops@stellaris.loans";
    const [signature] = req.body.args || [];
    const result = await airtable.updateUserSignature(email, signature);
    res.json(result);
  } catch (err) {
    console.error("updateUserSignature error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/updateUserPreference", async (req, res) => {
  try {
    const email = req.session?.user?.email || "admin.ops@stellaris.loans";
    const [preferenceName, value] = req.body.args || [];
    const result = await airtable.updateUserPreference(email, preferenceName, value);
    res.json(result);
  } catch (err) {
    console.error("updateUserPreference error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/getRecentContacts", async (req, res) => {
  try {
    const statusFilter = req.body.args?.[0] || null;
    const contacts = await airtable.getRecentContacts(statusFilter);
    res.json(contacts);
  } catch (err) {
    console.error("getRecentContacts error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/searchContacts", async (req, res) => {
  try {
    const query = req.body.args?.[0] || "";
    const statusFilter = req.body.args?.[1] || null;
    const contacts = await airtable.searchContacts(query, statusFilter);
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

const LINK_FIELDS = {
  'Opportunities': {
    'Primary Applicant': 'Contacts',
    'Applicants': 'Contacts',
    'Guarantors': 'Contacts'
  }
};

app.post("/api/updateRecord", async (req, res) => {
  try {
    let [table, id, field, value] = req.body.args || [];
    const userEmail = req.session?.user?.email || null;
    const userContext = userEmail ? await airtable.getUserProfileByEmail(userEmail) : null;
    
    // Strip phone numbers for Contacts table phone fields
    const phoneFields = ['Mobile', 'Telephone1', 'Telephone2'];
    if (table === 'Contacts' && phoneFields.includes(field) && typeof value === 'string') {
      value = value.replace(/\D/g, '');
    }
    
    const linkConfig = LINK_FIELDS[table]?.[field];
    if (linkConfig && userContext && Array.isArray(value)) {
      const currentRecord = await airtable.getRecordFromTable(table, id);
      const currentIds = currentRecord?.fields?.[field] || [];
      const newIds = value;
      const addedIds = newIds.filter(x => !currentIds.includes(x));
      const removedIds = currentIds.filter(x => !newIds.includes(x));
      const changedIds = [...new Set([...addedIds, ...removedIds])];
      for (const contactId of changedIds) {
        await airtable.markRecordModified(linkConfig, contactId, userContext);
      }
    }
    
    await airtable.updateRecordInTable(table, id, field, value, userContext);
    res.json({ success: true });
  } catch (err) {
    console.error("updateRecord error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/createRecord", async (req, res) => {
  try {
    const [table, fields] = req.body.args || [];
    const userEmail = req.session?.user?.email || null;
    const userContext = userEmail ? await airtable.getUserProfileByEmail(userEmail) : null;
    if (table === "Contacts") {
      const contact = await airtable.createContact(fields, userContext);
      res.json(contact);
    } else {
      res.json({ success: false });
    }
  } catch (err) {
    console.error("createRecord error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/checkDuplicates", async (req, res) => {
  try {
    const { mobile, email, firstName, lastName } = req.body;
    const duplicates = [];
    
    if (mobile) {
      const cleanMobile = mobile.replace(/\D/g, '');
      if (cleanMobile.length >= 8) {
        const mobileMatches = await airtable.findContactsByField('Mobile', cleanMobile);
        mobileMatches.forEach(c => {
          if (!duplicates.find(d => d.id === c.id)) {
            duplicates.push({ ...c, matchType: 'mobile' });
          }
        });
      }
    }
    
    if (email) {
      const emailMatches = await airtable.findContactsByField('EmailAddress1', email.toLowerCase());
      emailMatches.forEach(c => {
        const existing = duplicates.find(d => d.id === c.id);
        if (existing) {
          existing.matchType = 'mobile+email';
        } else {
          duplicates.push({ ...c, matchType: 'email' });
        }
      });
    }
    
    if (firstName && lastName) {
      const nameMatches = await airtable.findContactsByName(firstName, lastName);
      nameMatches.forEach(c => {
        const existing = duplicates.find(d => d.id === c.id);
        if (existing) {
          existing.matchType = existing.matchType + '+name';
        } else {
          duplicates.push({ ...c, matchType: 'name' });
        }
      });
    }
    
    res.json({ duplicates });
  } catch (err) {
    console.error("checkDuplicates error:", err);
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

app.post("/api/createOpportunity", async (req, res) => {
  try {
    const [name, contactId, opportunityType, tacoFields] = req.body.args || [];
    const userEmail = req.session?.user?.email || null;
    const userContext = userEmail ? await airtable.getUserProfileByEmail(userEmail) : null;
    const record = await airtable.createOpportunity(name, contactId, opportunityType || "Home Loans", userContext, tacoFields || {});
    if (userContext && contactId) {
      await airtable.markRecordModified("Contacts", contactId, userContext);
    }
    
    // If Converted to Appt is true, also create an appointment record
    if (record && record.id && tacoFields && tacoFields["Taco: Converted to Appt"] === true) {
      try {
        // Parse Taco appointment time to ISO format - skip appointment creation if unparseable
        const parsedTime = parseTacoAppointmentTime(tacoFields["Taco: Appointment Time"]);
        
        if (!parsedTime && tacoFields["Taco: Appointment Time"]) {
          console.warn("Skipping appointment creation - could not parse time:", tacoFields["Taco: Appointment Time"]);
        } else {
          const apptFields = {
            "Appointment Time": parsedTime,
            "Type of Appointment": tacoFields["Taco: Type of Appointment"] || "Phone",
            "How Booked": tacoFields["Taco: How appt booked"] || "Calendly",
            "How Booked Other": tacoFields["Taco: How Appt Booked Other"] || null,
            "Phone Number": tacoFields["Taco: Appt Phone Number"] || null,
            "Video Meet URL": tacoFields["Taco: Appt Meet URL"] || null,
            "Need Evidence in Advance": tacoFields["Taco: Need Evidence in Advance"] === true,
            "Need Appt Reminder": tacoFields["Taco: Need Appt Reminder"] === true,
            "Notes": tacoFields["Taco: Appt Notes"] || null,
            "Appointment Status": parsedTime ? "Scheduled" : null
          };
          await airtable.createAppointment(record.id, apptFields, userContext);
          console.log("Created appointment record for new opportunity:", record.id, "with parsed time:", parsedTime);
        }
      } catch (apptErr) {
        console.error("Failed to create appointment for new opportunity:", apptErr.message);
        // Don't fail the whole request, just log the error
      }
    }
    
    res.json(record);
  } catch (err) {
    console.error("createOpportunity error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/parseTacoData", async (req, res) => {
  try {
    const [rawText] = req.body.args || [];
    const result = { parsed: {}, display: [], unmapped: [] };
    
    if (!rawText || typeof rawText !== 'string') {
      return res.json(result);
    }
    
    const knownKeys = new Set(Object.keys(TACO_FIELD_MAP));
    const lines = rawText.split('\n');
    let consecutiveUnknown = 0;
    
    for (const line of lines) {
      const colonIndex = line.indexOf(':');
      if (colonIndex === -1) {
        consecutiveUnknown++;
        if (consecutiveUnknown >= 2) break;
        continue;
      }
      
      const key = line.substring(0, colonIndex).trim().toLowerCase().replace(/\s+/g, '_').replace(/[^a-z0-9_]/g, '');
      let value = line.substring(colonIndex + 1).trim();
      
      if (!key) {
        consecutiveUnknown++;
        if (consecutiveUnknown >= 2) break;
        continue;
      }
      
      // Taco returns [field_name] when field has no data - treat as empty
      if (value.match(/^\[.+\]$/)) {
        value = '';
      }
      
      const mapping = TACO_FIELD_MAP[key];
      if (mapping) {
        consecutiveUnknown = 0;
        if (typeof mapping === 'object' && mapping.type === 'boolean_flag') {
          // For boolean flags like new_client/existing_client:
          // Only set if there's actual data (not empty after stripping [placeholder])
          if (value) {
            result.parsed[mapping.field] = mapping.value;
            result.display.push({ tacoField: key, airtableField: mapping.field, value: mapping.value });
          }
          // If empty (was [field_name] placeholder), skip - no selection made
        } else if (typeof mapping === 'object' && mapping.type === 'checkbox') {
          // Checkbox fields: empty = unchecked, any value = checked
          const boolVal = value !== '';
          result.parsed[mapping.field] = boolVal;
          result.display.push({ tacoField: key, airtableField: mapping.field, value: boolVal ? 'Yes' : 'No' });
        } else {
          if (value) {
            result.parsed[mapping] = value;
            result.display.push({ tacoField: key, airtableField: mapping, value: value });
          }
        }
      } else {
        consecutiveUnknown++;
        if (consecutiveUnknown >= 2) break;
      }
    }
    
    // Auto-fill logic
    // If Existing Client and Lead Source is empty, set to "Repeat Client"
    if (result.parsed['Taco: New or Existing Client'] === 'Existing Client' && !result.parsed['Taco: Lead Source']) {
      result.parsed['Taco: Lead Source'] = 'Repeat Client';
      result.display.push({ tacoField: '(auto)', airtableField: 'Taco: Lead Source', value: 'Repeat Client' });
    }
    
    // If Broker Assistant is blank, set to "Stephanie Gooch"
    if (!result.parsed['Taco: Broker Assistant']) {
      result.parsed['Taco: Broker Assistant'] = 'Stephanie Gooch';
      result.display.push({ tacoField: '(auto)', airtableField: 'Taco: Broker Assistant', value: 'Stephanie Gooch' });
    }
    
    // If Client Manager is blank, set to "Stephanie Gooch"
    if (!result.parsed['Taco: Client Manager']) {
      result.parsed['Taco: Client Manager'] = 'Stephanie Gooch';
      result.display.push({ tacoField: '(auto)', airtableField: 'Taco: Client Manager', value: 'Stephanie Gooch' });
    }
    
    res.json(result);
  } catch (err) {
    console.error("parseTacoData error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/updateOpportunity", async (req, res) => {
  try {
    const [id, field, value] = req.body.args || [];
    const userEmail = req.session?.user?.email || null;
    const userContext = userEmail ? await airtable.getUserProfileByEmail(userEmail) : null;
    const record = await airtable.updateOpportunity(id, field, value, userContext);
    res.json(record);
  } catch (err) {
    console.error("updateOpportunity error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/deleteContact", async (req, res) => {
  try {
    const [contactId] = req.body.args || [];
    const result = await airtable.deleteContact(contactId);
    res.json(result);
  } catch (err) {
    console.error("deleteContact error:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.post("/api/deleteOpportunity", async (req, res) => {
  try {
    const [opportunityId] = req.body.args || [];
    const result = await airtable.deleteOpportunity(opportunityId);
    res.json(result);
  } catch (err) {
    console.error("deleteOpportunity error:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// ==================== CONTACT ACTIONS ====================

app.post("/api/markContactDeceased", async (req, res) => {
  try {
    const [contactId, isDeceased] = req.body.args || [];
    const userEmail = req.session?.user?.email || null;
    const userContext = userEmail ? await airtable.getUserProfileByEmail(userEmail) : null;
    
    // Update Deceased field (and unsubscribe from marketing if marking as deceased)
    const updateFields = {
      'Deceased': isDeceased
    };
    if (isDeceased) {
      updateFields['Unsubscribed from Marketing'] = true;
    }
    
    const record = await airtable.updateContactMultipleFields(contactId, updateFields, userContext);
    if (!record) {
      return res.status(500).json({ success: false, error: 'Failed to update contact' });
    }
    res.json({ success: true, record });
  } catch (err) {
    console.error("markContactDeceased error:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// ==================== CONNECTIONS ====================

app.post("/api/getConnectionsForContact", async (req, res) => {
  try {
    const [contactId] = req.body.args || [];
    const connections = await airtable.getConnectionsForContact(contactId);
    res.json(connections);
  } catch (err) {
    console.error("getConnectionsForContact error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/getConnectionRoleTypes", async (req, res) => {
  try {
    const roleTypes = airtable.getConnectionRoleTypes();
    res.json(roleTypes);
  } catch (err) {
    console.error("getConnectionRoleTypes error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/createConnection", async (req, res) => {
  try {
    const [contact1Id, contact2Id, record1Role, record2Role] = req.body.args || [];
    const userEmail = req.session?.user?.email || null;
    const userContext = userEmail ? await airtable.getUserProfileByEmail(userEmail) : null;
    const result = await airtable.createConnection(contact1Id, contact2Id, record1Role, record2Role, userContext);
    res.json(result);
  } catch (err) {
    console.error("createConnection error:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.post("/api/deactivateConnection", async (req, res) => {
  try {
    const [connectionId] = req.body.args || [];
    const userEmail = req.session?.user?.email || null;
    const userContext = userEmail ? await airtable.getUserProfileByEmail(userEmail) : null;
    const result = await airtable.deactivateConnection(connectionId, userContext);
    res.json(result);
  } catch (err) {
    console.error("deactivateConnection error:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.post("/api/updateConnectionNote", async (req, res) => {
  try {
    const [connectionId, note] = req.body.args || [];
    const userEmail = req.session?.user?.email || null;
    const userContext = userEmail ? await airtable.getUserProfileByEmail(userEmail) : null;
    const result = await airtable.updateConnectionNote(connectionId, note, userContext);
    res.json(result);
  } catch (err) {
    console.error("updateConnectionNote error:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.post("/api/processForm", async (req, res) => {
  try {
    const formData = req.body.args?.[0] || {};
    const recordId = formData.recordId;
    const userEmail = req.session?.user?.email || null;
    const userContext = userEmail ? await airtable.getUserProfileByEmail(userEmail) : null;
    
    // Build fields object - text fields can be empty strings, but select fields must be omitted if empty
    const textFields = {
      FirstName: formData.firstName || "",
      MiddleName: formData.middleName || "",
      LastName: formData.lastName || "",
      PreferredName: formData.preferredName || "",
      "Does Not Like Being Called": formData.doesNotLike || "",
      Mobile: formData.mobilePhone || "",
      EmailAddress1: formData.email1 || "",
      EmailAddress1Comment: formData.email1Comment || "",
      EmailAddress2: formData.email2 || "",
      EmailAddress2Comment: formData.email2Comment || "",
      EmailAddress3: formData.email3 || "",
      EmailAddress3Comment: formData.email3Comment || "",
      Notes: formData.notes || "",
      "Gender - Other": formData.genderOther || ""
    };
    
    // Select fields - only include if they have a value (Airtable rejects empty strings for select fields)
    const selectFields = {};
    if (formData.gender) selectFields.Gender = formData.gender;
    
    // Date fields - only include if valid
    const dateFields = {};
    const dob = convertDDMMYYYYtoISO(formData.dateOfBirth);
    if (dob) dateFields["Date of Birth"] = dob;
    
    const fields = { ...textFields, ...selectFields, ...dateFields };
    
    if (recordId) {
      for (const [field, value] of Object.entries(fields)) {
        if (value !== undefined) {
          await airtable.updateContact(recordId, field, value, userContext);
        }
      }
      res.json({ type: "update", message: "Contact updated successfully" });
    } else {
      // Set Status to Active for new contacts
      fields.Status = "Active";
      const newRecord = await airtable.createContact(fields, userContext);
      res.json({ type: "create", record: newRecord });
    }
  } catch (err) {
    console.error("processForm error:", err);
    res.status(500).json({ error: err.message });
  }
});

// Convert DD/MM/YYYY to ISO date format for Airtable
function convertDDMMYYYYtoISO(dateStr) {
  if (!dateStr) return null;
  const match = dateStr.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!match) return null;
  const [, day, month, year] = match;
  return `${year}-${month}-${day}`;
}

// Format timestamp to "HH:MM DD/MM/YYYY" format matching Contact audit display
function formatAuditTimestamp(isoString) {
  if (!isoString) return 'Unknown';
  try {
    const date = new Date(isoString);
    if (isNaN(date.getTime())) return isoString;
    const hours = String(date.getHours()).padStart(2, '0');
    const mins = String(date.getMinutes()).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    return `${hours}:${mins} ${day}/${month}/${year}`;
  } catch (e) {
    return isoString;
  }
}

// Taco field mapping: Taco field name -> Airtable field name
const TACO_FIELD_MAP = {
  'new_client': { field: 'Taco: New or Existing Client', type: 'boolean_flag', value: 'New Client' },
  'existing_client': { field: 'Taco: New or Existing Client', type: 'boolean_flag', value: 'Existing Client' },
  'lead_source': 'Taco: Lead Source',
  'whats_the_last_thing_we_did_for_this_client': 'Taco: Last thing we did',
  'how_can_we_help_this_client': 'Taco: How can we help',
  'cm_notes_for_broker': 'Taco: CM notes',
  'broker': 'Taco: Broker',
  'broker_assistant': 'Taco: Broker Assistant',
  'client_manager': 'Taco: Client Manager',
  'converted_to_appointment': { field: 'Taco: Converted to Appt', type: 'checkbox' },
  'appointment_time': 'Taco: Appointment Time',
  'type_of_appointment': 'Taco: Type of Appointment',
  'appoint_phone_number': 'Taco: Appt Phone Number',
  'appointment_meet_url': 'Taco: Appt Meet URL',
  'how_was_appoint_booked': 'Taco: How appt booked',
  'how_was_appoint_booked_other': 'Taco: How Appt Booked Other',
  'need_evidence_in_advance': { field: 'Taco: Need Evidence in Advance', type: 'checkbox' },
  'need_app_reminder': { field: 'Taco: Need Appt Reminder', type: 'checkbox' },
  'appointment_confirmation_email_sent': { field: 'Taco: Appt Conf Email Sent', type: 'checkbox' },
  'appointment_confirmation_text_sent': { field: 'Taco: Appt Conf Text Sent', type: 'checkbox' }
};

const SCHEMA = {
  'Opportunities': {
    auditFields: ['Created', 'Modified'],
    fields: [
      { key: 'Opportunity Name', label: 'Opportunity Name' },
      { key: 'Status', label: 'Status', type: 'select', options: ['Won', 'Open', 'Lost'] },
      { key: 'Opportunity Type', label: 'Opportunity Type', type: 'select', options: ['Home Loans', 'Commercial Loans', 'Deposit Bonds', 'Insurance (General)', 'Insurance (Life)', 'Personal Loans', 'Asset Finance', 'Tax Depreciation Schedule'] },
      // Taco import fields - in order as specified
      { key: 'Taco: New or Existing Client', label: 'New or Existing Client', type: 'select', options: ['', 'New Client', 'Existing Client'], tacoField: true, tacoRow: 1 },
      { key: 'Taco: Lead Source', label: 'Lead Source', type: 'select', options: ['', 'Repeat Client', 'Client Referral', 'Business Referral', 'Agent Referral', 'Walk In', 'Internet', 'Marketing', 'Head Office'], tacoField: true, tacoRow: 1 },
      { key: 'Taco: Last thing we did', label: 'Last thing we did', type: 'long-text', tacoField: true, tacoRow: 2 },
      { key: 'Taco: How can we help', label: 'How can we help', type: 'long-text', tacoField: true, tacoRow: 2 },
      { key: 'Taco: CM notes', label: 'CM notes', type: 'long-text', tacoField: true, tacoRow: 2 },
      { key: 'Taco: Broker', label: 'Broker', tacoField: true, tacoRow: 3 },
      { key: 'Taco: Broker Assistant', label: 'Broker Assistant', tacoField: true, tacoRow: 3 },
      { key: 'Taco: Client Manager', label: 'Client Manager', tacoField: true, tacoRow: 3 },
      { key: 'Taco: Converted to Appt', label: 'Converted to Appt', type: 'checkbox', tacoField: true, tacoRow: 4 },
      { key: 'Taco: Appointment Time', label: 'Appointment Time', tacoField: true, tacoRow: 5, requiresAppt: true },
      { key: 'Taco: Type of Appointment', label: 'Type of Appointment', type: 'select', options: ['', 'Office', 'Phone', 'Video'], tacoField: true, tacoRow: 5, requiresAppt: true },
      { key: 'Taco: How appt booked', label: 'How Appt Booked', type: 'select', options: ['', 'Calendly', 'Email', 'Phone', 'Podium', 'Other'], tacoField: true, tacoRow: 5, requiresAppt: true },
      { key: 'Taco: Appt Phone Number', label: 'Appt Phone Number', tacoField: true, tacoRow: 6, requiresAppt: true, showIf: { field: 'Taco: Type of Appointment', value: 'Phone' } },
      { key: 'Taco: Appt Meet URL', label: 'Appt Meet URL', type: 'url', tacoField: true, tacoRow: 6, requiresAppt: true, showIf: { field: 'Taco: Type of Appointment', value: 'Video' } },
      { key: 'Taco: How Appt Booked Other', label: 'How Appt Booked Other', tacoField: true, tacoRow: 6, requiresAppt: true, showIf: { field: 'Taco: How appt booked', value: 'Other' } },
      { key: 'Taco: Need Evidence in Advance', label: 'Need Evidence in Advance', type: 'checkbox', tacoField: true, tacoRow: 7, requiresAppt: true },
      { key: 'Taco: Need Appt Reminder', label: 'Need Appt Reminder', type: 'checkbox', tacoField: true, tacoRow: 7, requiresAppt: true },
      { key: 'Taco: Appt Conf Email Sent', label: 'Appt Conf Email Sent', type: 'checkbox', tacoField: true, tacoRow: 8, requiresAppt: true },
      { key: 'Taco: Appt Conf Text Sent', label: 'Appt Conf Text Sent', type: 'checkbox', tacoField: true, tacoRow: 8, requiresAppt: true },
      // Other fields
      { key: 'Lead Source Major', label: 'Lead Source Major', type: 'readonly' },
      { key: 'Lead Source Minor', label: 'Lead Source Minor', type: 'readonly' },
      { key: 'Primary Applicant', nameKey: 'Primary Applicant Name', table: 'Contacts', label: 'Primary Applicant' },
      { key: 'Applicants', nameKey: 'Applicants Name', table: 'Contacts', label: 'Applicants' },
      { key: 'Guarantors', nameKey: 'Guarantors Name', table: 'Contacts', label: 'Guarantors' },
      { key: 'Loan Applications', nameKey: 'Loan Applications Name', table: 'Loan Applications', label: 'Loan Applications' },
      { key: 'CustomerIdName', label: 'Customer ID Name' },
      { key: 'mc_loanconsultantid', label: 'Loan Consultant ID', type: 'long-text' },
      { key: 'mc_LoanProcessorId', label: 'Loan Processor ID', type: 'long-text' },
      { key: 'Description', label: 'Description', type: 'long-text' },
      { key: 'mc_AmountRequested', label: 'Amount Requested', type: 'long-text' },
      { key: 'mc_amountrequested_Base', label: 'Amount Requested (Base)', type: 'long-text' },
      { key: 'Estimated Value', label: 'Estimated Value', type: 'long-text' },
      { key: 'Actual Value', label: 'Actual Value', type: 'long-text' },
      { key: 'ActualValue_Base', label: 'Actual Value (Base)', type: 'long-text' },
      { key: 'EstimatedValue_Base', label: 'Estimated Value (Base)', type: 'long-text' },
      { key: 'Tasks', nameKey: 'Tasks Name', table: 'Tasks', label: 'Tasks' },
      { key: 'Submitted Date', label: 'Submitted Date', type: 'date' },
      { key: 'mc_DateUnconditionalOpp', label: 'Date Unconditional', type: 'date' },
      { key: 'mc_DateSettlement', label: 'Date Settlement', type: 'date' },
      { key: 'mc_DateSettlementOpp', label: 'Date Settlement (Opp)', type: 'date' },
      { key: 'mc_DateDeclinedOpp', label: 'Date Declined', type: 'date' },
      { key: 'mc_DateWithdrawnOpp', label: 'Date Withdrawn', type: 'date' },
      { key: 'ActualCloseDate', label: 'Actual Close Date', type: 'date' },
      { key: 'mc_IsReferral', label: 'Is Referral', type: 'long-text' },
      { key: 'mc_ReferralNotes', label: 'Referral Notes', type: 'long-text' },
      { key: 'mc_SentReferral', label: 'Sent Referral', type: 'long-text' },
      { key: 'mc_referralsenton', label: 'Referral Sent On', type: 'date' },
      { key: 'mc_referralstatus', label: 'Referral Status', type: 'long-text' },
      { key: 'mc_franchiseleadsourceid', label: 'Franchise Lead Source ID', type: 'long-text' },
      { key: 'mc_FinancialAdviserId', label: 'Financial Adviser ID', type: 'long-text' },
      { key: 'mc_SupplierId', label: 'Supplier ID', type: 'long-text' },
      { key: 'mc_RelatedOpportunityId', label: 'Related Opportunity ID', type: 'long-text' }
    ]
  },
  'Loan Applications': {
    fields: [
      { key: 'Name', label: 'Application Name' },
      { key: 'Status', label: 'Status' },
      { key: 'Lender', nameKey: 'Lender Name', table: 'Lenders', label: 'Lender' },
      { key: 'Opportunity', nameKey: 'Opportunity Name', table: 'Opportunities', label: 'Linked Opportunity' }
    ]
  },
  'Tasks': {
    fields: [
      { key: 'Name', label: 'Task Name' },
      { key: 'Status', label: 'Status' },
      { key: 'Due Date', label: 'Due Date' },
      { key: 'Opportunity', nameKey: 'Opportunity Name', table: 'Opportunities', label: 'Linked Opportunity' }
    ]
  },
  'Contacts': {
    fields: [
      { key: 'FirstName', label: 'First Name' },
      { key: 'LastName', label: 'Last Name' },
      { key: 'Mobile', label: 'Mobile' },
      { key: 'EmailAddress1', label: 'Email' }
    ]
  },
  'Lenders': {
    fields: [
      { key: 'Name', label: 'Lender Name' },
      { key: 'BDM Name', label: 'BDM' },
      { key: 'Phone', label: 'Phone' }
    ]
  }
};

app.post("/api/getRecordDetail", async (req, res) => {
  try {
    const [tableName, id] = req.body.args || [];
    
    if (!SCHEMA[tableName]) {
      return res.json({ title: "Unknown Record", data: [{ label: "Error", value: `Table '${tableName}' not configured.` }] });
    }
    
    const schemaDef = SCHEMA[tableName];
    const record = await airtable.getRecordFromTable(tableName, id);
    if (!record) return res.json({ data: [], title: "Not Found" });
    
    const rawFields = record.fields;
    const processedData = [];
    
    schemaDef.fields.forEach(fieldDef => {
      const val = rawFields[fieldDef.key];
      
      if (fieldDef.table && fieldDef.nameKey) {
        const ids = Array.isArray(val) ? val : [];
        const names = rawFields[fieldDef.nameKey] || [];
        const links = ids.map((recId, index) => {
          let name = "Unknown";
          if (Array.isArray(names)) name = names[index] || "Unknown";
          else if (index === 0) name = names;
          return { id: recId, name: name, table: fieldDef.table };
        });
        processedData.push({ key: fieldDef.key, label: fieldDef.label, value: links, type: 'link' });
      } else {
        const dataItem = { key: fieldDef.key, label: fieldDef.label, value: val, type: fieldDef.type || 'text' };
        if (fieldDef.options) dataItem.options = fieldDef.options;
        if (fieldDef.tacoField) dataItem.tacoField = true;
        processedData.push(dataItem);
      }
    });
    
    const auditInfo = {};
    if (schemaDef.auditFields) {
      // Use the formula fields 'Created' and 'Modified' which already contain formatted timestamps and user names
      if (schemaDef.auditFields.includes('Created') && rawFields['Created']) {
        auditInfo.Created = rawFields['Created'];
      }
      if (schemaDef.auditFields.includes('Modified') && rawFields['Modified']) {
        auditInfo.Modified = rawFields['Modified'];
      }
    }
    
    res.json({ 
      title: rawFields[schemaDef.fields[0].key] || "Details", 
      data: processedData,
      audit: auditInfo
    });
  } catch (err) {
    console.error("getRecordDetail error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/getRecordById", async (req, res) => {
  try {
    const [tableName, id] = req.body.args || [];
    const record = await airtable.getRecordFromTable(tableName, id);
    if (!record) return res.json(null);
    res.json(record);
  } catch (err) {
    console.error("getRecordById error:", err);
    res.status(500).json({ error: err.message });
  }
});

// --- SETTINGS API ENDPOINTS ---
app.post("/api/getSetting", async (req, res) => {
  try {
    const [key] = req.body.args || [];
    const value = await airtable.getSetting(key);
    res.json(value);
  } catch (err) {
    console.error("getSetting error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/getAllSettings", async (req, res) => {
  try {
    const settings = await airtable.getAllSettings();
    res.json(settings);
  } catch (err) {
    console.error("getAllSettings error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/updateSetting", async (req, res) => {
  try {
    const [key, value] = req.body.args || [];
    const userEmail = req.session?.user?.email || "admin.ops@stellaris.loans";
    const result = await airtable.updateSetting(key, value, userEmail);
    res.json(result);
  } catch (err) {
    console.error("updateSetting error:", err);
    res.status(500).json({ error: err.message });
  }
});

// --- EMAIL TEMPLATES API ENDPOINTS ---
app.post("/api/getEmailTemplates", async (req, res) => {
  try {
    const templates = await airtable.getEmailTemplates();
    res.json(templates);
  } catch (err) {
    console.error("getEmailTemplates error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/getEmailTemplate", async (req, res) => {
  try {
    const [templateId] = req.body.args || [];
    const template = await airtable.getEmailTemplate(templateId);
    res.json(template);
  } catch (err) {
    console.error("getEmailTemplate error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/updateEmailTemplate", async (req, res) => {
  try {
    const [templateId, fields] = req.body.args || [];
    const userEmail = req.session?.user?.email || null;
    const userContext = userEmail ? await airtable.getUserProfileByEmail(userEmail) : null;
    const result = await airtable.updateEmailTemplate(templateId, fields, userContext);
    res.json(result);
  } catch (err) {
    console.error("updateEmailTemplate error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/createEmailTemplate", async (req, res) => {
  try {
    const [fields] = req.body.args || [];
    const userEmail = req.session?.user?.email || null;
    const userContext = userEmail ? await airtable.getUserProfileByEmail(userEmail) : null;
    const result = await airtable.createEmailTemplate(fields, userContext);
    res.json(result);
  } catch (err) {
    console.error("createEmailTemplate error:", err);
    res.status(500).json({ error: err.message });
  }
});

// --- APPOINTMENTS API ENDPOINTS ---
app.post("/api/getAppointmentsForOpportunity", async (req, res) => {
  try {
    const [opportunityId] = req.body.args || [];
    const appointments = await airtable.getAppointmentsForOpportunity(opportunityId);
    res.json(appointments);
  } catch (err) {
    console.error("getAppointmentsForOpportunity error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/createAppointment", async (req, res) => {
  try {
    const [opportunityId, fields] = req.body.args || [];
    const userEmail = req.session?.user?.email || null;
    const userContext = userEmail ? await airtable.getUserProfileByEmail(userEmail) : null;
    
    // Parse Taco-format appointment time if present - fail if unparseable
    if (fields && fields["Appointment Time"]) {
      const parsedTime = parseTacoAppointmentTime(fields["Appointment Time"]);
      if (!parsedTime) {
        return res.status(400).json({ error: `Could not parse appointment time: "${fields["Appointment Time"]}". Please use format DD/MM/YY H:MM AM/PM or enter time manually.` });
      }
      fields["Appointment Time"] = parsedTime;
    }
    
    const result = await airtable.createAppointment(opportunityId, fields, userContext);
    res.json(result);
  } catch (err) {
    console.error("createAppointment error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/updateAppointment", async (req, res) => {
  try {
    const [appointmentId, field, value] = req.body.args || [];
    const userEmail = req.session?.user?.email || null;
    const userContext = userEmail ? await airtable.getUserProfileByEmail(userEmail) : null;
    const result = await airtable.updateAppointment(appointmentId, field, value, userContext);
    res.json(result);
  } catch (err) {
    console.error("updateAppointment error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/updateAppointmentFields", async (req, res) => {
  try {
    const [appointmentId, fields] = req.body.args || [];
    const userEmail = req.session?.user?.email || null;
    const userContext = userEmail ? await airtable.getUserProfileByEmail(userEmail) : null;
    const result = await airtable.updateAppointmentFields(appointmentId, fields, userContext);
    res.json(result);
  } catch (err) {
    console.error("updateAppointmentFields error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/deleteAppointment", async (req, res) => {
  try {
    const [appointmentId] = req.body.args || [];
    const result = await airtable.deleteAppointment(appointmentId);
    res.json({ success: result });
  } catch (err) {
    console.error("deleteAppointment error:", err);
    res.status(500).json({ error: err.message });
  }
});

// --- ADDRESS HISTORY ---
app.post("/api/getAddressesForContact", async (req, res) => {
  try {
    const [contactId] = req.body.args || [];
    const addresses = await airtable.getAddressesForContact(contactId);
    res.json(addresses);
  } catch (err) {
    console.error("getAddressesForContact error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/getAddressById", async (req, res) => {
  try {
    const [addressId] = req.body.args || [];
    const address = await airtable.getAddressById(addressId);
    res.json(address);
  } catch (err) {
    console.error("getAddressById error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/createAddress", async (req, res) => {
  try {
    const [contactId, fields] = req.body.args || [];
    const userEmail = req.session?.user?.email || null;
    const userContext = userEmail ? await airtable.getUserProfileByEmail(userEmail) : null;
    const result = await airtable.createAddress(contactId, fields, userContext);
    res.json(result);
  } catch (err) {
    console.error("createAddress error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/updateAddress", async (req, res) => {
  try {
    const [addressId, fields] = req.body.args || [];
    const userEmail = req.session?.user?.email || null;
    const userContext = userEmail ? await airtable.getUserProfileByEmail(userEmail) : null;
    const result = await airtable.updateAddress(addressId, fields, userContext);
    res.json(result);
  } catch (err) {
    console.error("updateAddress error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/deleteAddress", async (req, res) => {
  try {
    const [addressId] = req.body.args || [];
    const userEmail = req.session?.user?.email || null;
    const userContext = userEmail ? await airtable.getUserProfileByEmail(userEmail) : null;
    const result = await airtable.deleteAddress(addressId, userContext);
    res.json(result);
  } catch (err) {
    console.error("deleteAddress error:", err);
    res.status(500).json({ error: err.message });
  }
});

// --- EMPLOYMENT HISTORY ---
app.post("/api/getEmploymentForContact", async (req, res) => {
  try {
    const [contactId] = req.body.args || [];
    const employment = await airtable.getEmploymentForContact(contactId);
    res.json(employment);
  } catch (err) {
    console.error("getEmploymentForContact error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/createEmployment", async (req, res) => {
  try {
    const [contactId, fields] = req.body.args || [];
    const userEmail = req.session?.user?.email || null;
    const userContext = userEmail ? await airtable.getUserProfileByEmail(userEmail) : null;
    const result = await airtable.createEmployment(contactId, fields, userContext);
    res.json(result);
  } catch (err) {
    console.error("createEmployment error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/updateEmployment", async (req, res) => {
  try {
    const [employmentId, fields] = req.body.args || [];
    const userEmail = req.session?.user?.email || null;
    const userContext = userEmail ? await airtable.getUserProfileByEmail(userEmail) : null;
    const result = await airtable.updateEmployment(employmentId, fields, userContext);
    res.json(result);
  } catch (err) {
    console.error("updateEmployment error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/deleteEmployment", async (req, res) => {
  try {
    const [employmentId] = req.body.args || [];
    const userEmail = req.session?.user?.email || null;
    const userContext = userEmail ? await airtable.getUserProfileByEmail(userEmail) : null;
    const result = await airtable.deleteEmployment(employmentId, userContext);
    res.json(result);
  } catch (err) {
    console.error("deleteEmployment error:", err);
    res.status(500).json({ error: err.message });
  }
});

// --- EMAIL SENDING via Gmail API ---
app.post("/api/sendEmail", async (req, res) => {
  try {
    const [to, subject, body] = req.body.args || [];
    if (!to || !subject || !body) {
      return res.json({ success: false, error: "Missing required fields: to, subject, or body" });
    }
    
    // Get Gmail tokens from session (production) or use Replit connector (development)
    const gmailTokens = req.session?.gmailTokens;
    const baseUrl = getBaseUrl(req);
    
    // Pass tokens and OAuth client config for token refresh
    const result = await gmail.sendEmail(to, subject, body, {
      tokens: gmailTokens,
      clientId: GOOGLE_CLIENT_ID,
      clientSecret: GOOGLE_CLIENT_SECRET,
      onTokenRefresh: (newTokens) => {
        // Update session with refreshed tokens
        if (req.session) {
          req.session.gmailTokens = newTokens;
        }
      }
    });
    
    res.json(result);
  } catch (err) {
    console.error("sendEmail error:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// --- EVIDENCE SYSTEM ---
app.post("/api/getEvidenceCategories", async (req, res) => {
  try {
    const categories = await airtable.getEvidenceCategories();
    res.json(categories);
  } catch (err) {
    console.error("getEvidenceCategories error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/getEvidenceTemplates", async (req, res) => {
  try {
    const [opportunityType, lender] = req.body.args || [];
    const templates = await airtable.getEvidenceTemplates(opportunityType, lender);
    res.json(templates);
  } catch (err) {
    console.error("getEvidenceTemplates error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/getEvidenceItemsForOpportunity", async (req, res) => {
  try {
    const [opportunityId] = req.body.args || [];
    const items = await airtable.getEvidenceItemsForOpportunity(opportunityId);
    res.json(items);
  } catch (err) {
    console.error("getEvidenceItemsForOpportunity error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/populateEvidenceForOpportunity", async (req, res) => {
  try {
    const [opportunityId, opportunityType, lender] = req.body.args || [];
    const userEmail = req.session?.user?.email || null;
    const userContext = userEmail ? await airtable.getUserProfileByEmail(userEmail) : null;
    const result = await airtable.populateEvidenceForOpportunity(opportunityId, opportunityType, lender, userContext);
    res.json(result);
  } catch (err) {
    console.error("populateEvidenceForOpportunity error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/updateEvidenceItem", async (req, res) => {
  try {
    const [itemId, fields] = req.body.args || [];
    const userEmail = req.session?.user?.email || null;
    const userContext = userEmail ? await airtable.getUserProfileByEmail(userEmail) : null;
    const result = await airtable.updateEvidenceItem(itemId, fields, userContext);
    res.json(result);
  } catch (err) {
    console.error("updateEvidenceItem error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/createEvidenceItem", async (req, res) => {
  try {
    const [opportunityId, fields] = req.body.args || [];
    const userEmail = req.session?.user?.email || null;
    const userContext = userEmail ? await airtable.getUserProfileByEmail(userEmail) : null;
    const result = await airtable.createEvidenceItem(opportunityId, fields, userContext);
    res.json(result);
  } catch (err) {
    console.error("createEvidenceItem error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/deleteEvidenceItem", async (req, res) => {
  try {
    const [itemId] = req.body.args || [];
    const result = await airtable.deleteEvidenceItem(itemId);
    res.json(result);
  } catch (err) {
    console.error("deleteEvidenceItem error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/markEvidenceItemsAsRequested", async (req, res) => {
  try {
    const [itemIds] = req.body.args || [];
    const userEmail = req.session?.user?.email || null;
    const userContext = userEmail ? await airtable.getUserProfileByEmail(userEmail) : null;
    const result = await airtable.markEvidenceItemsAsRequested(itemIds, userContext);
    res.json(result);
  } catch (err) {
    console.error("markEvidenceItemsAsRequested error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/updateEvidenceItemsOrder", async (req, res) => {
  try {
    const [items] = req.body.args || [];
    const result = await airtable.updateEvidenceItemsOrder(items);
    res.json(result);
  } catch (err) {
    console.error("updateEvidenceItemsOrder error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/getAllEvidenceTemplates", async (req, res) => {
  try {
    const templates = await airtable.getAllEvidenceTemplates();
    res.json(templates);
  } catch (err) {
    console.error("getAllEvidenceTemplates error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/createEvidenceTemplate", async (req, res) => {
  try {
    const [fields] = req.body.args || [];
    const userEmail = req.session?.user?.email || null;
    const userContext = userEmail ? await airtable.getUserProfileByEmail(userEmail) : null;
    const result = await airtable.createEvidenceTemplate(fields, userContext);
    res.json(result);
  } catch (err) {
    console.error("createEvidenceTemplate error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/updateEvidenceTemplate", async (req, res) => {
  try {
    const [templateId, fields] = req.body.args || [];
    const result = await airtable.updateEvidenceTemplate(templateId, fields);
    res.json(result);
  } catch (err) {
    console.error("updateEvidenceTemplate error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/deleteEvidenceTemplate", async (req, res) => {
  try {
    const [templateId] = req.body.args || [];
    const result = await airtable.deleteEvidenceTemplate(templateId);
    res.json(result);
  } catch (err) {
    console.error("deleteEvidenceTemplate error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/getAllContactsForExport", async (req, res) => {
  try {
    const contacts = await airtable.getAllContactsForExport();
    res.json(contacts);
  } catch (err) {
    console.error("getAllContactsForExport error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/getCampaigns", async (req, res) => {
  try {
    const campaigns = await airtable.getCampaigns();
    res.json(campaigns);
  } catch (err) {
    console.error("getCampaigns error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/getCampaignStats", async (req, res) => {
  try {
    const stats = await airtable.getCampaignStats();
    res.json(stats);
  } catch (err) {
    console.error("getCampaignStats error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/getCampaignLogs", async (req, res) => {
  try {
    const campaignId = req.body.args ? req.body.args[0] : req.body.campaignId;
    if (!campaignId) {
      return res.status(400).json({ error: "campaignId is required." });
    }
    const logs = await airtable.getCampaignLogs(campaignId);
    res.json(logs);
  } catch (err) {
    console.error("getCampaignLogs error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/getMarketingLogsForContact", async (req, res) => {
  try {
    const contactId = req.body.args ? req.body.args[0] : req.body.contactId;
    if (!contactId) {
      return res.status(400).json({ error: "contactId is required." });
    }
    const logs = await airtable.getMarketingLogsForContact(contactId);
    res.json(logs);
  } catch (err) {
    console.error("getMarketingLogsForContact error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/importCampaignResults", async (req, res) => {
  try {
    const payload = req.body.args ? req.body.args[0] : req.body;
    const { campaignId, rows } = payload || {};
    if (!campaignId || !rows || !Array.isArray(rows)) {
      return res.status(400).json({ error: "campaignId and rows are required." });
    }
    const result = await airtable.importCampaignResults({ campaignId, rows });
    res.json(result);
  } catch (err) {
    console.error("importCampaignResults error:", err);
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
