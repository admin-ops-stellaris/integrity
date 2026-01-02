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
    req.session.user = { email: "admin.ops@stellaris.loans", name: "Dev User" };
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

const LINK_FIELDS = {
  'Opportunities': {
    'Primary Applicant': 'Contacts',
    'Applicants': 'Contacts',
    'Guarantors': 'Contacts'
  }
};

app.post("/api/updateRecord", async (req, res) => {
  try {
    const [table, id, field, value] = req.body.args || [];
    const userEmail = req.session?.user?.email || null;
    const userContext = userEmail ? await airtable.getUserProfileByEmail(userEmail) : null;
    
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

app.post("/api/processForm", async (req, res) => {
  try {
    const formData = req.body.args?.[0] || {};
    const recordId = formData.recordId;
    const userEmail = req.session?.user?.email || null;
    const userContext = userEmail ? await airtable.getUserProfileByEmail(userEmail) : null;
    
    const fields = {
      FirstName: formData.firstName || "",
      MiddleName: formData.middleName || "",
      LastName: formData.lastName || "",
      PreferredName: formData.preferredName || "",
      Mobile: formData.mobilePhone || "",
      EmailAddress1: formData.email1 || "",
      Description: formData.description || ""
    };
    
    if (recordId) {
      for (const [field, value] of Object.entries(fields)) {
        if (value !== undefined) {
          await airtable.updateContact(recordId, field, value, userContext);
        }
      }
      res.json("Contact updated successfully");
    } else {
      await airtable.createContact(fields, userContext);
      res.json("Contact created successfully");
    }
  } catch (err) {
    console.error("processForm error:", err);
    res.status(500).json({ error: err.message });
  }
});

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
      { key: 'Taco: New or Existing Client', label: 'New or Existing Client', tacoField: true },
      { key: 'Taco: Lead Source', label: 'Lead Source', tacoField: true },
      { key: 'Taco: Last thing we did', label: 'Last thing we did', type: 'long-text', tacoField: true },
      { key: 'Taco: How can we help', label: 'How can we help', type: 'long-text', tacoField: true },
      { key: 'Taco: CM notes', label: 'CM notes', type: 'long-text', tacoField: true },
      { key: 'Taco: Broker', label: 'Broker', tacoField: true },
      { key: 'Taco: Broker Assistant', label: 'Broker Assistant', tacoField: true },
      { key: 'Taco: Client Manager', label: 'Client Manager', tacoField: true },
      { key: 'Taco: Converted to Appt', label: 'Converted to Appt', tacoField: true },
      { key: 'Taco: Appointment Time', label: 'Appointment Time', tacoField: true },
      { key: 'Taco: Type of Appointment', label: 'Type of Appointment', tacoField: true },
      { key: 'Taco: Appt Phone Number', label: 'Appt Phone Number', tacoField: true },
      { key: 'Taco: How appt booked', label: 'How appt booked', tacoField: true },
      { key: 'Taco: How Appt Booked Other', label: 'How Appt Booked Other', tacoField: true },
      { key: 'Taco: Need Evidence in Advance', label: 'Need Evidence in Advance', tacoField: true },
      { key: 'Taco: Need Appt Reminder', label: 'Need Appt Reminder', tacoField: true },
      { key: 'Taco: Appt Conf Email Sent', label: 'Appt Conf Email Sent', tacoField: true },
      { key: 'Taco: Appt Conf Text Sent', label: 'Appt Conf Text Sent', tacoField: true },
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
      // Use the formula fields 'Created' and 'Modified' which already contain formatted timestamps
      // Then append "by [Name]" from the user tracking fields
      if (schemaDef.auditFields.includes('Created') && rawFields['Created']) {
        const createdBy = rawFields['Creating Site User Name'] || null;
        auditInfo.Created = rawFields['Created'] + (createdBy ? ' by ' + createdBy : '');
      }
      if (schemaDef.auditFields.includes('Modified') && rawFields['Modified']) {
        const modifiedBy = rawFields['Last Site User Name'] || null;
        auditInfo.Modified = rawFields['Modified'] + (modifiedBy ? ' by ' + modifiedBy : '');
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
