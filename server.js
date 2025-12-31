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
    const [name, contactId, opportunityType] = req.body.args || [];
    const userEmail = req.session?.user?.email || null;
    const userContext = userEmail ? await airtable.getUserProfileByEmail(userEmail) : null;
    const record = await airtable.createOpportunity(name, contactId, opportunityType || "Home Loans", userContext);
    if (userContext && contactId) {
      await airtable.markRecordModified("Contacts", contactId, userContext);
    }
    res.json(record);
  } catch (err) {
    console.error("createOpportunity error:", err);
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

const SCHEMA = {
  'Opportunities': {
    auditFields: ['Created', 'Modified', 'Last Site User Name'],
    fields: [
      { key: 'Opportunity Name', label: 'Opportunity Name' },
      { key: 'Status', label: 'Status', type: 'select', options: ['Won', 'Open', 'Lost'] },
      { key: 'Opportunity Type', label: 'Opportunity Type', type: 'select', options: ['Home Loans', 'Commercial Loans', 'Deposit Bonds', 'Insurance (General)', 'Insurance (Life)', 'Personal Loans', 'Asset Finance', 'Tax Depreciation Schedule'] },
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
        processedData.push(dataItem);
      }
    });
    
    const auditInfo = {};
    if (schemaDef.auditFields) {
      schemaDef.auditFields.forEach(key => {
        auditInfo[key] = rawFields[key] || null;
      });
      // Format audit strings to match Contact format: "Created: HH:MM DD/MM/YYYY by Name"
      if (auditInfo.Created) {
        const createdBy = rawFields['Creating Site User Name'] || null;
        const createdDate = formatAuditTimestamp(rawFields['Created On']);
        auditInfo.Created = `Created: ${createdDate}${createdBy ? ' by ' + createdBy : ''}`;
      }
      if (auditInfo.Modified) {
        const modifiedBy = rawFields['Last Site User Name'] || null;
        const modifiedDate = formatAuditTimestamp(rawFields['Modified On']);
        auditInfo.Modified = `Modified: ${modifiedDate}${modifiedBy ? ' by ' + modifiedBy : ''}`;
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
