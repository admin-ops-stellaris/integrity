import Airtable from "airtable";

const AIRTABLE_API_KEY = process.env.AIRTABLE_API_KEY;
const AIRTABLE_BASE_ID = process.env.AIRTABLE_BASE_ID;

if (!AIRTABLE_API_KEY || !AIRTABLE_BASE_ID) {
  console.warn("Warning: Airtable credentials not configured. Using mock mode.");
}

const base = AIRTABLE_API_KEY && AIRTABLE_BASE_ID 
  ? new Airtable({ apiKey: AIRTABLE_API_KEY }).base(AIRTABLE_BASE_ID)
  : null;

function formatRecord(record) {
  return {
    id: record.id,
    fields: record.fields
  };
}

function formatAppointmentRecord(record) {
  const f = record.fields || {};
  return {
    id: record.id,
    appointmentTime: f["Appointment Time"] || null,
    typeOfAppointment: f["Type of Appointment"] || null,
    howBooked: f["How Booked"] || null,
    howBookedOther: f["How Booked Other"] || null,
    phoneNumber: f["Phone Number"] || null,
    videoMeetUrl: f["Video Meet URL"] || null,
    needEvidenceInAdvance: f["Need Evidence in Advance"] || false,
    needApptReminder: f["Need Appt Reminder"] || false,
    notes: f["Notes"] || null,
    appointmentStatus: f["Appointment Status"] || null,
    opportunityId: Array.isArray(f["Opportunity"]) ? f["Opportunity"][0] : f["Opportunity"]
  };
}

const userProfileCache = new Map();

export async function getUserProfileByEmail(email) {
  if (!base || !email) return null;
  
  const cacheKey = email.toLowerCase();
  if (userProfileCache.has(cacheKey)) {
    return userProfileCache.get(cacheKey);
  }
  
  try {
    const records = await base("Users")
      .select({
        filterByFormula: `LOWER({Email}) = "${cacheKey}"`,
        maxRecords: 1
      })
      .all();
    
    if (records.length > 0) {
      const profile = {
        id: records[0].id,
        name: records[0].fields["Name"] || null,
        email: email,
        title: records[0].fields["Title"] || null,
        signature: records[0].fields["Email Signature"] || null
      };
      userProfileCache.set(cacheKey, profile);
      return profile;
    }
    
    console.warn(`User not found in Users table for email: ${email}`);
    return { name: null, email: email, title: null, signature: null };
  } catch (err) {
    console.error("getUserProfileByEmail error:", err.message);
    return { name: null, email: email, title: null, signature: null };
  }
}

export async function getUserById(userId) {
  if (!base || !userId) return null;
  
  if (userProfileCache.has(userId)) {
    return userProfileCache.get(userId);
  }
  
  try {
    const record = await base("Users").find(userId);
    if (record) {
      const profile = {
        id: record.id,
        name: record.fields["Name"] || null,
        email: record.fields["Email"] || null
      };
      userProfileCache.set(userId, profile);
      return profile;
    }
    return null;
  } catch (err) {
    console.error("getUserById error:", err.message);
    return null;
  }
}

export async function getUserSignature(email) {
  if (!base || !email) return null;
  
  try {
    const records = await base("Users")
      .select({
        filterByFormula: `LOWER({Email}) = "${email.toLowerCase()}"`,
        maxRecords: 1
      })
      .all();
    
    if (records.length > 0) {
      return {
        id: records[0].id,
        name: records[0].fields["Name"] || "",
        title: records[0].fields["Title"] || "",
        signature: records[0].fields["Email Signature"] || ""
      };
    }
    return null;
  } catch (err) {
    console.error("getUserSignature error:", err.message);
    return null;
  }
}

export async function updateUserSignature(email, signature) {
  if (!base || !email) return null;
  
  try {
    const records = await base("Users")
      .select({
        filterByFormula: `LOWER({Email}) = "${email.toLowerCase()}"`,
        maxRecords: 1
      })
      .all();
    
    if (records.length > 0) {
      const updated = await base("Users").update(records[0].id, {
        "Email Signature": signature
      });
      userProfileCache.delete(email.toLowerCase());
      return formatRecord(updated);
    }
    return null;
  } catch (err) {
    console.error("updateUserSignature error:", err.message);
    return null;
  }
}

export async function getContactById(id) {
  if (!base) return null;
  try {
    const record = await base("Contacts").find(id);
    const formatted = formatRecord(record);
    console.log("Contact fields available:", Object.keys(formatted.fields));
    return formatted;
  } catch (err) {
    console.error("getContactById error:", err.message);
    return null;
  }
}

function parseModifiedDate(modifiedText) {
  if (!modifiedText) return new Date(0);
  const match = modifiedText.match(/(\d{2}):(\d{2})\s+(\d{2})\/(\d{2})\/(\d{4})/);
  if (match) {
    const [, hours, mins, day, month, year] = match;
    return new Date(year, month - 1, day, hours, mins);
  }
  return new Date(0);
}

export async function getRecentContacts() {
  if (!base) return [];
  try {
    const records = await base("Contacts")
      .select({
        maxRecords: 50,
        sort: [{ field: "Modified On", direction: "desc" }]
      })
      .all();
    return records.map(formatRecord);
  } catch (err) {
    console.error("getRecentContacts error:", err.message);
    return [];
  }
}

export async function searchContacts(query) {
  if (!base) return [];
  if (!query || query.trim() === "") return getRecentContacts();
  
  try {
    const terms = query.toLowerCase().split(/\s+/).filter(t => t.length > 0);
    const searchConditions = terms.map(term => `SEARCH("${term}", {SEARCH_INDEX})`);
    const formula = terms.length === 1 
      ? searchConditions[0]
      : `AND(${searchConditions.join(", ")})`;
    
    const records = await base("Contacts")
      .select({ filterByFormula: formula })
      .all();
    return records.map(formatRecord);
  } catch (err) {
    console.error("searchContacts error:", err.message);
    return [];
  }
}

export async function updateContact(id, field, value, userContext = null) {
  if (!base) return null;
  try {
    const updateFields = { [field]: value };
    if (userContext) {
      if (userContext.name) updateFields["Last Site User Name"] = userContext.name;
      if (userContext.email) updateFields["Last Site User Email"] = userContext.email;
    }
    const record = await base("Contacts").update(id, updateFields);
    return formatRecord(record);
  } catch (err) {
    console.error("updateContact error:", err.message);
    return null;
  }
}

export async function createContact(fields, userContext = null) {
  if (!base) return null;
  try {
    const createFields = { ...fields };
    if (userContext) {
      if (userContext.name) {
        createFields["Creating Site User Name"] = userContext.name;
        createFields["Last Site User Name"] = userContext.name;
      }
      if (userContext.email) {
        createFields["Creating Site User Email"] = userContext.email;
        createFields["Last Site User Email"] = userContext.email;
      }
    }
    const record = await base("Contacts").create(createFields);
    return formatRecord(record);
  } catch (err) {
    console.error("createContact error:", err.message);
    return null;
  }
}

export async function getOpportunityById(id) {
  if (!base) return null;
  try {
    const record = await base("Opportunities").find(id);
    return formatRecord(record);
  } catch (err) {
    console.error("getOpportunityById error:", err.message);
    return null;
  }
}

export async function getRecordFromTable(tableName, id) {
  if (!base) return null;
  try {
    const record = await base(tableName).find(id);
    return formatRecord(record);
  } catch (err) {
    console.error(`getRecordFromTable(${tableName}) error:`, err.message);
    return null;
  }
}

export async function getOpportunitiesById(ids) {
  if (!base || !ids || ids.length === 0) return [];
  try {
    const formula = `OR(${ids.map(id => `RECORD_ID()="${id}"`).join(",")})`;
    const records = await base("Opportunities")
      .select({ filterByFormula: formula })
      .all();
    return records.map(formatRecord);
  } catch (err) {
    console.error("getOpportunitiesById error:", err.message);
    return [];
  }
}

export async function updateOpportunity(id, field, value, userContext = null) {
  if (!base) return null;
  try {
    const updateFields = { [field]: value };
    if (userContext) {
      if (userContext.name) updateFields["Last Site User Name"] = userContext.name;
      if (userContext.email) updateFields["Last Site User Email"] = userContext.email;
    }
    const record = await base("Opportunities").update(id, updateFields);
    return formatRecord(record);
  } catch (err) {
    console.error("updateOpportunity error:", err.message);
    return null;
  }
}

const TRACKED_TABLES = ["Contacts", "Opportunities"];

export async function updateRecordInTable(tableName, id, field, value, userContext = null) {
  if (!base) return null;
  try {
    const updateFields = { [field]: value };
    if (userContext && TRACKED_TABLES.includes(tableName)) {
      if (userContext.name) updateFields["Last Site User Name"] = userContext.name;
      if (userContext.email) updateFields["Last Site User Email"] = userContext.email;
    }
    const record = await base(tableName).update(id, updateFields);
    return formatRecord(record);
  } catch (err) {
    console.error(`updateRecordInTable(${tableName}) error:`, err.message);
    return null;
  }
}

export async function markRecordModified(tableName, id, userContext) {
  if (!base || !userContext || !TRACKED_TABLES.includes(tableName)) return null;
  try {
    const updateFields = {};
    if (userContext.name) updateFields["Last Site User Name"] = userContext.name;
    if (userContext.email) updateFields["Last Site User Email"] = userContext.email;
    if (Object.keys(updateFields).length === 0) return null;
    const record = await base(tableName).update(id, updateFields);
    return formatRecord(record);
  } catch (err) {
    console.error(`markRecordModified(${tableName}) error:`, err.message);
    return null;
  }
}

export async function createOpportunity(name, contactId, opportunityType = "Home Loans", userContext = null, additionalFields = {}) {
  if (!base) return null;
  try {
    const fields = {
      "Opportunity Name": name,
      "Primary Applicant": [contactId],
      "Status": "Open",
      "Opportunity Type": opportunityType,
      ...additionalFields
    };
    if (userContext) {
      if (userContext.name) {
        fields["Creating Site User Name"] = userContext.name;
        fields["Last Site User Name"] = userContext.name;
      }
      if (userContext.email) {
        fields["Creating Site User Email"] = userContext.email;
        fields["Last Site User Email"] = userContext.email;
      }
    }
    const record = await base("Opportunities").create(fields);
    return formatRecord(record);
  } catch (err) {
    console.error("createOpportunity error:", err.message);
    return null;
  }
}

export async function setSpouse(contactId, spouseId, action) {
  if (!base) return null;
  try {
    if (!contactId || !spouseId) {
      console.error("setSpouse: Missing contactId or spouseId");
      return null;
    }

    const actionValue = action.includes("disconnected") 
      ? "disconnected as spouse from" 
      : "connected as spouse to";

    await base("Spouse History").create({
      "Contact 1": [contactId],
      "Contact 2": [spouseId],
      "Connected or Disconnected": actionValue
    });

    await new Promise(resolve => setTimeout(resolve, 1500));

    const updated = await base("Contacts").find(contactId);
    return formatRecord(updated);
  } catch (err) {
    console.error("setSpouse error:", err.message);
    return null;
  }
}

export async function deleteContact(contactId) {
  if (!base) return { success: false, error: "Database not configured" };
  try {
    const contact = await base("Contacts").find(contactId);
    if (!contact) return { success: false, error: "Contact not found" };
    
    const connections = [];
    
    const spouse = contact.fields["Spouse"];
    if (spouse && spouse.length > 0) {
      const spouseName = contact.fields["Spouse Name"] && contact.fields["Spouse Name"][0];
      connections.push(spouseName ? `Spouse (${spouseName})` : "a Spouse");
    }
    
    const oppPrimary = contact.fields["Opportunities - Primary Applicant"];
    const oppApplicant = contact.fields["Opportunities - Applicant"];
    const totalOpps = (oppPrimary?.length || 0) + (oppApplicant?.length || 0);
    if (totalOpps > 0) {
      connections.push(`${totalOpps} Opportunit${totalOpps === 1 ? 'y' : 'ies'}`);
    }
    
    if (connections.length > 0) {
      return { 
        success: false, 
        error: `This Contact is currently connected to ${connections.join(" and ")}. If you're sure this Contact should be deleted, please first remove all connections.`
      };
    }
    
    await base("Contacts").destroy(contactId);
    return { success: true };
  } catch (err) {
    console.error("deleteContact error:", err.message);
    return { success: false, error: err.message };
  }
}

export async function deleteOpportunity(opportunityId) {
  if (!base) return { success: false, error: "Database not configured" };
  try {
    const opportunity = await base("Opportunities").find(opportunityId);
    if (!opportunity) return { success: false, error: "Opportunity not found" };
    
    const connections = [];
    
    // NOTE: Primary Applicant is allowed - user can delete opportunity with only Primary Applicant connected
    // All other connections must be removed first
    
    const applicants = opportunity.fields["Applicants"];
    if (applicants && applicants.length > 0) {
      connections.push(`${applicants.length} Applicant${applicants.length === 1 ? '' : 's'}`);
    }
    
    const guarantors = opportunity.fields["Guarantors"];
    if (guarantors && guarantors.length > 0) {
      connections.push(`${guarantors.length} Guarantor${guarantors.length === 1 ? '' : 's'}`);
    }
    
    const loanApps = opportunity.fields["Loan Applications"];
    if (loanApps && loanApps.length > 0) {
      connections.push(`${loanApps.length} Loan Application${loanApps.length === 1 ? '' : 's'}`);
    }
    
    const tasks = opportunity.fields["Tasks"];
    if (tasks && tasks.length > 0) {
      connections.push(`${tasks.length} Task${tasks.length === 1 ? '' : 's'}`);
    }
    
    if (connections.length > 0) {
      return { 
        success: false, 
        error: `This Opportunity is currently connected to ${connections.join(", ")}. If you're sure this Opportunity should be deleted, please first remove these connections. (Primary Applicant can remain connected.)`
      };
    }
    
    await base("Opportunities").destroy(opportunityId);
    return { success: true };
  } catch (err) {
    console.error("deleteOpportunity error:", err.message);
    return { success: false, error: err.message };
  }
}

// --- SETTINGS TABLE ---
const settingsCache = new Map();
const SETTINGS_CACHE_TTL = 60000; // 1 minute cache

export async function getSetting(key) {
  if (!base || !key) return null;
  
  const cacheKey = key.toLowerCase();
  const cached = settingsCache.get(cacheKey);
  if (cached && (Date.now() - cached.timestamp < SETTINGS_CACHE_TTL)) {
    return cached.value;
  }
  
  try {
    const records = await base("Settings")
      .select({
        filterByFormula: `LOWER({Setting Key}) = "${cacheKey}"`,
        maxRecords: 1
      })
      .all();
    
    if (records.length > 0) {
      const value = records[0].fields["Value"] || null;
      settingsCache.set(cacheKey, { value, timestamp: Date.now() });
      return value;
    }
    return null;
  } catch (err) {
    console.error("getSetting error:", err.message);
    return null;
  }
}

export async function getAllSettings() {
  if (!base) return {};
  
  try {
    const records = await base("Settings")
      .select({ maxRecords: 100 })
      .all();
    
    const settings = {};
    records.forEach(record => {
      const key = record.fields["Setting Key"];
      const value = record.fields["Value"];
      if (key) {
        settings[key] = value || null;
        settingsCache.set(key.toLowerCase(), { value, timestamp: Date.now() });
      }
    });
    return settings;
  } catch (err) {
    console.error("getAllSettings error:", err.message);
    return {};
  }
}

export async function updateSetting(key, value, userEmail = null) {
  if (!base || !key) return null;
  
  try {
    const records = await base("Settings")
      .select({
        filterByFormula: `LOWER({Setting Key}) = "${key.toLowerCase()}"`,
        maxRecords: 1
      })
      .all();
    
    let result;
    const updateFields = { "Value": value };
    
    if (userEmail) {
      const userRecords = await base("Users")
        .select({
          filterByFormula: `LOWER({Email}) = "${userEmail.toLowerCase()}"`,
          maxRecords: 1
        })
        .all();
      if (userRecords.length > 0) {
        updateFields["Last Updated By"] = [userRecords[0].id];
      }
    }
    
    if (records.length > 0) {
      result = await base("Settings").update(records[0].id, updateFields);
    } else {
      result = await base("Settings").create({
        "Setting Key": key,
        ...updateFields
      });
    }
    
    settingsCache.set(key.toLowerCase(), { value, timestamp: Date.now() });
    return formatRecord(result);
  } catch (err) {
    console.error("updateSetting error:", err.message);
    return null;
  }
}

export function clearSettingsCache() {
  settingsCache.clear();
}

// --- APPOINTMENTS CRUD ---

export async function getAppointmentsForOpportunity(opportunityId) {
  if (!base || !opportunityId) return [];
  
  try {
    console.log("Fetching appointments for opportunity:", opportunityId);
    
    // Fetch all appointments and filter by linked Opportunity record ID
    // ARRAYJOIN returns display names not record IDs, so we filter server-side
    const allRecords = await base("Appointments")
      .select({
        sort: [{ field: "Appointment Time", direction: "asc" }]
      })
      .all();
    
    // Filter to only appointments linked to this opportunity
    const records = allRecords.filter(r => {
      const linkedOpps = r.fields["Opportunity"] || [];
      return linkedOpps.includes(opportunityId);
    });
    
    console.log("Appointments found:", records.length, "out of", allRecords.length, "total");
    
    // Collect all unique user IDs for batch lookup
    const userIds = new Set();
    records.forEach(r => {
      const createdBy = Array.isArray(r.fields["Created By"]) ? r.fields["Created By"][0] : null;
      const modifiedBy = Array.isArray(r.fields["Modified By"]) ? r.fields["Modified By"][0] : null;
      if (createdBy) userIds.add(createdBy);
      if (modifiedBy) userIds.add(modifiedBy);
    });
    
    // Look up all user names in parallel
    const userLookups = await Promise.all([...userIds].map(id => getUserById(id)));
    const userMap = new Map();
    userLookups.forEach(user => {
      if (user && user.id) userMap.set(user.id, user.name || 'Unknown');
    });
    
    return records.map(r => {
      const createdById = Array.isArray(r.fields["Created By"]) ? r.fields["Created By"][0] : null;
      const modifiedById = Array.isArray(r.fields["Modified By"]) ? r.fields["Modified By"][0] : null;
      
      return {
        id: r.id,
        appointmentTime: r.fields["Appointment Time"] || null,
        typeOfAppointment: r.fields["Type of Appointment"] || null,
        howBooked: r.fields["How Booked"] || null,
        howBookedOther: r.fields["How Booked Other"] || null,
        phoneNumber: r.fields["Phone Number"] || null,
        videoMeetUrl: r.fields["Video Meet URL"] || null,
        needEvidenceInAdvance: r.fields["Need Evidence in Advance"] || false,
        needApptReminder: r.fields["Need Appt Reminder"] || false,
        confEmailSent: r.fields["Conf Email Sent"] || false,
        confTextSent: r.fields["Conf Text Sent"] || false,
        appointmentStatus: r.fields["Appointment Status"] || null,
        notes: r.fields["Notes"] || null,
        createdTime: r.fields["Created Time"] || null,
        modifiedTime: r.fields["Modified Time"] || null,
        createdById,
        modifiedById,
        createdByName: createdById ? userMap.get(createdById) : null,
        modifiedByName: modifiedById ? userMap.get(modifiedById) : null
      };
    });
  } catch (err) {
    console.error("getAppointmentsForOpportunity error:", err.message);
    return [];
  }
}

export async function createAppointment(opportunityId, fields, userContext = null) {
  if (!base || !opportunityId) return null;
  
  try {
    // Build create fields - only include fields that have values
    const createFields = {
      "Opportunity": [opportunityId]
    };
    
    // Map frontend field names to Airtable field names and only include non-empty values
    if (fields["Appointment Time"]) createFields["Appointment Time"] = fields["Appointment Time"];
    if (fields["Type of Appointment"]) createFields["Type of Appointment"] = fields["Type of Appointment"];
    if (fields["How Booked"]) createFields["How Booked"] = fields["How Booked"];
    if (fields["How Booked Other"]) createFields["How Booked Other"] = fields["How Booked Other"];
    if (fields["Phone Number"]) createFields["Phone Number"] = fields["Phone Number"];
    if (fields["Video Meet URL"]) createFields["Video Meet URL"] = fields["Video Meet URL"];
    if (fields["Notes"]) createFields["Notes"] = fields["Notes"];
    
    // Booleans - always set
    createFields["Need Evidence in Advance"] = fields["Need Evidence in Advance"] === true;
    createFields["Need Appt Reminder"] = fields["Need Appt Reminder"] === true;
    
    // Only set Appointment Status if the field exists in Airtable (skip to avoid select option errors)
    // The field should default in Airtable or we skip it
    
    if (userContext && userContext.id) {
      createFields["Created By"] = [userContext.id];
    }
    
    console.log("Creating appointment with fields:", JSON.stringify(createFields, null, 2));
    
    const record = await base("Appointments").create(createFields);
    return formatRecord(record);
  } catch (err) {
    console.error("createAppointment error:", err.message);
    throw err; // Re-throw so the API can return proper error to frontend
  }
}

export async function updateAppointment(appointmentId, field, value, userContext = null) {
  if (!base || !appointmentId) return null;
  
  // Map frontend field names to Airtable field names
  const fieldMap = {
    'appointmentTime': 'Appointment Time',
    'typeOfAppointment': 'Type of Appointment',
    'howBooked': 'How Booked',
    'howBookedOther': 'How Booked Other',
    'phoneNumber': 'Phone Number',
    'videoMeetUrl': 'Video Meet URL',
    'needEvidenceInAdvance': 'Need Evidence in Advance',
    'needApptReminder': 'Need Appt Reminder',
    'confEmailSent': 'Conf Email Sent',
    'confTextSent': 'Conf Text Sent',
    'appointmentStatus': 'Appointment Status',
    'notes': 'Notes'
  };
  
  const airtableField = fieldMap[field] || field;
  
  try {
    const updateFields = {
      [airtableField]: value || null
    };
    
    // Set Modified By if user context is available
    if (userContext && userContext.id) {
      updateFields["Modified By"] = [userContext.id];
    }
    
    const record = await base("Appointments").update(appointmentId, updateFields);
    return formatRecord(record);
  } catch (err) {
    console.error("updateAppointment error:", err.message);
    return null;
  }
}

export async function updateAppointmentFields(appointmentId, fields, userContext = null) {
  if (!base || !appointmentId) return null;
  
  try {
    // Track who modified the record
    if (userContext && userContext.id) {
      fields["Modified By"] = [userContext.id];
    }
    
    const record = await base("Appointments").update(appointmentId, fields);
    return formatRecord(record);
  } catch (err) {
    console.error("updateAppointmentFields error:", err.message);
    return null;
  }
}

export async function deleteAppointment(appointmentId) {
  if (!base || !appointmentId) return false;
  
  try {
    await base("Appointments").destroy(appointmentId);
    return true;
  } catch (err) {
    console.error("deleteAppointment error:", err.message);
    return false;
  }
}
