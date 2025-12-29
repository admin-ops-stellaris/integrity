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
        name: records[0].fields["Name"] || null,
        email: email
      };
      userProfileCache.set(cacheKey, profile);
      return profile;
    }
    
    console.warn(`User not found in Users table for email: ${email}`);
    return { name: null, email: email };
  } catch (err) {
    console.error("getUserProfileByEmail error:", err.message);
    return { name: null, email: email };
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

export async function updateOpportunity(id, field, value) {
  if (!base) return null;
  try {
    const record = await base("Opportunities").update(id, {
      [field]: value
    });
    return formatRecord(record);
  } catch (err) {
    console.error("updateOpportunity error:", err.message);
    return null;
  }
}

export async function updateRecordInTable(tableName, id, field, value, userContext = null) {
  if (!base) return null;
  try {
    const updateFields = { [field]: value };
    if (userContext && tableName === "Contacts") {
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

export async function createOpportunity(name, contactId, userContext = null) {
  if (!base) return null;
  try {
    const fields = {
      "Opportunity Name": name,
      "Primary Applicant": [contactId],
      "Status": "Open"
    };
    if (userContext) {
      if (userContext.name) fields["Creating Site User Name"] = userContext.name;
      if (userContext.email) fields["Creating Site User Email"] = userContext.email;
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
    const contact = await base("Contacts").find(contactId);
    const spouse = await base("Contacts").find(spouseId);
    
    if (!contact || !spouse) return null;

    const timestamp = new Date().toISOString().split('T')[0];
    const historyEntry = `${timestamp}: ${action}`;

    if (action.includes("connected")) {
      const contactHistory = contact.fields["Spouse History Text"] || [];
      const spouseHistory = spouse.fields["Spouse History Text"] || [];
      
      await base("Contacts").update([
        {
          id: contactId,
          fields: {
            Spouse: [spouseId],
            "Spouse History Text": [...contactHistory, historyEntry]
          }
        },
        {
          id: spouseId,
          fields: {
            Spouse: [contactId],
            "Spouse History Text": [...spouseHistory, historyEntry]
          }
        }
      ]);
    } else if (action.includes("disconnected")) {
      const contactHistory = contact.fields["Spouse History Text"] || [];
      const spouseHistory = spouse.fields["Spouse History Text"] || [];
      
      await base("Contacts").update([
        {
          id: contactId,
          fields: {
            Spouse: [],
            "Spouse History Text": [...contactHistory, historyEntry]
          }
        },
        {
          id: spouseId,
          fields: {
            Spouse: [],
            "Spouse History Text": [...spouseHistory, historyEntry]
          }
        }
      ]);
    }

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
    
    const opportunities = contact.fields["Opportunities"];
    if (opportunities && opportunities.length > 0) {
      connections.push(`${opportunities.length} Opportunit${opportunities.length === 1 ? 'y' : 'ies'}`);
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
