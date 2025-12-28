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

export async function getContactById(id) {
  if (!base) return null;
  try {
    const record = await base("Contacts").find(id);
    return formatRecord(record);
  } catch (err) {
    console.error("getContactById error:", err.message);
    return null;
  }
}

export async function getRecentContacts() {
  if (!base) return [];
  try {
    const records = await base("Contacts")
      .select({
        maxRecords: 50,
        sort: [{ field: "Modified", direction: "desc" }]
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
    const q = query.toLowerCase();
    const records = await base("Contacts")
      .select({
        filterByFormula: `OR(
          SEARCH("${q}", LOWER({FirstName})),
          SEARCH("${q}", LOWER({LastName})),
          SEARCH("${q}", LOWER({PreferredName})),
          SEARCH("${q}", LOWER({EmailAddress1}))
        )`
      })
      .all();
    return records.map(formatRecord);
  } catch (err) {
    console.error("searchContacts error:", err.message);
    return [];
  }
}

export async function updateContact(id, field, value) {
  if (!base) return null;
  try {
    const record = await base("Contacts").update(id, {
      [field]: value
    });
    return formatRecord(record);
  } catch (err) {
    console.error("updateContact error:", err.message);
    return null;
  }
}

export async function createContact(fields) {
  if (!base) return null;
  try {
    const record = await base("Contacts").create(fields);
    return formatRecord(record);
  } catch (err) {
    console.error("createContact error:", err.message);
    return null;
  }
}

export async function getOpportunitiesById(ids) {
  if (!base || !ids || ids.length === 0) return [];
  try {
    const filterFormula = `OR(${ids.map(id => `RECORD_ID()="${id}"`).join(",")})`;
    const records = await base("Opportunities")
      .select({ filterByFormula })
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
