import Airtable from "airtable";

const AIRTABLE_API_KEY = process.env.AIRTABLE_API_KEY;
const AIRTABLE_BASE_ID = process.env.AIRTABLE_BASE_ID;

if (!AIRTABLE_API_KEY || !AIRTABLE_BASE_ID) {
  console.warn("Warning: Airtable credentials not configured. Using mock mode.");
}

const base = AIRTABLE_API_KEY && AIRTABLE_BASE_ID 
  ? new Airtable({ apiKey: AIRTABLE_API_KEY }).base(AIRTABLE_BASE_ID)
  : null;

// Helper to generate Perth time ISO string
function getPerthTimeISO() {
  const now = new Date();
  const perthOffset = 8 * 60; // Perth is UTC+8
  const perthTime = new Date(now.getTime() + (perthOffset + now.getTimezoneOffset()) * 60 * 1000);
  const year = perthTime.getFullYear();
  const month = String(perthTime.getMonth() + 1).padStart(2, '0');
  const day = String(perthTime.getDate()).padStart(2, '0');
  const hours = String(perthTime.getHours()).padStart(2, '0');
  const minutes = String(perthTime.getMinutes()).padStart(2, '0');
  const seconds = String(perthTime.getSeconds()).padStart(2, '0');
  return `${year}-${month}-${day}T${hours}:${minutes}:${seconds}.000+08:00`;
}

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
  // Match "HH:MM DD/MM/YYYY" format from Perth timezone (GMT+8)
  const match = modifiedText.match(/(\d{2}):(\d{2})\s+(\d{2})\/(\d{2})\/(\d{4})/);
  if (match) {
    const [, hours, mins, day, month, year] = match;
    // Create UTC timestamp and subtract 8 hours to convert Perth time to UTC
    const utcMs = Date.UTC(parseInt(year), parseInt(month) - 1, parseInt(day), parseInt(hours), parseInt(mins));
    const perthOffsetMs = 8 * 60 * 60 * 1000; // GMT+8
    return new Date(utcMs - perthOffsetMs);
  }
  return new Date(0);
}

export async function getRecentContacts(statusFilter = null) {
  if (!base) return [];
  try {
    // Build filter formula based on status
    let filterFormula = null;
    if (statusFilter && statusFilter !== 'All') {
      filterFormula = `{Status} = "${statusFilter}"`;
    }
    
    // Fetch contacts sorted by "Modified On (Web App)" descending at Airtable level
    const selectOptions = {
      maxRecords: 50,
      sort: [{ field: "Modified On (Web App)", direction: "desc" }]
    };
    if (filterFormula) {
      selectOptions.filterByFormula = filterFormula;
    }
    
    const records = await base("Contacts").select(selectOptions).all();
    const formatted = records.map(formatRecord);
    
    // Secondary sort in JS by Modified formula field for better ordering
    formatted.sort((a, b) => {
      const dateA = parseModifiedDate(a.fields.Modified);
      const dateB = parseModifiedDate(b.fields.Modified);
      return dateB - dateA; // Most recent first
    });
    return formatted;
  } catch (err) {
    console.error("getRecentContacts error:", err.message);
    return [];
  }
}

export async function searchContacts(query, statusFilter = null) {
  if (!base) return [];
  if (!query || query.trim() === "") return getRecentContacts(statusFilter);
  
  try {
    const terms = query.toLowerCase().split(/\s+/).filter(t => t.length > 0);
    const searchConditions = terms.map(term => `SEARCH("${term}", {SEARCH_INDEX})`);
    
    // Build formula with optional status filter
    let formula = terms.length === 1 
      ? searchConditions[0]
      : `AND(${searchConditions.join(", ")})`;
    
    if (statusFilter && statusFilter !== 'All') {
      formula = `AND(${formula}, {Status} = "${statusFilter}")`;
    }
    
    const records = await base("Contacts")
      .select({ filterByFormula: formula })
      .all();
    
    // Score and sort results by field priority
    const scored = records.map(record => {
      const f = record.fields;
      let score = 0;
      
      // Field weights (higher = more important)
      const fieldWeights = [
        { value: f.FirstName, weight: 100, bonus: 50 },
        { value: f.MiddleName, weight: 80, bonus: 30 },
        { value: f.LastName, weight: 70, bonus: 25 },
        { value: f.PreferredName, weight: 60, bonus: 20 },
        { value: f.EmailAddress1, weight: 40, bonus: 10 },
        { value: f.EmailAddress2, weight: 35, bonus: 10 },
        { value: f.EmailAddress3, weight: 30, bonus: 10 },
        { value: f.Mobile, weight: 25, bonus: 5 },
        { value: f.Notes, weight: 5, bonus: 0 }
      ];
      
      for (const term of terms) {
        for (const field of fieldWeights) {
          const val = (field.value || '').toLowerCase();
          if (!val) continue;
          
          if (val === term) {
            // Exact match - highest priority
            score += field.weight + field.bonus + 20;
          } else if (val.startsWith(term)) {
            // Starts with match - high priority
            score += field.weight + field.bonus;
          } else if (val.includes(term)) {
            // Contains match - normal priority
            score += field.weight;
          }
        }
      }
      
      return { record: formatRecord(record), score };
    });
    
    // Sort by score descending, then alphabetically by name
    scored.sort((a, b) => {
      if (b.score !== a.score) return b.score - a.score;
      const nameA = `${a.record.fields.FirstName || ''} ${a.record.fields.LastName || ''}`.toLowerCase();
      const nameB = `${b.record.fields.FirstName || ''} ${b.record.fields.LastName || ''}`.toLowerCase();
      return nameA.localeCompare(nameB);
    });
    
    return scored.map(s => s.record);
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
      if (userContext.name) {
        updateFields["Modified By (Web App User)"] = userContext.name;
      }
      if (userContext.email) {
        updateFields["Modified By (Web App User Email)"] = userContext.email;
      }
      // Add modified timestamp in ISO format with Perth timezone (GMT+8)
      updateFields["Modified On (Web App)"] = getPerthTimeISO();
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
        createFields["Created By (Web App User)"] = userContext.name;
      }
      if (userContext.email) {
        createFields["Created By (Web App User Email)"] = userContext.email;
      }
      // Add created timestamp in ISO format with Perth timezone (GMT+8)
      createFields["Created On (Web App)"] = getPerthTimeISO();
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

// Mark a contact as modified (updates Modified On/By fields)
async function markContactModified(contactId, userContext) {
  if (!base || !contactId || !userContext) return;
  try {
    const updateFields = {};
    if (userContext.name) updateFields["Modified By (Web App User)"] = userContext.name;
    if (userContext.email) updateFields["Modified By (Web App User Email)"] = userContext.email;
    updateFields["Modified On (Web App)"] = getPerthTimeISO();
    await base("Contacts").update(contactId, updateFields);
  } catch (err) {
    console.error("markContactModified error:", err.message);
  }
}

// Contact link fields on Opportunities
const CONTACT_LINK_FIELDS = ["Primary Applicant", "Applicants", "Guarantors"];

export async function updateOpportunity(id, field, value, userContext = null) {
  if (!base) return null;
  try {
    // If updating a contact link field, get old values first to mark removed contacts
    let oldContactIds = [];
    if (CONTACT_LINK_FIELDS.includes(field) && userContext) {
      const oldRecord = await base("Opportunities").find(id);
      oldContactIds = oldRecord.fields[field] || [];
    }
    
    const updateFields = { [field]: value };
    if (userContext) {
      if (userContext.name) updateFields["Modified By (Web App User)"] = userContext.name;
      if (userContext.email) updateFields["Modified By (Web App User Email)"] = userContext.email;
      updateFields["Modified On (Web App)"] = getPerthTimeISO();
    }
    const record = await base("Opportunities").update(id, updateFields);
    
    // If updating a contact link field, mark both old and new contacts as modified
    if (CONTACT_LINK_FIELDS.includes(field) && userContext) {
      const newContactIds = Array.isArray(value) ? value : (value ? [value] : []);
      const allContactIds = [...new Set([...oldContactIds, ...newContactIds])];
      for (const contactId of allContactIds) {
        await markContactModified(contactId, userContext);
      }
    }
    
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
      if (userContext.name) updateFields["Modified By (Web App User)"] = userContext.name;
      if (userContext.email) updateFields["Modified By (Web App User Email)"] = userContext.email;
      updateFields["Modified On (Web App)"] = getPerthTimeISO();
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
    if (userContext.name) updateFields["Modified By (Web App User)"] = userContext.name;
    if (userContext.email) updateFields["Modified By (Web App User Email)"] = userContext.email;
    updateFields["Modified On (Web App)"] = getPerthTimeISO();
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
    // Set Created and Modified timestamps (both set on creation)
    const perthTime = getPerthTimeISO();
    fields["Created On (Web App)"] = perthTime;
    fields["Modified On (Web App)"] = perthTime;
    
    if (userContext) {
      if (userContext.name) {
        fields["Created By (Web App User)"] = userContext.name;
        fields["Modified By (Web App User)"] = userContext.name;
      }
      if (userContext.email) {
        fields["Created By (Web App User Email)"] = userContext.email;
        fields["Modified By (Web App User Email)"] = userContext.email;
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

// ==================== CONNECTIONS ====================

// Valid connection role pairs (Record1Role -> Record2Role)
const CONNECTION_ROLE_PAIRS = [
  { role1: "Parent", role2: "Child", label: "Parent / Child" },
  { role1: "Child", role2: "Parent", label: "Child / Parent" },
  { role1: "Sibling", role2: "Sibling", label: "Sibling / Sibling" },
  { role1: "Friend", role2: "Friend", label: "Friend / Friend" },
  { role1: "Household Representative", role2: "Household Member", label: "Household Representative / Household Member" },
  { role1: "Household Member", role2: "Household Representative", label: "Household Member / Household Representative" },
  { role1: "Employer of", role2: "Employee of", label: "Employer / Employee" },
  { role1: "Employee of", role2: "Employer of", label: "Employee / Employer" },
  { role1: "Referred by", role2: "Has Referred", label: "Referred by / Has Referred" },
  { role1: "Has Referred", role2: "Referred by", label: "Has Referred / Referred by" }
];

export function getConnectionRoleTypes() {
  return CONNECTION_ROLE_PAIRS;
}

export async function getConnectionsForContact(contactId) {
  if (!base || !contactId) return [];
  try {
    // Get connection record IDs from the Contact's back-link fields
    const contact = await base("Contacts").find(contactId);
    if (!contact) return [];
    
    // Use the correct back-link field names
    const connectionsAsContact1 = contact.fields["BACK-LINK: Connections: Contact 1"] || [];
    const connectionsAsContact2 = contact.fields["BACK-LINK: Connections: Contact 2"] || [];
    const allConnectionIds = [...new Set([...connectionsAsContact1, ...connectionsAsContact2])];
    
    if (allConnectionIds.length === 0) return [];
    
    // BATCH FETCH: Get all connection records in one query using formula
    const connectionFormula = `OR(${allConnectionIds.map(id => `RECORD_ID()='${id}'`).join(',')})`;
    const records = await base("Connections").select({
      filterByFormula: connectionFormula
    }).all();
    
    // Collect all "other" contact IDs and user IDs for batch fetching
    const otherContactIds = new Set();
    const userIds = new Set();
    
    records.forEach(record => {
      const f = record.fields;
      const contact1Ids = f["Contact 1"] || [];
      const contact2Ids = f["Contact 2"] || [];
      const isContact1 = contact1Ids.includes(contactId);
      const otherContactId = isContact1 ? contact2Ids[0] : contact1Ids[0];
      if (otherContactId) otherContactIds.add(otherContactId);
      
      // Collect user IDs from Created By and Modified By linked fields
      const createdByIds = f["Created By"] || [];
      const modifiedByIds = f["Modified By"] || [];
      if (createdByIds[0]) userIds.add(createdByIds[0]);
      if (modifiedByIds[0]) userIds.add(modifiedByIds[0]);
    });
    
    // BATCH FETCH: Get all other contacts in one query
    const contactsMap = {};
    if (otherContactIds.size > 0) {
      const contactFormula = `OR(${[...otherContactIds].map(id => `RECORD_ID()='${id}'`).join(',')})`;
      const contactRecords = await base("Contacts").select({
        filterByFormula: contactFormula,
        fields: ["FirstName", "MiddleName", "LastName", "Calculated Name"]
      }).all();
      contactRecords.forEach(c => {
        const cf = c.fields;
        contactsMap[c.id] = cf["Calculated Name"] || 
          `${cf.FirstName || ''} ${cf.MiddleName || ''} ${cf.LastName || ''}`.replace(/\s+/g, ' ').trim();
      });
    }
    
    // BATCH FETCH: Get all users in one query for Created By / Modified By names
    const usersMap = {};
    if (userIds.size > 0) {
      const userFormula = `OR(${[...userIds].map(id => `RECORD_ID()='${id}'`).join(',')})`;
      const userRecords = await base("Users").select({
        filterByFormula: userFormula,
        fields: ["Name"]
      }).all();
      userRecords.forEach(u => {
        usersMap[u.id] = u.fields.Name || null;
      });
    }
    
    // Format connections with perspective from the given contact
    const connections = records.map(record => {
      const f = record.fields;
      const contact1Ids = f["Contact 1"] || [];
      const contact2Ids = f["Contact 2"] || [];
      const isContact1 = contact1Ids.includes(contactId);
      
      // Get the "other" contact's info
      const otherContactId = isContact1 ? contact2Ids[0] : contact1Ids[0];
      const fallbackName = isContact1 
        ? (f["Record2IdName"] || "Unknown") 
        : (f["Record1IdName"] || "Unknown");
      const myRole = isContact1 ? f["Record1Role"] : f["Record2Role"];
      const theirRole = isContact1 ? f["Record2Role"] : f["Record1Role"];
      
      // Use batch-fetched contact name or fallback
      const displayName = otherContactId && contactsMap[otherContactId] 
        ? contactsMap[otherContactId] 
        : fallbackName;
      
      // Get user names from batch-fetched Users
      const createdById = (f["Created By"] || [])[0];
      const modifiedById = (f["Modified By"] || [])[0];
      const createdByName = createdById ? usersMap[createdById] : null;
      const modifiedByName = modifiedById ? usersMap[modifiedById] : null;
      
      return {
        id: record.id,
        otherContactId: otherContactId || null,
        otherContactName: displayName,
        myRole: myRole,
        theirRole: theirRole,
        status: f["Status"] || "Active",
        createdOn: f["Created On"] || null,
        modifiedOn: f["Modified On"] || null,
        createdByName: createdByName,
        modifiedByName: modifiedByName
      };
    });
    
    // Filter to active only and sort by name
    return connections
      .filter(c => c.status === "Active")
      .sort((a, b) => (a.otherContactName || "").localeCompare(b.otherContactName || ""));
  } catch (err) {
    console.error("getConnectionsForContact error:", err.message);
    return [];
  }
}

export async function createConnection(contact1Id, contact2Id, record1Role, record2Role, userContext = null) {
  if (!base) return { success: false, error: "Database not configured" };
  if (!contact1Id || !contact2Id) return { success: false, error: "Both contacts are required" };
  if (!record1Role || !record2Role) return { success: false, error: "Relationship roles are required" };
  
  try {
    // Fetch contact names for Record1IdName and Record2IdName
    const [contact1, contact2] = await Promise.all([
      base("Contacts").find(contact1Id),
      base("Contacts").find(contact2Id)
    ]);
    
    const contact1Name = contact1 
      ? `${contact1.fields.FirstName || ''} ${contact1.fields.MiddleName || ''} ${contact1.fields.LastName || ''}`.replace(/\s+/g, ' ').trim()
      : "";
    const contact2Name = contact2 
      ? `${contact2.fields.FirstName || ''} ${contact2.fields.MiddleName || ''} ${contact2.fields.LastName || ''}`.replace(/\s+/g, ' ').trim()
      : "";
    
    const createFields = {
      "Contact 1": [contact1Id],
      "Contact 2": [contact2Id],
      "Record1Role": record1Role,
      "Record2Role": record2Role,
      "Record1IdName": contact1Name,
      "Record2IdName": contact2Name,
      "Status": "Active",
      "Created On": new Date().toISOString(),
      "Modified On": new Date().toISOString()
    };
    
    // Add user audit links if available
    if (userContext && userContext.id) {
      createFields["Created By"] = [userContext.id];
      createFields["Modified By"] = [userContext.id];
    }
    
    const record = await base("Connections").create(createFields);
    return { success: true, record: formatRecord(record) };
  } catch (err) {
    console.error("createConnection error:", err.message);
    return { success: false, error: err.message };
  }
}

export async function deactivateConnection(connectionId, userContext = null) {
  if (!base) return { success: false, error: "Database not configured" };
  if (!connectionId) return { success: false, error: "Connection ID is required" };
  
  try {
    const updateFields = {
      "Status": "Inactive",
      "Modified On": new Date().toISOString()
    };
    
    if (userContext && userContext.id) {
      updateFields["Modified By"] = [userContext.id];
    }
    
    await base("Connections").update(connectionId, updateFields);
    return { success: true };
  } catch (err) {
    console.error("deactivateConnection error:", err.message);
    return { success: false, error: err.message };
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
    
    // Set Appointment Status if provided
    if (fields["Appointment Status"]) createFields["Appointment Status"] = fields["Appointment Status"];
    
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

// --- EMAIL TEMPLATES CRUD ---

const templateCache = new Map();
const TEMPLATE_CACHE_TTL = 60000; // 1 minute cache

export async function getEmailTemplates() {
  if (!base) return [];
  
  // Check cache first
  const cached = templateCache.get('all');
  if (cached && (Date.now() - cached.timestamp < TEMPLATE_CACHE_TTL)) {
    return cached.value;
  }
  
  try {
    const records = await base("Email Templates")
      .select({
        filterByFormula: "{Active} = TRUE()",
        sort: [{ field: "Template Name", direction: "asc" }]
      })
      .all();
    
    const templates = records.map(r => ({
      id: r.id,
      name: r.fields["Template Name"] || "",
      type: r.fields["Template Type"] || "General",
      subject: r.fields["Subject Template"] || "",
      body: r.fields["Body Template"] || "",
      description: r.fields["Description"] || "",
      active: r.fields["Active"] || false
    }));
    
    templateCache.set('all', { value: templates, timestamp: Date.now() });
    return templates;
  } catch (err) {
    console.error("getEmailTemplates error:", err.message);
    return [];
  }
}

export async function getEmailTemplate(templateId) {
  if (!base || !templateId) return null;
  
  try {
    const record = await base("Email Templates").find(templateId);
    if (!record) return null;
    
    return {
      id: record.id,
      name: record.fields["Template Name"] || "",
      type: record.fields["Template Type"] || "General",
      subject: record.fields["Subject Template"] || "",
      body: record.fields["Body Template"] || "",
      description: record.fields["Description"] || "",
      active: record.fields["Active"] || false
    };
  } catch (err) {
    console.error("getEmailTemplate error:", err.message);
    return null;
  }
}

export async function updateEmailTemplate(templateId, fields, userContext = null) {
  if (!base || !templateId) return null;
  
  try {
    const updateFields = {};
    
    if (fields.name !== undefined) updateFields["Template Name"] = fields.name;
    if (fields.type !== undefined) updateFields["Template Type"] = fields.type;
    if (fields.subject !== undefined) updateFields["Subject Template"] = fields.subject;
    if (fields.body !== undefined) updateFields["Body Template"] = fields.body;
    if (fields.description !== undefined) updateFields["Description"] = fields.description;
    if (fields.active !== undefined) updateFields["Active"] = fields.active;
    
    // Add modified timestamp and user
    updateFields["Modified On"] = getPerthTimeISO();
    if (userContext && userContext.id) {
      updateFields["Modified By"] = [userContext.id];
    }
    
    const record = await base("Email Templates").update(templateId, updateFields);
    
    // Clear cache
    templateCache.clear();
    
    return {
      id: record.id,
      name: record.fields["Template Name"] || "",
      type: record.fields["Template Type"] || "General",
      subject: record.fields["Subject Template"] || "",
      body: record.fields["Body Template"] || "",
      description: record.fields["Description"] || "",
      active: record.fields["Active"] || false
    };
  } catch (err) {
    console.error("updateEmailTemplate error:", err.message);
    return null;
  }
}

export async function createEmailTemplate(fields, userContext = null) {
  if (!base) return null;
  
  try {
    const perthTime = getPerthTimeISO();
    const createFields = {
      "Template Name": fields.name || "New Template",
      "Template Type": fields.type || "General",
      "Subject Template": fields.subject || "",
      "Body Template": fields.body || "",
      "Description": fields.description || "",
      "Active": fields.active !== undefined ? fields.active : true,
      "Created On": perthTime,
      "Modified On": perthTime
    };
    
    if (userContext && userContext.id) {
      createFields["Created By"] = [userContext.id];
      createFields["Modified By"] = [userContext.id];
    }
    
    const record = await base("Email Templates").create(createFields);
    
    // Clear cache
    templateCache.clear();
    
    return {
      id: record.id,
      name: record.fields["Template Name"] || "",
      type: record.fields["Template Type"] || "General",
      subject: record.fields["Subject Template"] || "",
      body: record.fields["Body Template"] || "",
      description: record.fields["Description"] || "",
      active: record.fields["Active"] || false
    };
  } catch (err) {
    console.error("createEmailTemplate error:", err.message);
    return null;
  }
}

export function clearTemplateCache() {
  templateCache.clear();
}
