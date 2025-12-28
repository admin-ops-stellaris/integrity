// Mock data store
let contacts = [
  {
    id: "c1",
    fields: {
      FirstName: "John",
      MiddleName: "Michael",
      LastName: "Smith",
      PreferredName: "J.M.",
      Mobile: "555-0101",
      EmailAddress1: "john.smith@example.com",
      Description: "Loyal client",
      Created: "14:30 01/15/2023",
      Modified: "10:45 12/28/2024",
      Duplicate: null,
      "Duplicate Warning": null,
      Spouse: [],
      "Spouse Name": [],
      "Spouse History Text": [],
      "Opportunities - Primary Applicant": ["o1"],
      "Opportunities - Applicant": [],
      "Opportunities - Guarantor": []
    }
  },
  {
    id: "c2",
    fields: {
      FirstName: "Sarah",
      MiddleName: "Elizabeth",
      LastName: "Johnson",
      PreferredName: "Liz",
      Mobile: "555-0102",
      EmailAddress1: "sarah.johnson@example.com",
      Description: "Project lead",
      Created: "09:15 03/22/2023",
      Modified: "11:20 12/27/2024",
      Duplicate: null,
      "Duplicate Warning": null,
      Spouse: ["c3"],
      "Spouse Name": ["David Johnson"],
      "Spouse History Text": ["2024-06-15: connected as spouse to David Johnson"],
      "Opportunities - Primary Applicant": ["o2"],
      "Opportunities - Applicant": ["o3"],
      "Opportunities - Guarantor": []
    }
  },
  {
    id: "c3",
    fields: {
      FirstName: "David",
      MiddleName: "",
      LastName: "Johnson",
      PreferredName: "",
      Mobile: "555-0103",
      EmailAddress1: "david.johnson@example.com",
      Description: "Spouse of Sarah",
      Created: "12:45 05/10/2023",
      Modified: "11:20 12/27/2024",
      Duplicate: null,
      "Duplicate Warning": null,
      Spouse: ["c2"],
      "Spouse Name": ["Sarah Johnson"],
      "Spouse History Text": ["2024-06-15: connected as spouse to Sarah Johnson"],
      "Opportunities - Primary Applicant": ["o2"],
      "Opportunities - Applicant": [],
      "Opportunities - Guarantor": ["o3"]
    }
  }
];

let opportunities = [
  {
    id: "o1",
    fields: {
      "Opportunity Name": "Home Purchase - Smith",
      "Opportunity Type": "Primary Residence",
      Status: "In Progress",
      Amount: "$350,000"
    }
  },
  {
    id: "o2",
    fields: {
      "Opportunity Name": "Investment Property - Johnson",
      "Opportunity Type": "Investment",
      Status: "Qualifying",
      Amount: "$500,000"
    }
  },
  {
    id: "o3",
    fields: {
      "Opportunity Name": "Refinance - Johnson",
      "Opportunity Type": "Refinance",
      Status: "Documentation",
      Amount: "$450,000"
    }
  }
];

export function getContactById(id) {
  return contacts.find(c => c.id === id) || null;
}

export function getRecentContacts() {
  return [...contacts].sort((a, b) => 
    new Date(b.fields.Modified) - new Date(a.fields.Modified)
  );
}

export function searchContacts(query) {
  const q = query.toLowerCase();
  return contacts.filter(c => {
    const f = c.fields;
    const fullName = `${f.FirstName} ${f.MiddleName} ${f.LastName} ${f.PreferredName}`.toLowerCase();
    const email = (f.EmailAddress1 || "").toLowerCase();
    return fullName.includes(q) || email.includes(q);
  });
}

export function updateContact(id, field, value) {
  const contact = contacts.find(c => c.id === id);
  if (contact) {
    contact.fields[field] = value;
    contact.fields.Modified = new Date().toLocaleString();
  }
  return contact;
}

export function createContact(fields) {
  const id = `c${Date.now()}`;
  const contact = {
    id,
    fields: {
      ...fields,
      Created: new Date().toLocaleString(),
      Modified: new Date().toLocaleString()
    }
  };
  contacts.push(contact);
  return contact;
}

export function getOpportunitiesById(ids) {
  return opportunities.filter(o => ids.includes(o.id));
}

export function setSpouse(contactId, spouseId, action) {
  const contact = contacts.find(c => c.id === contactId);
  const spouse = contacts.find(c => c.id === spouseId);
  if (!contact || !spouse) return null;

  if (action.includes("connected")) {
    contact.fields.Spouse = [spouseId];
    contact.fields["Spouse Name"] = [`${spouse.fields.FirstName} ${spouse.fields.LastName}`];
    spouse.fields.Spouse = [contactId];
    spouse.fields["Spouse Name"] = [`${contact.fields.FirstName} ${contact.fields.LastName}`];
    
    const timestamp = new Date().toISOString().split('T')[0];
    const historyEntry = `${timestamp}: ${action}`;
    if (!contact.fields["Spouse History Text"]) contact.fields["Spouse History Text"] = [];
    if (!spouse.fields["Spouse History Text"]) spouse.fields["Spouse History Text"] = [];
    contact.fields["Spouse History Text"].push(historyEntry);
    spouse.fields["Spouse History Text"].push(historyEntry);
  } else if (action.includes("disconnected")) {
    contact.fields.Spouse = [];
    contact.fields["Spouse Name"] = [];
    spouse.fields.Spouse = [];
    spouse.fields["Spouse Name"] = [];
  }

  return contact;
}
