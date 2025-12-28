  let searchTimeout;
  let spouseSearchTimeout;
  let linkedSearchTimeout;
  let loadingTimer;
  let pollInterval;
  let pollAttempts = 0;
  let panelHistory = []; 
  let currentContactRecord = null; 
  let currentOppRecords = []; 
  let currentOppSortDirection = 'desc'; 
  let pendingLinkedEdits = {}; 
  let currentPanelData = {}; 
  let pendingRemovals = {}; 

  window.onload = function() { loadContacts(); checkUserIdentity(); };

  function checkUserIdentity() {
    google.script.run.withSuccessHandler(function(email) {
       const display = email ? email : "Unknown";
       document.getElementById('debugUser').innerText = display;
       if (!email) alert("Warning: The system cannot detect your email address.");
    }).getEffectiveUserEmail();
  }

  function updateHeaderTitle(isEditing) {
    const fName = document.getElementById('firstName').value || "";
    const mName = document.getElementById('middleName').value || "";
    const lName = document.getElementById('lastName').value || "";
    const pName = document.getElementById('preferredName').value || "";
    let fullName = [fName, mName, lName].filter(Boolean).join(" ");
    if (pName) fullName += ` (${pName})`;
    if (!fullName.trim()) { document.getElementById('formTitle').innerText = "New Contact"; return; }
    document.getElementById('formTitle').innerText = isEditing ? `Editing ${fullName}` : fullName;
  }

  function toggleProfileView(show) {
    if(show) {
      document.getElementById('emptyState').style.display = 'none';
      document.getElementById('profileContent').style.display = 'flex';
      document.getElementById('formDivider').style.display = 'block';
    } else {
      document.getElementById('emptyState').style.display = 'flex';
      document.getElementById('profileContent').style.display = 'none';
      document.getElementById('formDivider').style.display = 'none';
      document.getElementById('formTitle').innerText = "Contact";
      document.getElementById('editBtn').style.visibility = 'hidden';
      document.getElementById('refreshBtn').style.display = 'none'; 
      document.getElementById('auditSection').style.display = 'none';
      document.getElementById('duplicateWarningBox').style.display = 'none'; 
    }
  }

  function enableEditMode() {
    const inputs = document.querySelectorAll('#contactForm input, #contactForm textarea');
    inputs.forEach(input => { input.classList.remove('locked'); input.readOnly = false; });
    document.getElementById('actionRow').style.display = 'flex';
    updateHeaderTitle(true); 
  }

  function disableEditMode() {
    const inputs = document.querySelectorAll('#contactForm input, #contactForm textarea');
    inputs.forEach(input => { input.classList.add('locked'); input.readOnly = true; });
    document.getElementById('actionRow').style.display = 'none';
    updateHeaderTitle(false); 
  }

  function selectContact(record) {
    toggleProfileView(true);
    currentContactRecord = record; 
    const f = record.fields;

    document.getElementById('recordId').value = record.id;
    document.getElementById('firstName').value = f.FirstName || "";
    document.getElementById('middleName').value = f.MiddleName || "";
    document.getElementById('lastName').value = f.LastName || "";
    document.getElementById('preferredName').value = f.PreferredName || "";
    document.getElementById('mobilePhone').value = f.Mobile || "";
    document.getElementById('email1').value = f.EmailAddress1 || "";
    document.getElementById('description').value = f.Description || "";

    disableEditMode(); 
    document.getElementById('editBtn').style.visibility = 'visible';
    document.getElementById('refreshBtn').style.display = 'inline';

    const warnBox = document.getElementById('duplicateWarningBox');
    if (f['Duplicate Warning']) {
       document.getElementById('duplicateWarningText').innerText = f['Duplicate Warning'];
       warnBox.style.display = 'flex'; 
    } else {
       warnBox.style.display = 'none';
    }

    renderHistory(f);
    loadOpportunities(f);
    renderSpouseSection(f); 
    closeOppPanel();
  }

  function refreshCurrentContact() {
     if (!currentContactRecord) return;
     const btn = document.getElementById('refreshBtn');
     btn.classList.add('spin-anim'); 
     setTimeout(() => { btn.classList.remove('spin-anim'); }, 1000); 

     const id = currentContactRecord.id;
     google.script.run.withSuccessHandler(function(r) {
        if (r && r.fields) selectContact(r);
     }).getContactById(id);
  }

  // --- SPOUSE LOGIC ---
  function renderSpouseSection(f) {
     const statusEl = document.getElementById('spouseStatusText');
     const historyList = document.getElementById('spouseHistoryList');
     const linkEl = document.getElementById('spouseEditLink');

     const spouseName = (f['Spouse Name'] && f['Spouse Name'].length > 0) ? f['Spouse Name'][0] : null;
     const spouseId = (f['Spouse'] && f['Spouse'].length > 0) ? f['Spouse'][0] : null;

     if (spouseName && spouseId) {
        statusEl.innerHTML = `Spouse: <a class="data-link" onclick="loadPanelRecord('Contacts', '${spouseId}')">${spouseName}</a>`;
        linkEl.innerText = "Edit"; 
     } else {
        statusEl.innerHTML = "Single";
        linkEl.innerText = "Edit"; 
     }

     linkEl.style.display = 'inline'; 

     historyList.innerHTML = '';
     const rawLogs = f['Spouse History Text']; 

     if (rawLogs && Array.isArray(rawLogs) && rawLogs.length > 0) {
        const sortedLogs = rawLogs.sort((a, b) => b.localeCompare(a));
        const showLimit = 3;
        const initialSet = sortedLogs.slice(0, showLimit);
        initialSet.forEach(logString => { renderHistoryItem(logString, historyList); });
        if (sortedLogs.length > showLimit) {
           const remaining = sortedLogs.slice(showLimit);
           const expandLink = document.createElement('div');
           expandLink.className = 'expand-link';
           expandLink.innerText = `Show ${remaining.length} older records...`;
           expandLink.onclick = function() {
              remaining.forEach(logString => { renderHistoryItem(logString, historyList); });
              expandLink.style.display = 'none'; 
           };
           historyList.appendChild(expandLink);
        }
     } else {
        historyList.innerHTML = '<li class="spouse-history-item" style="border:none;">No history recorded.</li>';
     }
  }

  function renderHistoryItem(logString, container) {
     const parts = logString.split(': ');
     if(parts.length < 2) return; 
     const datePartISO = parts[0]; 
     const textPart = parts.slice(1).join(': '); 
     let displayDate = datePartISO;
     const dateMatch = datePartISO.match(/^(\d{4})-(\d{2})-(\d{2})/);
     if(dateMatch) { displayDate = `${dateMatch[3]}/${dateMatch[2]}/${dateMatch[1]}`; }
     const li = document.createElement('li');
     li.className = 'spouse-history-item';
     li.innerText = `${displayDate}: ${textPart}`;
     const expandLink = container.querySelector('.expand-link');
     if(expandLink) { container.insertBefore(li, expandLink); } else { container.appendChild(li); }
  }

  // --- INLINE EDIT LOGIC ---
  function toggleFieldEdit(fieldKey) {
     document.getElementById('view_' + fieldKey).style.display = 'none';
     document.getElementById('edit_' + fieldKey).style.display = 'block';
  }
  function cancelFieldEdit(fieldKey) {
     document.getElementById('view_' + fieldKey).style.display = 'block';
     document.getElementById('edit_' + fieldKey).style.display = 'none';
  }
  function saveFieldEdit(table, id, fieldKey) {
     const input = document.getElementById('input_' + fieldKey);
     const val = input.value;
     const btn = document.getElementById('btn_save_' + fieldKey);
     const originalText = btn.innerText;
     btn.innerText = "Saving..."; btn.disabled = true;
     google.script.run.withSuccessHandler(function(res) {
        document.getElementById('display_' + fieldKey).innerText = val;
        cancelFieldEdit(fieldKey);
        btn.innerText = originalText; btn.disabled = false;
        if(fieldKey === 'Opportunity Name') {
           document.getElementById('panelTitle').innerText = val;
           if(currentContactRecord) { loadOpportunities(currentContactRecord.fields); }
        }
     }).updateRecord(table, id, fieldKey, val);
  }

  // --- LINKED RECORD EDITOR (TAGS) ---
  function toggleLinkedEdit(key) {
     document.getElementById('view_' + key).style.display = 'none';
     document.getElementById('edit_' + key).style.display = 'block';

     const currentLinks = currentPanelData[key] || [];
     pendingLinkedEdits[key] = currentLinks.map(link => ({...link}));
     pendingRemovals = {}; 
     renderLinkedEditorState(key);
     document.getElementById('error_' + key).innerText = ''; 
  }

  function cancelLinkedEdit(key) {
     document.getElementById('view_' + key).style.display = 'block';
     document.getElementById('edit_' + key).style.display = 'none';
     pendingLinkedEdits[key] = [];
     pendingRemovals = {};
  }

  function renderLinkedEditorState(key) {
     const container = document.getElementById('chip_container_' + key);
     container.innerHTML = '';
     const links = pendingLinkedEdits[key];

     if(links.length === 0) {
        container.innerHTML = '<span style="font-size:11px; color:#999; font-style:italic;">No links selected</span>';
     } else {
        links.forEach(link => {
           const chip = document.createElement('div');
           chip.className = 'link-chip';
           chip.innerHTML = `<span>${link.name}</span><span class="link-chip-remove" onclick="removePendingLink('${key}', '${link.id}')">✕</span>`;
           container.appendChild(chip);
        });
     }
  }

  function removePendingLink(key, id) {
     pendingLinkedEdits[key] = pendingLinkedEdits[key].filter(l => l.id !== id);
     renderLinkedEditorState(key);
     document.getElementById('error_' + key).innerText = ''; 
  }

  function handleLinkedSearch(event, key) {
     const query = event.target.value;
     const resultsDiv = document.getElementById('results_' + key);

     document.getElementById('error_' + key).innerText = '';

     clearTimeout(linkedSearchTimeout);
     if(query.length < 2) { resultsDiv.style.display = 'none'; return; }

     resultsDiv.style.display = 'block';
     resultsDiv.innerHTML = '<div style="padding:6px; color:#999; font-style:italic;">Searching...</div>';

     linkedSearchTimeout = setTimeout(() => {
        google.script.run.withSuccessHandler(function(records) {
           resultsDiv.innerHTML = '';
           if(records.length === 0) {
              resultsDiv.innerHTML = '<div style="padding:6px; color:#999; font-style:italic;">No results</div>';
           } else {
              records.forEach(r => {
                 const name = formatName(r.fields);
                 const details = formatDetails(r.fields);
                 const div = document.createElement('div');
                 div.className = 'link-result-item';
                 div.innerHTML = `<strong>${name}</strong> <span style="color:#888;">${details}</span>`;
                 div.onclick = function() {
                    addPendingLink(key, {id: r.id, name: name});
                    resultsDiv.style.display = 'none';
                    event.target.value = '';
                 };
                 resultsDiv.appendChild(div);
              });
           }
        }).searchContacts(query);
     }, 400);
  }

  function addPendingLink(key, newLink) {
     const errorEl = document.getElementById('error_' + key);
     errorEl.innerText = ''; 

     // 1. Enforce Primary Single
     if(key === 'Primary Applicant') {
        pendingLinkedEdits[key] = [newLink]; 
     } else {
        // 2. Check Duplicates in current list
        if(pendingLinkedEdits[key].some(l => l.id === newLink.id)) {
           errorEl.innerText = "Already added.";
           return; 
        }
        // 3. Enforce Mutually Exclusive Logic WITH PROMPT
        const exclusiveKeys = ['Primary Applicant', 'Applicants', 'Guarantors'];
        let conflictFound = false;

        exclusiveKeys.forEach(otherKey => {
           if(otherKey === key) return;
           const otherLinks = currentPanelData[otherKey] || [];
           if(otherLinks.some(l => l.id === newLink.id)) {
              conflictFound = true;
              if(confirm(`${newLink.name} is currently a '${otherKey}'.\n\nDo you want to move them to '${key}'?`)) {
                 if(!pendingRemovals[otherKey]) pendingRemovals[otherKey] = [];
                 pendingRemovals[otherKey].push(newLink.id);
                 pendingLinkedEdits[key].push(newLink);
              } else {
                 return; 
              }
           }
        });

        if(conflictFound) {
           renderLinkedEditorState(key);
           return; 
        }

        pendingLinkedEdits[key].push(newLink);
     }
     renderLinkedEditorState(key);
  }

  function saveLinkedEdit(table, id, key) {
     const btn = document.getElementById('btn_save_' + key);
     const originalText = btn.innerText;
     btn.innerText = "Saving..."; btn.disabled = true;

     const operations = [];

     for (const [otherKey, idsToRemove] of Object.entries(pendingRemovals)) {
         const current = currentPanelData[otherKey] || [];
         const kept = current.filter(l => !idsToRemove.includes(l.id)).map(l => l.id);
         if(idsToRemove.length > 0) {
            operations.push({ field: otherKey, val: kept });
         }
     }

     const finalIds = pendingLinkedEdits[key].map(l => l.id);
     operations.push({ field: key, val: finalIds });

     executeQueue(table, id, operations, function() {
         loadPanelRecord(table, id); 
     });
  }

  function executeQueue(table, id, ops, callback) {
     if(ops.length === 0) {
        callback();
        return;
     }
     const currentOp = ops.shift(); 
     google.script.run.withSuccessHandler(function() {
        executeQueue(table, id, ops, callback); 
     }).updateRecord(table, id, currentOp.field, currentOp.val);
  }

  // --- END LINKED EDITOR ---

  // ... (Spouse Modal Logic & General Logic remains the same) ...
  function openSpouseModal() {
     const f = currentContactRecord.fields;
     const spouseName = (f['Spouse Name'] && f['Spouse Name'].length > 0) ? f['Spouse Name'][0] : null;
     const spouseId = (f['Spouse'] && f['Spouse'].length > 0) ? f['Spouse'][0] : null;
     document.getElementById('spouseModal').style.display = 'block';
     document.getElementById('connectForm').style.display = 'none';
     document.getElementById('confirmConnectForm').style.display = 'none';
     document.getElementById('disconnectForm').style.display = 'none';
     if (spouseId) {
        document.getElementById('disconnectForm').style.display = 'flex';
        document.getElementById('currentSpouseName').innerText = spouseName;
        document.getElementById('currentSpouseId').value = spouseId;
     } else {
        document.getElementById('connectForm').style.display = 'flex';
        document.getElementById('spouseSearchInput').value = '';
        document.getElementById('spouseSearchResults').innerHTML = '';
        document.getElementById('spouseSearchResults').style.display = 'none';
        loadRecentContactsForModal();
     }
  }
  function closeSpouseModal() { document.getElementById('spouseModal').style.display = 'none'; }
  function backToSearch() {
     document.getElementById('confirmConnectForm').style.display = 'none';
     document.getElementById('connectForm').style.display = 'flex';
  }
  function loadRecentContactsForModal() {
     const resultsDiv = document.getElementById('spouseSearchResults');
     const inputVal = document.getElementById('spouseSearchInput').value;
     if(inputVal.length > 0) return; 
     resultsDiv.style.display = 'block';
     resultsDiv.innerHTML = '<div style="padding:10px; color:#999; font-style:italic;">Loading recent...</div>';
     google.script.run.withSuccessHandler(function(records) {
         resultsDiv.innerHTML = '<div style="padding:5px 10px; font-size:10px; color:#999; text-transform:uppercase; font-weight:700;">Recently Modified</div>';
         if (!records || records.length === 0) { resultsDiv.innerHTML += '<div style="padding:8px; font-style:italic; color:#999;">No recent contacts</div>'; } else {
            records.forEach(r => {
               if(r.id === currentContactRecord.id) return;
               renderSearchResultItem(r, resultsDiv);
            });
         }
     }).getRecentContacts();
  }
  function handleSpouseSearch(event) {
     const query = event.target.value;
     const resultsDiv = document.getElementById('spouseSearchResults');
     clearTimeout(spouseSearchTimeout);
     if(query.length === 0) { loadRecentContactsForModal(); return; }
     resultsDiv.style.display = 'block';
     resultsDiv.innerHTML = '<div style="padding:10px; color:#999; font-style:italic;">Searching...</div>';
     spouseSearchTimeout = setTimeout(() => {
        google.script.run.withSuccessHandler(function(records) {
           resultsDiv.innerHTML = '';
           if (records.length === 0) { resultsDiv.innerHTML = '<div style="padding:8px; font-style:italic; color:#999;">No results</div>'; } else {
              records.forEach(r => {
                 if(r.id === currentContactRecord.id) return;
                 renderSearchResultItem(r, resultsDiv);
              });
           }
        }).searchContacts(query);
     }, 500);
  }
  function renderSearchResultItem(r, container) {
     const name = formatName(r.fields);
     const details = formatDetails(r.fields); 
     const div = document.createElement('div');
     div.className = 'search-option';
     div.innerHTML = `<span style="font-weight:700; display:block;">${name}</span><span style="font-size:11px; color:#666;">${details}</span>`;
     div.onclick = function() {
        document.getElementById('targetSpouseName').innerText = name;
        document.getElementById('targetSpouseId').value = r.id;
        document.getElementById('connectForm').style.display = 'none';
        document.getElementById('confirmConnectForm').style.display = 'flex';
     };
     container.appendChild(div);
  }
  function executeSpouseChange(action) {
     const myId = currentContactRecord.id;
     let statusStr = ""; let otherId = ""; let expectHasSpouse = false; 
     if (action === 'disconnect') {
        statusStr = "disconnected as spouse from";
        otherId = document.getElementById('currentSpouseId').value;
        expectHasSpouse = false;
     } else {
        statusStr = "connected as spouse to";
        otherId = document.getElementById('targetSpouseId').value;
        expectHasSpouse = true;
     }
     closeSpouseModal();
     const statusEl = document.getElementById('spouseStatusText');
     statusEl.innerHTML = `<span style="color:var(--color-star); font-style:italic; font-weight:700; display:inline-flex; align-items:center;">Updating <span class="pulse-dot"></span><span class="pulse-dot"></span><span class="pulse-dot"></span></span>`;
     document.getElementById('spouseEditLink').style.display = 'none'; 
     google.script.run.withSuccessHandler(function(res) {
        pollAttempts = 0; startPolling(myId, expectHasSpouse);
     }).setSpouseStatus(myId, otherId, statusStr);
  }
  function startPolling(contactId, expectHasSpouse) {
     if(pollInterval) clearInterval(pollInterval);
     pollInterval = setInterval(() => {
        pollAttempts++;
        if (pollAttempts > 20) { 
           clearInterval(pollInterval);
           const statusEl = document.getElementById('spouseStatusText');
           if(statusEl) { statusEl.innerHTML = `<span style="color:#A00;">Update delayed.</span> <a class="data-link" onclick="forceReload('${contactId}')">Refresh</a>`; }
           return;
        }
        google.script.run.withSuccessHandler(function(r) {
           if(r && r.fields) {
              const currentSpouseId = (r.fields['Spouse'] && r.fields['Spouse'].length > 0) ? r.fields['Spouse'][0] : null;
              const match = expectHasSpouse ? (currentSpouseId !== null) : (currentSpouseId === null);
              if (match) {
                 clearInterval(pollInterval);
                 if(currentContactRecord && currentContactRecord.id === contactId) { selectContact(r); }
              }
           }
        }).getContactById(contactId); 
     }, 2000); 
  }
  function forceReload(id) {
     clearInterval(pollInterval);
     google.script.run.withSuccessHandler(function(r) { if(r && r.fields) selectContact(r); }).getContactById(id);
  }
  function resetForm() {
    toggleProfileView(true); document.getElementById('contactForm').reset();
    document.getElementById('recordId').value = ""; enableEditMode();
    document.getElementById('formTitle').innerText = "New Contact";
    document.getElementById('submitBtn').innerText = "Save Contact";
    document.getElementById('editBtn').style.visibility = 'hidden';
    document.getElementById('oppList').innerHTML = '<li style="color:#CCC; font-size:12px; font-style:italic;">No opportunities linked.</li>';
    document.getElementById('auditSection').style.display = 'none';
    document.getElementById('duplicateWarningBox').style.display = 'none';
    document.getElementById('spouseStatusText').innerHTML = "Single";
    document.getElementById('spouseHistoryList').innerHTML = "";
    document.getElementById('spouseEditLink').style.display = 'inline';
    document.getElementById('refreshBtn').style.display = 'none';
    closeOppPanel();
  }
  function handleSearch(event) {
    const query = event.target.value; const status = document.getElementById('searchStatus');
    clearTimeout(loadingTimer); 
    if(query.length === 0) { status.innerText = ""; loadContacts(); return; }
    clearTimeout(searchTimeout); status.innerText = "Typing...";
    searchTimeout = setTimeout(() => {
      status.innerText = "Searching...";
      google.script.run.withSuccessHandler(function(records) {
         status.innerText = records.length > 0 ? `Found ${records.length} matches` : "No matches found";
         renderList(records);
      }).searchContacts(query);
    }, 500);
  }
  function loadContacts() {
    const loadingDiv = document.getElementById('loading'); const list = document.getElementById('contactList');
    list.innerHTML = ''; loadingDiv.style.display = 'block'; loadingDiv.innerHTML = 'Loading directory...';
    clearTimeout(loadingTimer);

    // --- RESTORED CORRECT MESSAGE ---
    loadingTimer = setTimeout(() => { 
       loadingDiv.innerHTML = `
         <div style="margin-top:15px; text-align:center;">
           <button onclick="loadContacts()" class="wake-btn">Wake up Google!</button>
           <p class="wake-note">Google isn't constantly awake in the background waiting for us to use this site, so it goes to sleep if we haven't used it for a little while. You might need to hit the button a few times to get it to pay attention. One day we'll make it work differently so that this problem goes away.</p>
         </div>
       `; 
    }, 4000);

    google.script.run.withSuccessHandler(function(records) {
         clearTimeout(loadingTimer); document.getElementById('loading').style.display = 'none';
         if (!records || records.length === 0) { 
           list.innerHTML = '<li style="padding:20px; color:#999; text-align:center; font-size:13px;">No contacts found</li>'; 
           return; 
         }
         renderList(records);
      }).getRecentContacts();
  }
  function renderList(records) {
    const list = document.getElementById('contactList'); 
    document.getElementById('loading').style.display = 'none'; 
    list.innerHTML = '';
    records.forEach(record => {
      const f = record.fields; const item = document.createElement('li'); item.className = 'contact-item';
      item.innerHTML = `<span class="contact-name">${formatName(f)}</span><span class="contact-detail">${formatDetails(f)}</span>`;
      item.onclick = function() { selectContact(record); }; list.appendChild(item);
    });
  }
  function formatName(f) {
    let n = `${f.FirstName || ''} ${f.MiddleName || ''} ${f.LastName || ''}`.replace(/\s+/g, ' ').trim();
    if (f.PreferredName) n += ` (${f.PreferredName})`; return n;
  }
  function formatDetails(f) { return [f.EmailAddress1, f.Mobile].filter(Boolean).join(" • "); }

  // --- CORRECTED HISTORY DATE LOGIC ---
  function renderHistory(f) {
    const section = document.getElementById('auditSection');
    section.innerHTML = ''; section.style.display = 'block';
    const createdStr = f.Created || "";
    const dateMatch = createdStr.match(/(\d{2}):(\d{2})\s+(\d{2})\/(\d{2})\/(\d{4})/);
    let durationText = "unavailable";

    if (dateMatch) {
       const hours = parseInt(dateMatch[1], 10);
       const minutes = parseInt(dateMatch[2], 10);
       const day = parseInt(dateMatch[3], 10);
       const month = parseInt(dateMatch[4], 10) - 1; 
       const year = parseInt(dateMatch[5], 10);
       const createdDate = new Date(year, month, day, hours, minutes);
       const now = new Date();
       const diffMs = now - createdDate;
       const diffMinsTotal = Math.floor(diffMs / (1000 * 60));
       const diffHoursTotal = Math.floor(diffMs / (1000 * 60 * 60));
       const diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));

       if (diffDays > 730) { const years = Math.floor(diffDays / 365); durationText = `over ${years} years`; } 
       else if (diffDays > 60) { const months = Math.floor(diffDays / 30); durationText = `over ${months} months`; } 
       else if (diffDays >= 1) { durationText = (diffDays === 1) ? "1 day" : `${diffDays} days`; } 
       else {
          const mins = diffMinsTotal % 60;
          const hStr = (diffHoursTotal === 1) ? "hr" : "hrs";
          durationText = `${diffHoursTotal} ${hStr} and ${mins} minutes`;
       }
    }
    const nameToUse = f.PreferredName || f.FirstName || 'This client';
    const line1 = document.createElement('div'); line1.className = 'audit-text';
    line1.innerText = `${nameToUse} has been a contact in our database for ${durationText}.`;
    section.appendChild(line1);

    if (f.Created) { const line2 = document.createElement('div'); line2.className = 'audit-modified'; line2.innerText = f.Created; section.appendChild(line2); }
    if (f.Modified) { const line3 = document.createElement('div'); line3.className = 'audit-modified'; line3.innerText = f.Modified; section.appendChild(line3); }
  }

  function loadOpportunities(f) {
    const oppList = document.getElementById('oppList'); const loader = document.getElementById('oppLoading');
    document.getElementById('oppSortBtn').style.display = 'none'; oppList.innerHTML = ''; loader.style.display = 'block';
    let oppsToFetch = []; let roleMap = {};
    const addIds = (ids, roleName) => { if(!ids) return; (Array.isArray(ids) ? ids : [ids]).forEach(id => { oppsToFetch.push(id); roleMap[id] = roleName; }); };
    addIds(f['Opportunities - Primary Applicant'], 'Primary Applicant');
    addIds(f['Opportunities - Applicant'], 'Applicant');
    addIds(f['Opportunities - Guarantor'], 'Guarantor');
    if(oppsToFetch.length === 0) { loader.style.display = 'none'; oppList.innerHTML = '<li style="color:#CCC; font-size:12px; font-style:italic;">No opportunities linked.</li>'; return; }
    google.script.run.withSuccessHandler(function(oppRecords) {
       loader.style.display = 'none';
       oppRecords.forEach(r => r._role = roleMap[r.id] || "Linked");
       if(oppRecords.length > 1) { document.getElementById('oppSortBtn').style.display = 'inline'; }
       currentOppRecords = oppRecords; renderOppList();
    }).getLinkedOpportunities(oppsToFetch);
  }
  function toggleOppSort() {
     if(currentOppSortDirection === 'asc') currentOppSortDirection = 'desc'; else currentOppSortDirection = 'asc';
     renderOppList();
  }
  function renderOppList() {
     const oppList = document.getElementById('oppList'); oppList.innerHTML = '';
     const sorted = [...currentOppRecords].sort((a, b) => {
         const nameA = (a.fields['Opportunity Name'] || "").toLowerCase();
         const nameB = (b.fields['Opportunity Name'] || "").toLowerCase();
         if(currentOppSortDirection === 'asc') return nameA.localeCompare(nameB); return nameB.localeCompare(nameA);
     });
     sorted.forEach(opp => {
         const fields = opp.fields; const name = fields['Opportunity Name'] || "Unnamed Opportunity"; const role = opp._role;
         const li = document.createElement('li'); li.className = 'opp-item';
         li.innerHTML = `<span class="opp-title">${name}</span> <span class="opp-role">${role}</span>`;
         li.onclick = function() { panelHistory = []; loadPanelRecord('Opportunities', opp.id); }; oppList.appendChild(li);
     });
  }
  function handleFormSubmit(formObject) {
    event.preventDefault();
    const btn = document.getElementById('submitBtn'); const status = document.getElementById('status');
    btn.disabled = true; btn.innerText = "Saving...";
    google.script.run.withSuccessHandler(function(response) {
         status.innerText = "✅ " + response; status.className = "status-success";
         loadContacts(); if(!document.getElementById('recordId').value) resetForm();
         btn.disabled = false; btn.innerText = "Update Contact"; disableEditMode(); 
         setTimeout(() => { status.innerText = ""; status.className = ""; }, 3000);
      }).withFailureHandler(function(err) { status.innerText = "❌ " + err.message; status.className = "status-error"; btn.disabled = false; btn.innerText = "Try Again"; }).processForm(formObject);
  }
  function loadPanelRecord(table, id) {
    const panel = document.getElementById('oppDetailPanel'); const content = document.getElementById('panelContent');
    const titleEl = document.getElementById('panelTitle'); const backBtn = document.getElementById('panelBackBtn');
    panel.classList.add('open'); content.innerHTML = `<div style="text-align:center; color:#999; margin-top:50px;">Loading...</div>`;
    google.script.run.withSuccessHandler(function(response) {
      if (!response || !response.data) { content.innerHTML = "Error loading."; return; }

      currentPanelData = {};
      response.data.forEach(item => { if(item.type === 'link') currentPanelData[item.key] = item.value; });

      panelHistory.push({ table: table, id: id, title: response.title });
      updateBackButton(); titleEl.innerText = response.title;
      let html = '';
      response.data.forEach(item => {
         if (item.key === 'Opportunity Name') {
            const safeValue = (item.value || "").toString().replace(/"/g, "&quot;");
            html += `<div class="detail-group"><div class="detail-label">${item.label}</div><div id="view_${item.key}"><div class="detail-value" style="display:flex; justify-content:space-between; align-items:center;"><span id="display_${item.key}">${item.value}</span><span class="edit-field-icon" onclick="toggleFieldEdit('${item.key}')">✎</span></div></div><div id="edit_${item.key}" style="display:none;"><div class="edit-wrapper"><input type="text" id="input_${item.key}" value="${safeValue}" class="edit-input"><div class="edit-btn-row"><button onclick="cancelFieldEdit('${item.key}')" class="btn-cancel-field">Cancel</button><button id="btn_save_${item.key}" onclick="saveFieldEdit('${table}', '${id}', '${item.key}')" class="btn-save-field">Save</button></div></div></div></div>`;
            return;
         }
         if (['Primary Applicant', 'Applicants', 'Guarantors'].includes(item.key)) {
            let linkHtml = '';
            if (item.value.length === 0) linkHtml = '<span style="color:#CCC; font-style:italic;">None</span>';
            else item.value.forEach(link => { linkHtml += `<a class="data-link" onclick="loadPanelRecord('${link.table}', '${link.id}')">${link.name}</a>`; });
            html += `<div class="detail-group"><div class="detail-label">${item.label}</div><div id="view_${item.key}"><div class="detail-value" style="display:flex; justify-content:space-between; align-items:center;"><span>${linkHtml}</span><span class="edit-field-icon" onclick="toggleLinkedEdit('${item.key}')">✎</span></div></div><div id="edit_${item.key}" style="display:none;"><div class="edit-wrapper"><div id="chip_container_${item.key}" class="link-chip-container"></div><input type="text" placeholder="Add contact..." class="link-search-input" onkeyup="handleLinkedSearch(event, '${item.key}')"><div id="error_${item.key}" class="input-error"></div><div id="results_${item.key}" class="link-results"></div><div class="edit-btn-row" style="margin-top:10px;"><button onclick="closeLinkedEdit('${item.key}')" class="btn-cancel-field">Done</button></div></div></div></div>`;
            return;
         }
         if (item.type === 'link') {
            const links = item.value; let linkHtml = '';
            if (links.length === 0) linkHtml = '<span style="color:#CCC; font-style:italic;">None</span>';
            else { links.forEach(link => { linkHtml += `<a class="data-link" onclick="loadPanelRecord('${link.table}', '${link.id}')">${link.name}</a>`; }); }
            html += `<div class="detail-group"><div class="detail-label">${item.label}</div><div class="detail-value" style="border:none;">${linkHtml}</div></div>`;
         } else {
            let display = item.value; if (display === undefined || display === null || display === "") return;
            html += `<div class="detail-group"><div class="detail-label">${item.label}</div><div class="detail-value">${display}</div></div>`;
         }
      });
      content.innerHTML = html;
    }).getRecordDetail(table, id);
  }
  function popHistory() { if (panelHistory.length <= 1) return; panelHistory.pop(); const prev = panelHistory[panelHistory.length - 1]; panelHistory.pop(); loadPanelRecord(prev.table, prev.id); }
  function updateBackButton() { const btn = document.getElementById('panelBackBtn'); if (panelHistory.length > 1) { btn.style.display = 'block'; } else { btn.style.display = 'none'; } }
  function closeOppPanel() { document.getElementById('oppDetailPanel').classList.remove('open'); panelHistory = []; }