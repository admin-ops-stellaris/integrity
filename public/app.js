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

  window.onload = function() { 
    loadContacts(); 
    checkUserIdentity(); 
    initKeyboardShortcuts();
    initDarkMode();
    initScreensaver();
  };

  // --- KEYBOARD SHORTCUTS ---
  function initKeyboardShortcuts() {
    document.addEventListener('keydown', function(e) {
      const activeEl = document.activeElement;
      const isTyping = activeEl.tagName === 'INPUT' || activeEl.tagName === 'TEXTAREA' || activeEl.isContentEditable;
      
      if (e.key === 'Escape') {
        closeOppPanel();
        closeSpouseModal();
        closeNewOppModal();
        closeShortcutsModal();
        closeDeleteConfirmModal();
        closeAlertModal();
        if (document.getElementById('actionRow').style.display === 'flex') disableEditMode();
        return;
      }
      
      if (isTyping) return;
      
      if (e.key === '?') {
        e.preventDefault();
        showShortcutsHelp();
        return;
      }
      
      if (e.key === '/') {
        e.preventDefault();
        document.getElementById('searchInput').focus();
      } else if (e.key === 'n' || e.key === 'N') {
        e.preventDefault();
        resetForm();
      } else if (e.key === 'e' || e.key === 'E') {
        if (currentContactRecord && document.getElementById('editBtn').style.visibility !== 'hidden') {
          e.preventDefault();
          enableEditMode();
        }
      }
    });
  }

  // --- DARK MODE ---
  function initDarkMode() {
    const savedTheme = localStorage.getItem('integrity-theme');
    if (savedTheme === 'dark') document.body.classList.add('dark-mode');
  }
  function toggleDarkMode() {
    document.body.classList.toggle('dark-mode');
    const isDark = document.body.classList.contains('dark-mode');
    localStorage.setItem('integrity-theme', isDark ? 'dark' : 'light');
  }

  // --- SCREENSAVER ---
  let screensaverTimer = null;
  const SCREENSAVER_DELAY = 120000; // 2 minutes
  
  function initScreensaver() {
    function resetScreensaverTimer() {
      if (document.body.classList.contains('screensaver-active')) {
        document.body.classList.remove('screensaver-active');
      }
      clearTimeout(screensaverTimer);
      screensaverTimer = setTimeout(() => {
        document.body.classList.add('screensaver-active');
      }, SCREENSAVER_DELAY);
    }
    
    ['mousemove', 'mousedown', 'keydown', 'touchstart', 'scroll'].forEach(event => {
      document.addEventListener(event, resetScreensaverTimer, { passive: true });
    });
    
    resetScreensaverTimer();
  }

  // --- QUICK ADD OPPORTUNITY ---
  function quickAddOpportunity() {
    if (!currentContactRecord) { alert('Please select a contact first.'); return; }
    const f = currentContactRecord.fields;
    const contactName = formatName(f);
    const today = new Date().toLocaleDateString('en-AU', { day: '2-digit', month: '2-digit', year: 'numeric' });
    const defaultName = `${contactName} - ${today}`;
    
    const spouseName = (f['Spouse Name'] && f['Spouse Name'].length > 0) ? f['Spouse Name'][0] : null;
    const spouseId = (f['Spouse'] && f['Spouse'].length > 0) ? f['Spouse'][0] : null;
    
    openNewOppModal(defaultName, contactName, spouseName, spouseId);
  }
  
  function openNewOppModal(defaultName, contactName, spouseName, spouseId) {
    const modal = document.getElementById('newOppModal');
    document.getElementById('newOppName').value = defaultName;
    document.getElementById('newOppContactName').innerText = contactName;
    
    const spouseSection = document.getElementById('newOppSpouseSection');
    if (spouseName && spouseId) {
      document.getElementById('newOppSpouseName').innerText = spouseName;
      document.getElementById('addSpouseAsApplicant').checked = false;
      updateSpouseCheckboxLabel();
      spouseSection.style.display = 'block';
    } else {
      spouseSection.style.display = 'none';
    }
    
    openModal('newOppModal');
    setTimeout(() => {
      document.getElementById('newOppName').focus();
      document.getElementById('newOppName').select();
    }, 50);
  }
  
  function closeNewOppModal() {
    closeModal('newOppModal');
  }
  
  function updateSpouseCheckboxLabel() {
    const checkbox = document.getElementById('addSpouseAsApplicant');
    const prefix = document.getElementById('spouseCheckboxPrefix');
    const suffix = document.getElementById('spouseCheckboxSuffix');
    if (checkbox.checked) {
      prefix.innerText = 'Adding ';
      suffix.innerText = ' as Applicant';
    } else {
      prefix.innerText = 'Also add ';
      suffix.innerText = ' as Applicant?';
    }
  }
  
  function showShortcutsHelp() {
    openModal('shortcutsModal');
  }
  
  function closeShortcutsModal() {
    closeModal('shortcutsModal');
  }
  
  document.addEventListener('click', function(e) {
    const shortcutsModal = document.getElementById('shortcutsModal');
    if (shortcutsModal && e.target === shortcutsModal) closeShortcutsModal();
    const deleteConfirmModal = document.getElementById('deleteConfirmModal');
    if (deleteConfirmModal && e.target === deleteConfirmModal) closeDeleteConfirmModal();
    const alertModal = document.getElementById('alertModal');
    if (alertModal && e.target === alertModal) closeAlertModal();
  });
  
  function submitNewOpportunity() {
    const oppName = document.getElementById('newOppName').value.trim();
    if (!oppName) { alert('Please enter an opportunity name.'); return; }
    
    const oppType = document.getElementById('newOppType').value;
    const f = currentContactRecord.fields;
    const spouseId = (f['Spouse'] && f['Spouse'].length > 0) ? f['Spouse'][0] : null;
    const addSpouse = document.getElementById('addSpouseAsApplicant')?.checked && spouseId;
    
    closeNewOppModal();
    
    google.script.run.withSuccessHandler(function(res) {
      if (res && res.id) {
        const finishUp = () => {
          google.script.run.withSuccessHandler(function(updatedContact) {
            if (updatedContact) {
              currentContactRecord = updatedContact;
              loadOpportunities(updatedContact.fields);
            }
            setTimeout(() => loadPanelRecord('Opportunities', res.id), 300);
          }).getContactById(currentContactRecord.id);
        };
        
        if (addSpouse) {
          google.script.run.withSuccessHandler(finishUp).updateOpportunity(res.id, 'Applicants', [spouseId]);
        } else {
          finishUp();
        }
      }
    }).createOpportunity(oppName, currentContactRecord.id, oppType);
  }

  // --- CELEBRATION ---
  function triggerWonCelebration() {
    const container = document.createElement('div');
    container.className = 'confetti-container';
    document.body.appendChild(container);
    const colors = ['#BB9934', '#7B8B64', '#19414C', '#D0DFE6', '#F2F0E9'];
    for (let i = 0; i < 50; i++) {
      const confetti = document.createElement('div');
      confetti.className = 'confetti';
      confetti.style.left = Math.random() * 100 + '%';
      confetti.style.backgroundColor = colors[Math.floor(Math.random() * colors.length)];
      confetti.style.animationDelay = Math.random() * 0.5 + 's';
      confetti.style.animationDuration = (Math.random() * 1 + 1.5) + 's';
      container.appendChild(confetti);
    }
    setTimeout(() => container.remove(), 3000);
  }

  // --- AVATAR HELPERS ---
  function getInitials(firstName, lastName) {
    const f = (firstName || '').charAt(0).toUpperCase();
    const l = (lastName || '').charAt(0).toUpperCase();
    return f + l || '?';
  }
  function getAvatarColor(name) {
    const colors = ['#19414C', '#7B8B64', '#BB9934', '#2C2622', '#6B5B4F'];
    let hash = 0;
    for (let i = 0; i < name.length; i++) hash = name.charCodeAt(i) + ((hash << 5) - hash);
    return colors[Math.abs(hash) % colors.length];
  }

  function checkUserIdentity() {
    google.script.run.withSuccessHandler(function(email) {
       const display = email ? email : "Unknown";
       document.getElementById('debugUser').innerText = display;
       document.getElementById('userEmail').innerText = email || "Not signed in";
       if (!email) alert("Warning: The system cannot detect your email address.");
    }).getEffectiveUserEmail();
  }

  function updateHeaderTitle(isEditing) {
    const fName = document.getElementById('firstName').value || "";
    const mName = document.getElementById('middleName').value || "";
    const lName = document.getElementById('lastName').value || "";
    let fullName = [fName, mName, lName].filter(Boolean).join(" ");
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
      document.getElementById('formSubtitle').innerText = '';
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
    document.getElementById('cancelBtn').style.display = 'inline-block';
    document.getElementById('editBtn').style.visibility = 'hidden';
    updateHeaderTitle(true); 
  }

  function disableEditMode() {
    const inputs = document.querySelectorAll('#contactForm input, #contactForm textarea');
    inputs.forEach(input => { input.classList.add('locked'); input.readOnly = true; });
    document.getElementById('actionRow').style.display = 'none';
    document.getElementById('cancelBtn').style.display = 'none';
    updateHeaderTitle(false); 
  }
  
  function cancelEditMode() {
    const recordId = document.getElementById('recordId').value;
    document.getElementById('cancelBtn').style.display = 'none';
    if (recordId && currentContactRecord) {
      selectContact(currentContactRecord);
    } else {
      document.getElementById('contactForm').reset();
      toggleProfileView(false);
    }
    disableEditMode();
  }

  function selectContact(record) {
    document.getElementById('cancelBtn').style.display = 'none';
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

    document.getElementById('formSubtitle').innerText = formatSubtitle(f);
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

  function confirmDeleteContact() {
    if (!currentContactRecord) return;
    const f = currentContactRecord.fields;
    const name = formatName(f);
    
    document.getElementById('deleteConfirmMessage').innerText = `Are you sure you want to delete "${name}"? This action cannot be undone.`;
    openModal('deleteConfirmModal');
  }
  
  function closeDeleteConfirmModal() {
    closeModal('deleteConfirmModal');
  }
  
  function executeDeleteContact() {
    if (!currentContactRecord) return;
    const f = currentContactRecord.fields;
    const name = formatName(f);
    const contactId = currentContactRecord.id;
    
    closeModal('deleteConfirmModal', function() {
      google.script.run
        .withSuccessHandler(function(result) {
          if (result.success) {
            showAlert('Success', `"${name}" has been deleted.`, 'success');
            currentContactRecord = null;
            toggleProfileView(false);
            loadContacts();
          } else {
            showAlert('Cannot Delete', result.error || 'Failed to delete contact.', 'error');
          }
        })
        .withFailureHandler(function(err) {
          showAlert('Error', err.message, 'error');
        })
        .deleteContact(contactId);
    });
  }
  
  function openModal(modalId) {
    const modal = document.getElementById(modalId);
    modal.classList.add('visible');
    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        modal.classList.add('showing');
      });
    });
  }
  
  function closeModal(modalId, callback) {
    const modal = document.getElementById(modalId);
    modal.classList.remove('showing');
    setTimeout(() => {
      modal.classList.remove('visible');
      if (callback) callback();
    }, 250);
  }
  
  function showAlert(title, message, type) {
    const modal = document.getElementById('alertModal');
    const sidebar = document.getElementById('alertModalSidebar');
    const icon = document.getElementById('alertModalIcon');
    
    document.getElementById('alertModalTitle').innerText = title;
    document.getElementById('alertModalMessage').innerText = message;
    
    if (type === 'success') {
      sidebar.style.background = 'var(--color-cedar)';
      icon.innerText = '✓';
    } else if (type === 'error') {
      sidebar.style.background = '#A00';
      icon.innerText = '✕';
    } else {
      sidebar.style.background = 'var(--color-star)';
      icon.innerText = 'ℹ';
    }
    
    openModal('alertModal');
  }
  
  function closeAlertModal() {
    closeModal('alertModal');
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
        if(fieldKey === 'Status' && val === 'Won') {
           triggerWonCelebration();
        }
        if(fieldKey === 'Status' && currentContactRecord) {
           loadOpportunities(currentContactRecord.fields);
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
     openModal('spouseModal');
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
  function closeSpouseModal() { closeModal('spouseModal'); }
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
    document.getElementById('formSubtitle').innerText = '';
    document.getElementById('submitBtn').innerText = "Save Contact";
    document.getElementById('cancelBtn').style.display = 'inline-block';
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
      const fullName = formatName(f);
      const initials = getInitials(f.FirstName, f.LastName);
      const avatarColor = getAvatarColor(fullName);
      const modifiedTooltip = formatModifiedTooltip(f);
      const modifiedShort = formatModifiedShort(f);
      item.innerHTML = `<div class="contact-avatar" style="background-color:${avatarColor}">${initials}</div><div class="contact-info"><span class="contact-name">${fullName}</span><div class="contact-details-row">${formatDetailsRow(f)}</div></div>${modifiedShort ? `<span class="contact-modified" title="${modifiedTooltip || ''}">${modifiedShort}</span>` : ''}`;
      item.onclick = function() { selectContact(record); }; list.appendChild(item);
    });
  }
  function formatName(f) {
    return `${f.FirstName || ''} ${f.MiddleName || ''} ${f.LastName || ''}`.replace(/\s+/g, ' ').trim();
  }
  function formatDetailsRow(f) {
    const parts = [];
    if (f.EmailAddress1) parts.push(`<span>${f.EmailAddress1}</span>`);
    if (f.Mobile) parts.push(`<span>${f.Mobile}</span>`);
    return parts.join('');
  }
  function formatModifiedTooltip(f) {
    const modifiedOn = f['Modified On'];
    const modifiedBy = f['Last Site User Name'];
    if (!modifiedOn) return null;
    
    const dateMatch = modifiedOn.match(/(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/);
    if (!dateMatch) return null;
    
    const modDate = new Date(dateMatch[1], dateMatch[2] - 1, dateMatch[3], dateMatch[4], dateMatch[5]);
    const now = new Date();
    const diffMs = now - modDate;
    const diffMins = Math.floor(diffMs / (1000 * 60));
    const diffHours = Math.floor(diffMs / (1000 * 60 * 60));
    const diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));
    
    let timeAgo;
    if (diffMins < 1) timeAgo = 'just now';
    else if (diffMins < 60) timeAgo = `${diffMins} min${diffMins > 1 ? 's' : ''} ago`;
    else if (diffHours < 24) timeAgo = `${diffHours} hr${diffHours > 1 ? 's' : ''} ago`;
    else if (diffDays === 1) timeAgo = 'yesterday';
    else if (diffDays < 7) timeAgo = `${diffDays} days ago`;
    else timeAgo = modDate.toLocaleDateString('en-AU');
    
    if (modifiedBy) return `Modified ${timeAgo} by ${modifiedBy}`;
    return `Modified ${timeAgo}`;
  }
  function formatModifiedShort(f) {
    const modifiedOn = f['Modified On'];
    if (!modifiedOn) return null;
    
    const dateMatch = modifiedOn.match(/(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/);
    if (!dateMatch) return null;
    
    const modDate = new Date(dateMatch[1], dateMatch[2] - 1, dateMatch[3], dateMatch[4], dateMatch[5]);
    const now = new Date();
    const diffMs = now - modDate;
    const diffMins = Math.floor(diffMs / (1000 * 60));
    const diffHours = Math.floor(diffMs / (1000 * 60 * 60));
    const diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));
    
    if (diffMins < 1) return 'now';
    if (diffMins < 60) return `${diffMins}m`;
    if (diffHours < 24) return `${diffHours}h`;
    if (diffDays < 7) return `${diffDays}d`;
    return modDate.toLocaleDateString('en-AU', { day: 'numeric', month: 'short' });
  }
  function formatSubtitle(f) {
    const preferredName = f.PreferredName || f.FirstName || '';
    const tenure = calculateTenure(f.Created);
    if (!preferredName && !tenure) return '';
    const parts = [];
    if (preferredName) parts.push(`prefers ${preferredName}`);
    if (tenure) parts.push(`in our database for ${tenure}`);
    return parts.join(' · ');
  }
  function calculateTenure(createdStr) {
    if (!createdStr) return null;
    const dateMatch = createdStr.match(/(\d{2}):(\d{2})\s+(\d{2})\/(\d{2})\/(\d{4})/);
    if (!dateMatch) return null;
    const day = parseInt(dateMatch[3], 10);
    const month = parseInt(dateMatch[4], 10) - 1;
    const year = parseInt(dateMatch[5], 10);
    const createdDate = new Date(year, month, day);
    const diffMs = new Date() - createdDate;
    const diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));
    if (diffDays > 730) return `${Math.floor(diffDays / 365)}+ years`;
    if (diffDays > 60) return `${Math.floor(diffDays / 30)}+ months`;
    if (diffDays >= 1) return diffDays === 1 ? "1 day" : `${diffDays} days`;
    return "today";
  }

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
    if (f.Created) { const line1 = document.createElement('div'); line1.className = 'audit-modified'; line1.innerText = f.Created; section.appendChild(line1); }
    if (f.Modified) { const line2 = document.createElement('div'); line2.className = 'audit-modified'; line2.innerText = f.Modified; section.appendChild(line2); }
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
         const status = fields['Status'] || '';
         const oppType = fields['Opportunity Type'] || '';
         const statusClass = status === 'Won' ? 'status-won' : status === 'Lost' ? 'status-lost' : '';
         const li = document.createElement('li'); li.className = `opp-item ${statusClass}`;
         const statusBadge = status ? `<span class="opp-status-badge ${statusClass}">${status}</span>` : '';
         const typeLabel = oppType ? `<span class="opp-type">${oppType}</span>` : '';
         li.innerHTML = `<span class="opp-title">${name}${typeLabel}</span><span class="opp-role-wrapper">${statusBadge}<span class="opp-role">${role}</span></span>`;
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
         if (item.type === 'select') {
            const currentVal = item.value || '';
            const options = item.options || [];
            let optionsHtml = options.map(opt => `<option value="${opt}" ${opt === currentVal ? 'selected' : ''}>${opt}</option>`).join('');
            html += `<div class="detail-group"><div class="detail-label">${item.label}</div><div id="view_${item.key}"><div class="detail-value" style="display:flex; justify-content:space-between; align-items:center;"><span id="display_${item.key}">${currentVal || '<span style="color:#CCC; font-style:italic;">Not set</span>'}</span><span class="edit-field-icon" onclick="toggleFieldEdit('${item.key}')">✎</span></div></div><div id="edit_${item.key}" style="display:none;"><div class="edit-wrapper"><select id="input_${item.key}" class="edit-input">${optionsHtml}</select><div class="edit-btn-row"><button onclick="cancelFieldEdit('${item.key}')" class="btn-cancel-field">Cancel</button><button id="btn_save_${item.key}" onclick="saveFieldEdit('${table}', '${id}', '${item.key}')" class="btn-save-field">Save</button></div></div></div></div>`;
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