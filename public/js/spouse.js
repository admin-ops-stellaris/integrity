/**
 * Spouse Module
 * Spouse section rendering, modal, and history
 */
(function() {
  'use strict';
  
  const state = window.IntegrityState;
  
  // ============================================================
  // Render Spouse Section
  // ============================================================
  
  window.renderSpouseSection = function(f) {
    const statusText = document.getElementById('spouseStatusText');
    const historyAccordion = document.getElementById('spouseHistoryAccordion');
    const historyList = document.getElementById('spouseHistoryList');
    const historyDate = document.getElementById('spouseHistoryDate');
    const editLink = document.getElementById('spouseEditLink');
    
    const spouseId = f.Spouse?.[0] || null;
    const spouseName = f['Spouse Calculated Name'] || null;
    const historyText = f['Spouse History Text'] || '';
    const relationshipType = f.RelationshipType || '';
    
    // Parse history entries
    const historyEntries = historyText
      .split('\n')
      .filter(line => line.trim())
      .map(line => parseSpouseHistoryEntry(line))
      .filter(Boolean);
    
    // Show/hide history accordion
    if (historyEntries.length > 0) {
      historyAccordion.style.display = 'flex';
      historyList.innerHTML = '';
      historyEntries.forEach(entry => renderHistoryItem(entry, historyList));
    } else {
      historyAccordion.style.display = 'none';
      historyList.innerHTML = '';
    }
    
    // Update status text
    if (spouseId && spouseName) {
      const displayType = relationshipType || 'Married';
      statusText.innerHTML = `
        <span class="spouse-name" onclick="loadContactById('${spouseId}', true)">${escapeHtml(spouseName)}</span>
      `;
      historyDate.textContent = displayType;
    } else {
      statusText.innerHTML = 'Single';
      historyDate.textContent = '';
    }
    
    // Show edit link
    editLink.style.display = 'inline';
    
    // Collapse history by default
    historyList.style.display = 'none';
    const arrow = document.getElementById('spouseHistoryArrow');
    if (arrow) arrow.classList.remove('expanded');
  };
  
  // ============================================================
  // Spouse History Parsing
  // ============================================================
  
  window.parseSpouseHistoryEntry = function(logString) {
    const match = logString.match(/^(.+?) - (.+?) \((.+?)\) by (.+)$/);
    if (match) {
      return { action: match[1], name: match[2], date: match[3], by: match[4] };
    }
    return null;
  };
  
  function renderHistoryItem(entry, container) {
    const li = document.createElement('li');
    li.className = 'spouse-history-item';
    li.innerHTML = `
      <span class="history-action">${escapeHtml(entry.action)}</span>
      <span class="history-name">${escapeHtml(entry.name)}</span>
      <span class="history-date">${escapeHtml(entry.date)}</span>
      <span class="history-by">by ${escapeHtml(entry.by)}</span>
    `;
    container.appendChild(li);
  }
  
  // ============================================================
  // Toggle Spouse History
  // ============================================================
  
  window.toggleSpouseHistory = function() {
    const historyList = document.getElementById('spouseHistoryList');
    const arrow = document.getElementById('spouseHistoryArrow');
    
    if (historyList.style.display === 'none') {
      historyList.style.display = 'block';
      arrow?.classList.add('expanded');
    } else {
      historyList.style.display = 'none';
      arrow?.classList.remove('expanded');
    }
  };
  
  // ============================================================
  // Spouse Modal
  // ============================================================
  
  window.openSpouseModal = function() {
    const modal = document.getElementById('spouseModal');
    const searchInput = document.getElementById('spouseSearchInput');
    const resultsList = document.getElementById('spouseSearchResults');
    const step1 = document.getElementById('spouseStep1');
    const step2 = document.getElementById('spouseStep2');
    
    if (searchInput) searchInput.value = '';
    if (resultsList) resultsList.innerHTML = '';
    if (step1) step1.style.display = 'block';
    if (step2) step2.style.display = 'none';
    
    openModal('spouseModal');
    loadRecentContactsForModal();
    
    setTimeout(() => searchInput?.focus(), 100);
  };
  
  window.closeSpouseModal = function() {
    closeModal('spouseModal');
  };
  
  window.backToSearch = function() {
    document.getElementById('spouseStep1').style.display = 'block';
    document.getElementById('spouseStep2').style.display = 'none';
  };
  
  // ============================================================
  // Recent Contacts for Modal
  // ============================================================
  
  function loadRecentContactsForModal() {
    google.script.run.withSuccessHandler(function(records) {
      const container = document.getElementById('spouseSearchResults');
      if (!container) return;
      
      container.innerHTML = '';
      records.forEach(r => renderSearchResultItem(r, container));
    }).getRecentContacts();
  }
  
  // ============================================================
  // Spouse Search
  // ============================================================
  
  window.handleSpouseSearch = function(event) {
    const query = event.target.value;
    const container = document.getElementById('spouseSearchResults');
    
    clearTimeout(state.spouseSearchTimeout);
    
    if (query.length === 0) {
      loadRecentContactsForModal();
      return;
    }
    
    state.spouseSearchTimeout = setTimeout(() => {
      google.script.run.withSuccessHandler(function(records) {
        container.innerHTML = '';
        records.forEach(r => renderSearchResultItem(r, container));
      }).searchContacts(query);
    }, 300);
  };
  
  function renderSearchResultItem(r, container) {
    const f = r.fields;
    const name = formatName(f);
    const details = formatDetailsRow(f);
    
    // Skip current contact
    if (state.currentContactRecord && r.id === state.currentContactRecord.id) return;
    
    // Skip current spouse
    const currentSpouseId = state.currentContactRecord?.fields?.Spouse?.[0];
    if (r.id === currentSpouseId) return;
    
    const div = document.createElement('div');
    div.className = 'spouse-search-result';
    div.innerHTML = `
      <div class="spouse-result-name">${escapeHtml(name)}</div>
      <div class="spouse-result-details">${escapeHtml(details)}</div>
    `;
    div.onclick = () => selectSpouse(r);
    container.appendChild(div);
  }
  
  function selectSpouse(record) {
    const step1 = document.getElementById('spouseStep1');
    const step2 = document.getElementById('spouseStep2');
    const selectedName = document.getElementById('selectedSpouseName');
    const relationshipSelect = document.getElementById('relationshipType');
    
    step1.style.display = 'none';
    step2.style.display = 'block';
    
    selectedName.textContent = formatName(record.fields);
    selectedName.dataset.spouseId = record.id;
    
    if (relationshipSelect) relationshipSelect.value = 'Married';
  }
  
  // ============================================================
  // Execute Spouse Change
  // ============================================================
  
  window.executeSpouseChange = function(action) {
    const currentId = state.currentContactRecord?.id;
    if (!currentId) return;
    
    if (action === 'link') {
      const spouseId = document.getElementById('selectedSpouseName')?.dataset?.spouseId;
      const relationshipType = document.getElementById('relationshipType')?.value || 'Married';
      
      if (!spouseId) return;
      
      closeSpouseModal();
      document.getElementById('spouseStatusText').innerHTML = '<span style="color:#888;">Linking...</span>';
      
      google.script.run.withSuccessHandler(function(result) {
        if (result.success) {
          startPolling(currentId, true);
        } else {
          document.getElementById('spouseStatusText').innerHTML = '<span style="color:#A00;">Link failed</span>';
        }
      }).linkSpouse(currentId, spouseId, relationshipType);
      
    } else if (action === 'unlink') {
      closeSpouseModal();
      document.getElementById('spouseStatusText').innerHTML = '<span style="color:#888;">Unlinking...</span>';
      
      google.script.run.withSuccessHandler(function(result) {
        if (result.success) {
          startPolling(currentId, false);
        } else {
          document.getElementById('spouseStatusText').innerHTML = '<span style="color:#A00;">Unlink failed</span>';
        }
      }).unlinkSpouse(currentId);
    }
  };
  
  // ============================================================
  // Polling for Spouse Updates
  // ============================================================
  
  function startPolling(contactId, expectHasSpouse) {
    if (state.pollInterval) clearInterval(state.pollInterval);
    state.pollAttempts = 0;
    
    state.pollInterval = setInterval(() => {
      state.pollAttempts++;
      
      if (state.pollAttempts > 20) {
        clearInterval(state.pollInterval);
        const statusEl = document.getElementById('spouseStatusText');
        if (statusEl) {
          statusEl.innerHTML = `<span style="color:#A00;">Update delayed.</span> <a class="data-link" onclick="forceReload('${contactId}')">Refresh</a>`;
        }
        return;
      }
      
      google.script.run.withSuccessHandler(function(r) {
        if (r && r.fields) {
          const currentSpouseId = (r.fields['Spouse'] && r.fields['Spouse'].length > 0) ? r.fields['Spouse'][0] : null;
          const match = expectHasSpouse ? (currentSpouseId !== null) : (currentSpouseId === null);
          if (match) {
            clearInterval(state.pollInterval);
            if (state.currentContactRecord && state.currentContactRecord.id === contactId) {
              selectContact(r);
            }
          }
        }
      }).getContactById(contactId);
    }, 2000);
  }
  
  window.forceReload = function(id) {
    clearInterval(state.pollInterval);
    google.script.run.withSuccessHandler(function(r) {
      if (r && r.fields) selectContact(r);
    }).getContactById(id);
  };
  
})();
