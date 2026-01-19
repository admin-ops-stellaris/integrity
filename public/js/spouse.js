/**
 * Spouse Module
 * Handles spouse relationship management, history rendering, and modal logic
 */
(function() {
  'use strict';
  
  const state = window.IntegrityState;
  
  function renderSpouseSection(f) {
    const badgeEl = document.getElementById('spouseBadge');
    const statusEl = document.getElementById('spouseStatusText');
    const dateEl = document.getElementById('spouseHistoryDate');
    const accordionEl = document.getElementById('spouseHistoryAccordion');
    const historyList = document.getElementById('spouseHistoryList');
    const arrowEl = document.getElementById('spouseHistoryArrow');

    const spouseName = (f['Spouse Name'] && f['Spouse Name'].length > 0) ? f['Spouse Name'][0] : null;
    const spouseId = (f['Spouse'] && f['Spouse'].length > 0) ? f['Spouse'][0] : null;

    if (accordionEl) accordionEl.style.display = 'none';
    if (historyList) { historyList.innerHTML = ''; historyList.style.display = 'none'; }
    if (arrowEl) arrowEl.classList.remove('expanded');

    let connectionDate = '';
    const rawLogs = f['Spouse History Text']; 
    let parsedLogs = [];
    if (rawLogs && Array.isArray(rawLogs) && rawLogs.length > 0) {
      parsedLogs = rawLogs.map(parseSpouseHistoryEntry).filter(Boolean);
      parsedLogs.sort((a, b) => b.timestamp - a.timestamp);
      const connLog = parsedLogs.find(e => e.displayText.toLowerCase().includes('connected to'));
      if (connLog) connectionDate = connLog.displayDate;
    }

    if (spouseName && spouseId) {
      if (badgeEl) badgeEl.style.display = 'inline-block';
      statusEl.innerHTML = spouseName;
      statusEl.className = 'connection-name';
      statusEl.setAttribute('data-contact-id', spouseId);
      if (typeof attachQuickViewToElement === 'function') {
        attachQuickViewToElement(statusEl, spouseId);
      }
      
      if (parsedLogs.length > 1 && accordionEl && historyList) {
        if (dateEl) dateEl.textContent = '';
        accordionEl.style.display = 'inline-flex';
        parsedLogs.forEach(entry => { renderHistoryItem(entry, historyList); });
      } else {
        if (dateEl) dateEl.textContent = connectionDate;
      }
    } else {
      if (badgeEl) badgeEl.style.display = 'none';
      statusEl.innerHTML = "Single";
      statusEl.className = 'connection-name single';
      statusEl.removeAttribute('data-contact-id');
      if (dateEl) dateEl.textContent = '';
    }
  }
  
  function toggleSpouseHistory() {
    const historyList = document.getElementById('spouseHistoryList');
    const arrowEl = document.getElementById('spouseHistoryArrow');
    if (historyList && arrowEl) {
      const isExpanded = arrowEl.classList.contains('expanded');
      historyList.style.display = isExpanded ? 'none' : 'block';
      arrowEl.classList.toggle('expanded');
    }
  }
  
  function parseSpouseHistoryEntry(logString) {
    const match = logString.match(/^(\d{4})-(\d{2})-(\d{2})\s+(\d{2}):(\d{2}):(\d{2}):\s*(connected as spouse to|disconnected as spouse from)\s+(.+)$/);
    if (!match) return null;
    const [, year, month, day, hours, mins, secs, action, spouseName] = match;
    const timestamp = new Date(parseInt(year), parseInt(month) - 1, parseInt(day), parseInt(hours), parseInt(mins), parseInt(secs));
    const displayDate = `${day}/${month}/${year}`;
    const shortAction = action.replace(' as spouse', '');
    const displayText = `${shortAction} ${spouseName}`;
    return { timestamp, displayDate, displayText };
  }

  function renderHistoryItem(entry, container) {
    const li = document.createElement('li');
    li.className = 'spouse-history-item';
    li.innerText = `${entry.displayDate}: ${entry.displayText}`;
    const expandLink = container.querySelector('.expand-link');
    if(expandLink) { container.insertBefore(li, expandLink); } else { container.appendChild(li); }
  }

  function openSpouseModal() {
    const f = state.currentContactRecord.fields;
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

  function closeSpouseModal() { 
    closeModal('spouseModal'); 
  }

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
      if (!records || records.length === 0) { 
        resultsDiv.innerHTML += '<div style="padding:8px; font-style:italic; color:#999;">No recent contacts</div>'; 
      } else {
        records.forEach(r => {
          if(r.id === state.currentContactRecord.id) return;
          renderSearchResultItem(r, resultsDiv);
        });
      }
    }).getRecentContacts();
  }

  function handleSpouseSearch(event) {
    const query = event.target.value;
    const resultsDiv = document.getElementById('spouseSearchResults');
    clearTimeout(state.spouseSearchTimeout);
    if(query.length === 0) { loadRecentContactsForModal(); return; }
    resultsDiv.style.display = 'block';
    resultsDiv.innerHTML = '<div style="padding:10px; color:#999; font-style:italic;">Searching...</div>';
    state.spouseSearchTimeout = setTimeout(() => {
      google.script.run.withSuccessHandler(function(records) {
        resultsDiv.innerHTML = '';
        if (records.length === 0) { 
          resultsDiv.innerHTML = '<div style="padding:8px; font-style:italic; color:#999;">No results</div>'; 
        } else {
          records.forEach(r => {
            if(r.id === state.currentContactRecord.id) return;
            renderSearchResultItem(r, resultsDiv);
          });
        }
      }).searchContacts(query);
    }, 500);
  }

  function renderSearchResultItem(r, container) {
    const name = formatName(r.fields);
    const details = formatDetailsRow(r.fields); 
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
    const myId = state.currentContactRecord.id;
    let statusStr = ""; 
    let otherId = ""; 
    let expectHasSpouse = false; 
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
      state.pollAttempts = 0; 
      startPolling(myId, expectHasSpouse);
    }).setSpouseStatus(myId, otherId, statusStr);
  }

  function startPolling(contactId, expectHasSpouse) {
    if(state.pollInterval) clearInterval(state.pollInterval);
    state.pollInterval = setInterval(() => {
      state.pollAttempts++;
      if (state.pollAttempts > 20) { 
        clearInterval(state.pollInterval);
        const statusEl = document.getElementById('spouseStatusText');
        if(statusEl) { 
          statusEl.innerHTML = `<span style="color:#A00;">Update delayed.</span> <a class="data-link" onclick="forceReload('${contactId}')">Refresh</a>`; 
        }
        return;
      }
      google.script.run.withSuccessHandler(function(r) {
        if(r && r.fields) {
          const currentSpouseId = (r.fields['Spouse'] && r.fields['Spouse'].length > 0) ? r.fields['Spouse'][0] : null;
          const match = expectHasSpouse ? (currentSpouseId !== null) : (currentSpouseId === null);
          if (match) {
            clearInterval(state.pollInterval);
            if(state.currentContactRecord && state.currentContactRecord.id === contactId) { 
              selectContact(r); 
            }
          }
        }
      }).getContactById(contactId); 
    }, 2000); 
  }

  function forceReload(id) {
    clearInterval(state.pollInterval);
    google.script.run.withSuccessHandler(function(r) { 
      if(r && r.fields) selectContact(r); 
    }).getContactById(id);
  }

  window.renderSpouseSection = renderSpouseSection;
  window.toggleSpouseHistory = toggleSpouseHistory;
  window.parseSpouseHistoryEntry = parseSpouseHistoryEntry;
  window.renderHistoryItem = renderHistoryItem;
  window.openSpouseModal = openSpouseModal;
  window.closeSpouseModal = closeSpouseModal;
  window.backToSearch = backToSearch;
  window.loadRecentContactsForModal = loadRecentContactsForModal;
  window.handleSpouseSearch = handleSpouseSearch;
  window.renderSearchResultItem = renderSearchResultItem;
  window.executeSpouseChange = executeSpouseChange;
  window.startPolling = startPolling;
  window.forceReload = forceReload;
  
})();
