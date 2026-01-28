/**
 * Connections Module
 * Handles relationship/connection management between contacts
 * Functions exposed synchronously to window object
 */
'use strict';

function escapeHtmlForAttr(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function getRoleBadgeClass(role) {
  const roleClasses = {
    'Spouse': 'spouse',
    'Partner': 'spouse',
    'Child': 'family',
    'Parent': 'family',
    'Sibling': 'family',
    'Grandchild': 'family',
    'Grandparent': 'family',
    'Accountant': 'professional',
    'Referred By': 'professional',
    'Financial Planner': 'professional',
    'Lawyer / Solicitor': 'professional',
    'Real Estate Agent': 'professional',
    'Friend': 'personal',
    'Colleague': 'personal',
    'Other': 'other'
  };
  return roleClasses[role] || 'other';
}

function renderConnectionsList() {
  const state = window.IntegrityState;
  if (!state) return;
  
  const leftCol = document.getElementById('connectionsListLeft');
  const rightCol = document.getElementById('connectionsListRight');
  const accordionWrapper = document.getElementById('connectionsAccordionWrapper');
  const noAccordionAdd = document.getElementById('connectionsNoAccordionAdd');
  
  if (!leftCol || !rightCol) {
    console.error('[Connections] Connection list columns not found');
    return;
  }
  
  const connections = state.allConnectionsData;
  
  leftCol.innerHTML = '';
  rightCol.innerHTML = '';
  
  if (!connections || connections.length === 0) {
    leftCol.innerHTML = '<li style="font-size: 11px; color: #999; font-style: italic;">No connections</li>';
    if (accordionWrapper) accordionWrapper.style.display = 'none';
    if (noAccordionAdd) noAccordionAdd.style.display = 'flex';
    return;
  }
  
  // Server already filters to active, but check status field just in case
  // Data is flattened: { id, status, myRole, otherContactName, ... }
  const activeConnections = connections.filter(c => c.status === 'Active' || !c.status);
  
  if (activeConnections.length === 0) {
    leftCol.innerHTML = '<li style="font-size: 11px; color: #999; font-style: italic;">No active connections</li>';
    if (accordionWrapper) accordionWrapper.style.display = 'none';
    if (noAccordionAdd) noAccordionAdd.style.display = 'flex';
    return;
  }
  
  activeConnections.sort((a, b) => {
    const roleOrder = { 'Spouse': 0, 'Partner': 1, 'Child': 2, 'Parent': 3, 'Sibling': 4 };
    const aOrder = roleOrder[a.myRole] ?? 99;
    const bOrder = roleOrder[b.myRole] ?? 99;
    if (aOrder !== bOrder) return aOrder - bOrder;
    const aName = a.otherContactName || '';
    const bName = b.otherContactName || '';
    return aName.localeCompare(bName);
  });
  
  const initialDisplay = state.connectionsExpanded ? activeConnections : activeConnections.slice(0, 6);
  const hasMore = activeConnections.length > 6;
  
  initialDisplay.forEach((conn, index) => {
    const relatedName = conn.otherContactName || 'Unknown';
    const relatedId = conn.otherContactId || null;
    const role = conn.myRole || 'Connection';
    const notes = conn.note || '';
    const roleClass = getRoleBadgeClass(role);
    
    const li = document.createElement('li');
    li.className = 'connection-item';
    li.innerHTML = `
      <div class="connection-item-main">
        <span class="connection-role-badge ${roleClass}">${role}</span>
        <span class="connection-name" data-contact-id="${relatedId || ''}">${relatedName}</span>
        ${notes ? `<span class="connection-notes-indicator" title="${escapeHtmlForAttr(notes)}">📝</span>` : ''}
      </div>
      <span class="connection-deactivate" onclick="window.openDeactivateConnectionModal('${conn.id}', '${escapeHtmlForAttr(relatedName)}', '${escapeHtmlForAttr(role)}')" title="Deactivate connection">×</span>
    `;
    
    const nameEl = li.querySelector('.connection-name');
    if (nameEl && relatedId && typeof attachQuickViewToElement === 'function') {
      attachQuickViewToElement(nameEl, relatedId);
    }
    
    if (index % 2 === 0) {
      leftCol.appendChild(li);
    } else {
      rightCol.appendChild(li);
    }
  });
  
  if (hasMore) {
    const toggleLi = document.createElement('li');
    toggleLi.className = 'connections-toggle';
    toggleLi.style.paddingTop = '6px';
    if (state.connectionsExpanded) {
      toggleLi.innerHTML = '<span class="expand-link" onclick="window.collapseConnections()">Show less</span>';
    } else {
      const remaining = activeConnections.length - 6;
      toggleLi.innerHTML = `<span class="expand-link" onclick="window.expandConnections()">+${remaining} more</span>`;
    }
    leftCol.appendChild(toggleLi);
  }
  
  if (accordionWrapper) accordionWrapper.style.display = activeConnections.length > 3 ? 'block' : 'none';
  if (noAccordionAdd) noAccordionAdd.style.display = activeConnections.length <= 3 ? 'flex' : 'none';
}

function loadConnections(contactId) {
  console.log('[Connections] loadConnections called with contactId:', contactId);
  const state = window.IntegrityState;
  const leftCol = document.getElementById('connectionsListLeft');
  const rightCol = document.getElementById('connectionsListRight');
  
  if (!leftCol || !rightCol) {
    console.error('[Connections] Connection list columns not found!');
    return;
  }
  
  leftCol.innerHTML = '<li style="font-size: 11px; color: #999; font-style: italic;">Loading connections...</li>';
  rightCol.innerHTML = '';
  
  google.script.run.withSuccessHandler(function(connections) {
    console.log('[Connections] getConnectionsForContact returned:', connections ? connections.length : 0, 'connections');
    if (state) state.allConnectionsData = connections || [];
    renderConnectionsList();
  }).withFailureHandler(function(err) {
    console.error('[Connections] getConnectionsForContact error:', err);
    leftCol.innerHTML = '<li style="font-size: 11px; color: #A00;">Error loading connections</li>';
  }).getConnectionsForContact(contactId);
}

function expandConnections() {
  const state = window.IntegrityState;
  if (state) state.connectionsExpanded = true;
  renderConnectionsList();
}

function collapseConnections() {
  const state = window.IntegrityState;
  if (state) state.connectionsExpanded = false;
  renderConnectionsList();
}

function openDeactivateConnectionModal(connectionId, contactName, role) {
  const state = window.IntegrityState;
  if (state) state.deactivatingConnectionId = connectionId;
  document.getElementById('deactivateConnectionName').innerText = contactName;
  document.getElementById('deactivateConnectionRole').innerText = role;
  
  const modal = document.getElementById('deactivateConnectionModal');
  modal.classList.add('visible');
  setTimeout(() => modal.classList.add('showing'), 10);
}

function closeDeactivateConnectionModal() {
  const state = window.IntegrityState;
  const modal = document.getElementById('deactivateConnectionModal');
  modal.classList.remove('showing');
  setTimeout(() => modal.classList.remove('visible'), 250);
  if (state) state.deactivatingConnectionId = null;
}

function executeDeactivateConnection() {
  const state = window.IntegrityState;
  if (!state) return;
  const connectionId = state.deactivatingConnectionId;
  if (!connectionId) return;
  
  const btn = document.querySelector('#deactivateConnectionModal .btn-danger');
  if (btn) {
    btn.disabled = true;
    btn.innerText = 'Deactivating...';
  }
  
  google.script.run.withSuccessHandler(function(result) {
    closeDeactivateConnectionModal();
    if (btn) {
      btn.disabled = false;
      btn.innerText = 'Deactivate';
    }
    if (state.currentContactRecord) {
      loadConnections(state.currentContactRecord.id);
    }
  }).withFailureHandler(function(err) {
    console.error('Deactivate connection error:', err);
    if (btn) {
      btn.disabled = false;
      btn.innerText = 'Deactivate';
    }
    alert('Error deactivating connection: ' + (err.message || err));
  }).deactivateConnection(connectionId);
}

function openAddConnectionModal() {
  console.log('[Connections] openAddConnectionModal called');
  const state = window.IntegrityState;
  if (!state || !state.currentContactRecord) {
    console.error('[Connections] No current contact record!');
    return;
  }
  
  const modal = document.getElementById('addConnectionModal');
  document.getElementById('connectionStep1').style.display = 'flex';
  document.getElementById('connectionStep2').style.display = 'none';
  document.getElementById('connectionSearchInput').value = '';
  document.getElementById('connectionSearchResults').innerHTML = '';
  document.getElementById('connectionSearchResults').style.display = 'none';
  
  if (state.connectionRoleTypes.length === 0) {
    google.script.run.withSuccessHandler(function(types) {
      state.connectionRoleTypes = types || [];
      populateConnectionRoleSelect();
    }).getConnectionRoleTypes();
  }
  
  modal.classList.add('visible');
  setTimeout(() => modal.classList.add('showing'), 10);
  
  loadRecentContactsForConnectionModal();
}

function closeAddConnectionModal() {
  const modal = document.getElementById('addConnectionModal');
  modal.classList.remove('showing');
  setTimeout(() => modal.classList.remove('visible'), 250);
}

function loadRecentContactsForConnectionModal() {
  console.log('[Connections] loadRecentContactsForConnectionModal called');
  const results = document.getElementById('connectionSearchResults');
  if (!results) {
    console.error('[Connections] connectionSearchResults element not found!');
    return;
  }
  results.innerHTML = '<div class="search-option" style="color:#999; font-style:italic;">Loading recent contacts...</div>';
  results.style.display = 'block';
  
  console.log('[Connections] Calling API: getRecentContacts');
  google.script.run.withSuccessHandler(function(contacts) {
    console.log('[Connections] getRecentContacts returned:', contacts ? contacts.length : 0, 'contacts');
    renderConnectionSearchResults(contacts);
  }).withFailureHandler(function(err) {
    console.error('[Connections] getRecentContacts error:', err);
    results.innerHTML = '<div class="search-option" style="color:#A00;">Error loading contacts</div>';
  }).getRecentContacts();
}

function handleConnectionSearch(event) {
  const query = event.target.value.trim();
  console.log('[Connections] handleConnectionSearch called with query:', query);
  
  const results = document.getElementById('connectionSearchResults');
  if (!results) {
    console.error('[Connections] connectionSearchResults element not found!');
    return;
  }
  
  if (query.length < 2) {
    console.log('[Connections] Query too short, loading recent contacts');
    loadRecentContactsForConnectionModal();
    return;
  }
  
  results.innerHTML = '<div class="search-option" style="color:#999; font-style:italic;">Searching...</div>';
  results.style.display = 'block';
  
  console.log('[Connections] Calling API: searchContacts with query:', query);
  google.script.run.withSuccessHandler(function(contacts) {
    console.log('[Connections] searchContacts returned:', contacts ? contacts.length : 0, 'contacts');
    renderConnectionSearchResults(contacts);
  }).withFailureHandler(function(err) {
    console.error('[Connections] searchContacts error:', err);
    results.innerHTML = '<div class="search-option" style="color:#A00;">Search error</div>';
  }).searchContacts(query);
}

function renderConnectionSearchResults(contacts) {
  const state = window.IntegrityState;
  const results = document.getElementById('connectionSearchResults');
  results.innerHTML = '';
  results.style.display = 'block';
  
  if (!contacts || contacts.length === 0) {
    results.innerHTML = '<div class="search-option" style="color:#999;">No contacts found</div>';
    return;
  }
  
  const currentId = state?.currentContactRecord?.id;
  const filtered = contacts.filter(c => c.id !== currentId);
  
  if (filtered.length === 0) {
    results.innerHTML = '<div class="search-option" style="color:#999;">No other contacts found</div>';
    return;
  }
  
  filtered.slice(0, 15).forEach(contact => {
    const f = contact.fields;
    const name = `${f.FirstName || ''} ${f.MiddleName || ''} ${f.LastName || ''}`.replace(/\s+/g, ' ').trim();
    const div = document.createElement('div');
    div.className = 'search-option';
    div.innerHTML = `<strong>${name}</strong>${f.EmailAddress1 ? `<br><span style="font-size:11px; color:#666;">${f.EmailAddress1}</span>` : ''}`;
    div.onclick = function() { selectConnectionTarget(contact.id, name); };
    results.appendChild(div);
  });
}

function selectConnectionTarget(contactId, contactName) {
  const state = window.IntegrityState;
  document.getElementById('targetConnectionContactId').value = contactId;
  document.getElementById('targetContactNameConn').innerText = contactName;
  
  if (state?.currentContactRecord) {
    const f = state.currentContactRecord.fields;
    const currentName = `${f.FirstName || ''} ${f.LastName || ''}`.trim();
    document.getElementById('currentContactNameConn').innerText = currentName;
  }
  
  populateConnectionRoleSelect();
  
  document.getElementById('connectionStep1').style.display = 'none';
  document.getElementById('connectionStep2').style.display = 'flex';
}

function populateConnectionRoleSelect() {
  const state = window.IntegrityState;
  const select = document.getElementById('connectionRoleSelect');
  select.innerHTML = '<option value="">-- Select relationship --</option>';
  
  if (state?.connectionRoleTypes) {
    state.connectionRoleTypes.forEach((pair, index) => {
      const option = document.createElement('option');
      option.value = index;
      option.textContent = `${pair.role1} of this contact`;
      select.appendChild(option);
    });
  }
}

function backToConnectionStep1() {
  document.getElementById('connectionStep1').style.display = 'flex';
  document.getElementById('connectionStep2').style.display = 'none';
}

function executeCreateConnection() {
  const state = window.IntegrityState;
  if (!state) return;
  
  const selectEl = document.getElementById('connectionRoleSelect');
  const roleIndex = selectEl.value;
  if (roleIndex === '') {
    alert('Please select a relationship type');
    return;
  }
  
  const rolePair = state.connectionRoleTypes[parseInt(roleIndex)];
  const targetContactId = document.getElementById('targetConnectionContactId').value;
  const currentContactId = state.currentContactRecord.id;
  
  const btn = document.getElementById('confirmConnectionBtn');
  if (btn) {
    btn.disabled = true;
    btn.innerText = 'Creating...';
  }
  
  google.script.run.withSuccessHandler(function(result) {
    closeAddConnectionModal();
    if (btn) {
      btn.disabled = false;
      btn.innerText = 'Create Connection';
    }
    loadConnections(currentContactId);
  }).withFailureHandler(function(err) {
    console.error('Create connection error:', err);
    if (btn) {
      btn.disabled = false;
      btn.innerText = 'Create Connection';
    }
    alert('Error creating connection: ' + (err.message || err));
  }).createConnection(currentContactId, targetContactId, rolePair.role1, rolePair.role2);
}

window.loadConnections = loadConnections;
window.renderConnectionsSection = loadConnections;
window.expandConnections = expandConnections;
window.collapseConnections = collapseConnections;
window.openDeactivateConnectionModal = openDeactivateConnectionModal;
window.closeDeactivateConnectionModal = closeDeactivateConnectionModal;
window.executeDeactivateConnection = executeDeactivateConnection;
window.openAddConnectionModal = openAddConnectionModal;
window.closeAddConnectionModal = closeAddConnectionModal;
window.loadRecentContactsForConnectionModal = loadRecentContactsForConnectionModal;
window.handleConnectionSearch = handleConnectionSearch;
window.renderConnectionSearchResults = renderConnectionSearchResults;
window.selectConnectionTarget = selectConnectionTarget;
window.populateConnectionRoleSelect = populateConnectionRoleSelect;
window.backToConnectionStep1 = backToConnectionStep1;
window.executeCreateConnection = executeCreateConnection;
window.escapeHtmlForAttr = escapeHtmlForAttr;

console.log('SUCCESS: loadConnections is now global');
console.log('[Connections Module] All functions exposed to window synchronously');
