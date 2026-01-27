/**
 * Connections Module
 * Handles relationship/connection management between contacts
 */
(function() {
  'use strict';
  
  const state = window.IntegrityState;
  
  // Verify state is available
  if (!state) {
    console.error('[Connections Module] ERROR: IntegrityState not found! Module will not work.');
    return;
  }
  console.log('[Connections Module] Loaded successfully, state available');

  // ============================================================
  // HELPER FUNCTIONS
  // ============================================================

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

  // ============================================================
  // CONNECTION RENDERING
  // ============================================================

  window.renderConnectionsSection = function(contactId) {
    const container = document.getElementById('connectionsList');
    if (!container) return;
    
    container.innerHTML = '<div style="font-size: 11px; color: #999; font-style: italic;">Loading connections...</div>';
    
    google.script.run.withSuccessHandler(function(connections) {
      state.allConnectionsData = connections || [];
      renderConnectionsList();
    }).getConnectionsForContact(contactId);
  };
  
  function renderConnectionsList() {
    const container = document.getElementById('connectionsList');
    if (!container) return;
    
    const connections = state.allConnectionsData;
    
    if (!connections || connections.length === 0) {
      container.innerHTML = '<div style="font-size: 11px; color: #999; font-style: italic;">No connections</div>';
      return;
    }
    
    const activeConnections = connections.filter(c => c.fields && c.fields.Status === 'Active');
    
    if (activeConnections.length === 0) {
      container.innerHTML = '<div style="font-size: 11px; color: #999; font-style: italic;">No active connections</div>';
      return;
    }
    
    activeConnections.sort((a, b) => {
      const roleOrder = { 'Spouse': 0, 'Partner': 1, 'Child': 2, 'Parent': 3, 'Sibling': 4 };
      const aOrder = roleOrder[a.fields.Role] ?? 99;
      const bOrder = roleOrder[b.fields.Role] ?? 99;
      if (aOrder !== bOrder) return aOrder - bOrder;
      const aName = a.fields['Related Contact Name']?.[0] || '';
      const bName = b.fields['Related Contact Name']?.[0] || '';
      return aName.localeCompare(bName);
    });
    
    const initialDisplay = state.connectionsExpanded ? activeConnections : activeConnections.slice(0, 3);
    const hasMore = activeConnections.length > 3;
    
    container.innerHTML = '';
    
    initialDisplay.forEach(conn => {
      const f = conn.fields;
      const relatedName = f['Related Contact Name']?.[0] || 'Unknown';
      const relatedId = f['Related Contact']?.[0] || null;
      const role = f.Role || 'Connection';
      const notes = f.Notes || '';
      const roleClass = getRoleBadgeClass(role);
      
      const div = document.createElement('div');
      div.className = 'connection-item';
      div.innerHTML = `
        <div class="connection-item-main">
          <span class="connection-role-badge ${roleClass}">${role}</span>
          <span class="connection-name" data-contact-id="${relatedId || ''}">${relatedName}</span>
          ${notes ? `<span class="connection-notes-indicator" title="${escapeHtmlForAttr(notes)}">📝</span>` : ''}
        </div>
        <span class="connection-deactivate" onclick="window.openDeactivateConnectionModal('${conn.id}', '${escapeHtmlForAttr(relatedName)}', '${escapeHtmlForAttr(role)}')" title="Deactivate connection">×</span>
      `;
      
      const nameEl = div.querySelector('.connection-name');
      if (nameEl && relatedId && typeof attachQuickViewToElement === 'function') {
        attachQuickViewToElement(nameEl, relatedId);
      }
      
      container.appendChild(div);
    });
    
    if (hasMore) {
      const toggleDiv = document.createElement('div');
      toggleDiv.className = 'connections-toggle';
      if (state.connectionsExpanded) {
        toggleDiv.innerHTML = '<span class="connections-toggle-link" onclick="window.collapseConnections()">Show less</span>';
      } else {
        const remaining = activeConnections.length - 3;
        toggleDiv.innerHTML = `<span class="connections-toggle-link" onclick="window.expandConnections()">+${remaining} more</span>`;
      }
      container.appendChild(toggleDiv);
    }
  }

  window.expandConnections = function() {
    state.connectionsExpanded = true;
    renderConnectionsList();
  };

  window.collapseConnections = function() {
    state.connectionsExpanded = false;
    renderConnectionsList();
  };

  // ============================================================
  // DEACTIVATE CONNECTION MODAL
  // ============================================================

  window.openDeactivateConnectionModal = function(connectionId, contactName, role) {
    state.deactivatingConnectionId = connectionId;
    document.getElementById('deactivateConnectionName').innerText = contactName;
    document.getElementById('deactivateConnectionRole').innerText = role;
    
    const modal = document.getElementById('deactivateConnectionModal');
    modal.classList.add('visible');
    setTimeout(() => modal.classList.add('showing'), 10);
  };

  window.closeDeactivateConnectionModal = function() {
    const modal = document.getElementById('deactivateConnectionModal');
    modal.classList.remove('showing');
    setTimeout(() => modal.classList.remove('visible'), 250);
    state.deactivatingConnectionId = null;
  };

  window.executeDeactivateConnection = function() {
    const connectionId = state.deactivatingConnectionId;
    if (!connectionId) return;
    
    const btn = document.querySelector('#deactivateConnectionModal .btn-danger');
    if (btn) {
      btn.disabled = true;
      btn.innerText = 'Deactivating...';
    }
    
    google.script.run.withSuccessHandler(function(result) {
      window.closeDeactivateConnectionModal();
      if (btn) {
        btn.disabled = false;
        btn.innerText = 'Deactivate';
      }
      if (state.currentContactRecord) {
        window.renderConnectionsSection(state.currentContactRecord.id);
      }
    }).withFailureHandler(function(err) {
      console.error('Deactivate connection error:', err);
      if (btn) {
        btn.disabled = false;
        btn.innerText = 'Deactivate';
      }
      alert('Error deactivating connection: ' + (err.message || err));
    }).deactivateConnection(connectionId);
  };

  // ============================================================
  // ADD CONNECTION MODAL
  // ============================================================

  window.openAddConnectionModal = function() {
    console.log('[Connections] openAddConnectionModal called');
    if (!state.currentContactRecord) {
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
        window.populateConnectionRoleSelect();
      }).getConnectionRoleTypes();
    }
    
    modal.classList.add('visible');
    setTimeout(() => modal.classList.add('showing'), 10);
    
    window.loadRecentContactsForConnectionModal();
  };
  
  window.closeAddConnectionModal = function() {
    const modal = document.getElementById('addConnectionModal');
    modal.classList.remove('showing');
    setTimeout(() => modal.classList.remove('visible'), 250);
  };

  // ============================================================
  // CONNECTION SEARCH FUNCTIONS - CRITICAL
  // ============================================================
  
  window.loadRecentContactsForConnectionModal = function() {
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
      window.renderConnectionSearchResults(contacts);
    }).withFailureHandler(function(err) {
      console.error('[Connections] getRecentContacts error:', err);
      results.innerHTML = '<div class="search-option" style="color:#A00;">Error loading contacts</div>';
    }).getRecentContacts();
  };
  
  window.handleConnectionSearch = function(event) {
    const query = event.target.value.trim();
    console.log('[Connections] handleConnectionSearch called with query:', query);
    
    const results = document.getElementById('connectionSearchResults');
    if (!results) {
      console.error('[Connections] connectionSearchResults element not found!');
      return;
    }
    
    if (query.length < 2) {
      console.log('[Connections] Query too short, loading recent contacts');
      window.loadRecentContactsForConnectionModal();
      return;
    }
    
    results.innerHTML = '<div class="search-option" style="color:#999; font-style:italic;">Searching...</div>';
    results.style.display = 'block';
    
    console.log('[Connections] Calling API: searchContacts with query:', query);
    google.script.run.withSuccessHandler(function(contacts) {
      console.log('[Connections] searchContacts returned:', contacts ? contacts.length : 0, 'contacts');
      window.renderConnectionSearchResults(contacts);
    }).withFailureHandler(function(err) {
      console.error('[Connections] searchContacts error:', err);
      results.innerHTML = '<div class="search-option" style="color:#A00;">Search error</div>';
    }).searchContacts(query);
  };
  
  window.renderConnectionSearchResults = function(contacts) {
    const results = document.getElementById('connectionSearchResults');
    results.innerHTML = '';
    results.style.display = 'block';
    
    if (!contacts || contacts.length === 0) {
      results.innerHTML = '<div class="search-option" style="color:#999;">No contacts found</div>';
      return;
    }
    
    const currentId = state.currentContactRecord?.id;
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
      div.onclick = function() { window.selectConnectionTarget(contact.id, name); };
      results.appendChild(div);
    });
  };
  
  window.selectConnectionTarget = function(contactId, contactName) {
    document.getElementById('targetConnectionContactId').value = contactId;
    document.getElementById('targetContactNameConn').innerText = contactName;
    
    const f = state.currentContactRecord.fields;
    const currentName = `${f.FirstName || ''} ${f.LastName || ''}`.trim();
    document.getElementById('currentContactNameConn').innerText = currentName;
    
    window.populateConnectionRoleSelect();
    
    document.getElementById('connectionStep1').style.display = 'none';
    document.getElementById('connectionStep2').style.display = 'flex';
  };
  
  window.populateConnectionRoleSelect = function() {
    const select = document.getElementById('connectionRoleSelect');
    select.innerHTML = '<option value="">-- Select relationship --</option>';
    
    state.connectionRoleTypes.forEach((pair, index) => {
      const option = document.createElement('option');
      option.value = index;
      option.textContent = `${pair.role1} of this contact`;
      select.appendChild(option);
    });
  };
  
  window.backToConnectionStep1 = function() {
    document.getElementById('connectionStep1').style.display = 'flex';
    document.getElementById('connectionStep2').style.display = 'none';
  };
  
  window.executeCreateConnection = function() {
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
      window.closeAddConnectionModal();
      if (btn) {
        btn.disabled = false;
        btn.innerText = 'Create Connection';
      }
      window.renderConnectionsSection(currentContactId);
    }).withFailureHandler(function(err) {
      console.error('Create connection error:', err);
      if (btn) {
        btn.disabled = false;
        btn.innerText = 'Create Connection';
      }
      alert('Error creating connection: ' + (err.message || err));
    }).createConnection(currentContactId, targetContactId, rolePair.role1, rolePair.role2);
  };

  // Expose helper functions
  window.escapeHtmlForAttr = escapeHtmlForAttr;

  console.log('[Connections Module] All functions exposed to window');
})();
