/**
 * Connections Module
 * Connection management between contacts
 */
(function() {
  'use strict';
  
  const state = window.IntegrityState;
  
  // ============================================================
  // Load Connections
  // ============================================================
  
  window.loadConnections = function(contactId) {
    google.script.run.withSuccessHandler(function(connections) {
      renderConnectionsList(connections || []);
    }).getConnectionsForContact(contactId);
  };
  
  // ============================================================
  // Toggle Connections Accordion
  // ============================================================
  
  window.toggleConnectionsAccordion = function() {
    const list = document.getElementById('connectionsList');
    const arrow = document.getElementById('connectionsAccordionArrow');
    
    if (list) {
      const isHidden = list.style.display === 'none';
      list.style.display = isHidden ? 'block' : 'none';
      arrow?.classList.toggle('expanded', isHidden);
    }
  };
  
  window.toggleConnectionsExpand = function() {
    toggleConnectionsAccordion();
  };
  
  window.toggleConnectionGroup = function(groupId) {
    const content = document.getElementById(groupId);
    const header = content?.previousElementSibling;
    const icon = header?.querySelector('.collapsible-icon');
    
    if (content) {
      const isHidden = content.style.display === 'none';
      content.style.display = isHidden ? 'block' : 'none';
      icon?.classList.toggle('expanded', isHidden);
    }
  };
  
  // ============================================================
  // Render Connections List
  // ============================================================
  
  window.renderConnectionsList = function(connections) {
    const accordionWrapper = document.getElementById('connectionsAccordionWrapper');
    const noAccordionAdd = document.getElementById('connectionsNoAccordionAdd');
    const connectionsList = document.getElementById('connectionsList');
    
    // Filter active connections only
    const activeConnections = connections.filter(c => c.isActive !== false);
    
    if (activeConnections.length === 0) {
      accordionWrapper.style.display = 'none';
      noAccordionAdd.style.display = 'block';
      connectionsList.innerHTML = '';
      return;
    }
    
    accordionWrapper.style.display = 'block';
    noAccordionAdd.style.display = 'none';
    
    // Group by role
    const groups = {};
    activeConnections.forEach(conn => {
      const role = conn.role || 'Other';
      if (!groups[role]) groups[role] = [];
      groups[role].push(conn);
    });
    
    // Sort roles
    const roleOrder = ['Parent', 'Child', 'Sibling', 'Friend', 'Household Rep', 'Household Member', 
                       'Employer', 'Employee', 'Referred By', 'Referred To', 'Other'];
    const sortedRoles = Object.keys(groups).sort((a, b) => {
      const aIdx = roleOrder.indexOf(a);
      const bIdx = roleOrder.indexOf(b);
      return (aIdx === -1 ? 999 : aIdx) - (bIdx === -1 ? 999 : bIdx);
    });
    
    // Build HTML
    let html = '';
    sortedRoles.forEach(role => {
      const conns = groups[role];
      const groupId = `connGroup_${role.replace(/\s+/g, '_')}`;
      const badgeClass = getRoleBadgeClass(role);
      
      html += `
        <div class="connection-group">
          <div class="connection-group-header" onclick="toggleConnectionGroup('${groupId}')">
            <span class="collapsible-icon expanded">&#9654;</span>
            <span class="connection-role-badge ${badgeClass}">${escapeHtml(role)}</span>
            <span class="connection-count">(${conns.length})</span>
          </div>
          <div class="connection-group-content" id="${groupId}">
      `;
      
      conns.forEach(conn => {
        const hasNote = conn.note && conn.note.trim().length > 0;
        html += `
          <div class="connection-item" onclick="openConnectionDetailsModal(${escapeHtmlForAttr(JSON.stringify(conn))})">
            <span class="connection-name">${escapeHtml(conn.otherContactName || 'Unknown')}</span>
            <span class="connection-note-icon ${hasNote ? 'has-note' : ''}" 
                  onclick="event.stopPropagation(); openConnectionNotePopover(this, '${conn.id}', '${escapeHtmlForAttr(conn.note || '')}')"
                  title="${hasNote ? 'View/edit note' : 'Add note'}">
              ${hasNote ? '&#9998;' : '&#9998;'}
            </span>
          </div>
        `;
      });
      
      html += `
          </div>
        </div>
      `;
    });
    
    connectionsList.innerHTML = html;
    connectionsList.style.display = 'block';
  };
  
  // ============================================================
  // Connection Role Badge Classes
  // ============================================================
  
  function getRoleBadgeClass(role) {
    const classes = {
      'Parent': 'role-parent',
      'Child': 'role-child',
      'Sibling': 'role-sibling',
      'Friend': 'role-friend',
      'Household Rep': 'role-household-rep',
      'Household Member': 'role-household-member',
      'Employer': 'role-employer',
      'Employee': 'role-employee',
      'Referred By': 'role-referred-by',
      'Referred To': 'role-referred-to'
    };
    return classes[role] || 'role-other';
  }
  
  // ============================================================
  // Connection Details Modal
  // ============================================================
  
  window.openConnectionDetailsModal = function(conn) {
    const modal = document.getElementById('connectionDetailsModal');
    const nameEl = document.getElementById('connectionDetailName');
    const roleEl = document.getElementById('connectionDetailRole');
    const noteEl = document.getElementById('connectionDetailNote');
    const viewBtn = document.getElementById('connectionViewBtn');
    const deactivateBtn = document.getElementById('connectionDeactivateBtn');
    
    if (!modal) return;
    
    nameEl.textContent = conn.otherContactName || 'Unknown';
    roleEl.innerHTML = `<span class="connection-role-badge ${getRoleBadgeClass(conn.role)}">${escapeHtml(conn.role || 'Other')}</span>`;
    noteEl.textContent = conn.note || 'No note';
    
    viewBtn.onclick = () => {
      closeModal('connectionDetailsModal');
      loadContactById(conn.otherContactId, true);
    };
    
    deactivateBtn.onclick = () => {
      closeModal('connectionDetailsModal');
      state.deactivatingConnectionId = conn.id;
      openModal('deactivateConnectionModal');
    };
    
    openModal('connectionDetailsModal');
  };
  
  window.closeConnectionDetailsModal = function() {
    closeModal('connectionDetailsModal');
  };
  
  // ============================================================
  // Add Connection Modal
  // ============================================================
  
  window.openAddConnectionModal = function() {
    const modal = document.getElementById('addConnectionModal');
    const searchInput = document.getElementById('connectionSearchInput');
    const resultsList = document.getElementById('connectionSearchResults');
    const step1 = document.getElementById('connectionStep1');
    const step2 = document.getElementById('connectionStep2');
    
    if (searchInput) searchInput.value = '';
    if (resultsList) resultsList.innerHTML = '';
    if (step1) step1.style.display = 'block';
    if (step2) step2.style.display = 'none';
    
    state.selectedConnectionTarget = null;
    
    openModal('addConnectionModal');
    loadRecentContactsForConnectionModal();
    
    setTimeout(() => searchInput?.focus(), 100);
  };
  
  window.closeAddConnectionModal = function() {
    closeModal('addConnectionModal');
    state.selectedConnectionTarget = null;
  };
  
  // ============================================================
  // Connection Search
  // ============================================================
  
  function loadRecentContactsForConnectionModal() {
    google.script.run.withSuccessHandler(function(records) {
      renderConnectionSearchResults(records || []);
    }).getRecentContacts();
  }
  
  window.handleConnectionSearch = function(event) {
    const query = event.target.value;
    
    clearTimeout(state.linkedSearchTimeout);
    
    if (query.length === 0) {
      loadRecentContactsForConnectionModal();
      return;
    }
    
    state.linkedSearchTimeout = setTimeout(() => {
      google.script.run.withSuccessHandler(function(records) {
        renderConnectionSearchResults(records || []);
      }).searchContacts(query);
    }, 300);
  };
  
  function renderConnectionSearchResults(contacts) {
    const container = document.getElementById('connectionSearchResults');
    if (!container) return;
    
    container.innerHTML = '';
    
    contacts.forEach(r => {
      const f = r.fields;
      const name = formatName(f);
      const details = formatDetailsRow(f);
      
      // Skip current contact
      if (state.currentContactRecord && r.id === state.currentContactRecord.id) return;
      
      const div = document.createElement('div');
      div.className = 'connection-search-result';
      div.innerHTML = `
        <div class="connection-result-name">${escapeHtml(name)}</div>
        <div class="connection-result-details">${escapeHtml(details)}</div>
      `;
      div.onclick = () => selectConnectionTarget(r.id, name);
      container.appendChild(div);
    });
  }
  
  function selectConnectionTarget(contactId, contactName) {
    state.selectedConnectionTarget = { id: contactId, name: contactName };
    
    document.getElementById('connectionStep1').style.display = 'none';
    document.getElementById('connectionStep2').style.display = 'block';
    document.getElementById('selectedConnectionName').textContent = contactName;
    
    populateConnectionRoleSelect();
  }
  
  function populateConnectionRoleSelect() {
    const select = document.getElementById('connectionRoleSelect');
    if (!select) return;
    
    const roles = ['Parent', 'Child', 'Sibling', 'Friend', 'Household Rep', 'Household Member',
                   'Employer', 'Employee', 'Referred By', 'Referred To', 'Other'];
    
    select.innerHTML = roles.map(r => `<option value="${r}">${r}</option>`).join('');
  }
  
  window.backToConnectionStep1 = function() {
    document.getElementById('connectionStep1').style.display = 'block';
    document.getElementById('connectionStep2').style.display = 'none';
    state.selectedConnectionTarget = null;
  };
  
  // ============================================================
  // Create Connection
  // ============================================================
  
  window.executeCreateConnection = function() {
    const currentId = state.currentContactRecord?.id;
    const targetId = state.selectedConnectionTarget?.id;
    const role = document.getElementById('connectionRoleSelect')?.value;
    
    if (!currentId || !targetId || !role) return;
    
    closeAddConnectionModal();
    
    google.script.run.withSuccessHandler(function(result) {
      if (result.success) {
        loadConnections(currentId);
      } else {
        showAlert('Error', result.error || 'Failed to create connection', 'error');
      }
    }).createConnection(currentId, targetId, role);
  };
  
  // ============================================================
  // Deactivate Connection
  // ============================================================
  
  window.closeDeactivateConnectionModal = function() {
    closeModal('deactivateConnectionModal');
    state.deactivatingConnectionId = null;
  };
  
  window.executeDeactivateConnection = function() {
    const connectionId = state.deactivatingConnectionId;
    const currentId = state.currentContactRecord?.id;
    
    if (!connectionId || !currentId) return;
    
    closeDeactivateConnectionModal();
    
    google.script.run.withSuccessHandler(function(result) {
      if (result.success) {
        loadConnections(currentId);
      } else {
        showAlert('Error', result.error || 'Failed to deactivate connection', 'error');
      }
    }).deactivateConnection(connectionId);
  };
  
})();
