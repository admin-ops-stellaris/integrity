/**
 * Connections Module
 * Connection management between contacts with role-based relationships
 */
(function() {
  'use strict';
  
  const state = window.IntegrityState;
  const CONNECTIONS_COLLAPSED_LIMIT = 8;
  
  function toggleConnectionsAccordion() {
    const content = document.getElementById('connectionsContent');
    const arrowEl = document.getElementById('connectionsAccordionArrow');
    if (content && arrowEl) {
      const isExpanded = arrowEl.classList.contains('expanded');
      content.style.display = isExpanded ? 'none' : 'flex';
      arrowEl.classList.toggle('expanded');
    }
  }
  
  function loadConnections(contactId) {
    const leftList = document.getElementById('connectionsListLeft');
    const rightList = document.getElementById('connectionsListRight');
    if (!leftList || !rightList) return;
    leftList.innerHTML = '<li class="connections-empty">Loading...</li>';
    rightList.innerHTML = '';
    
    google.script.run.withSuccessHandler(function(connections) {
      state.allConnectionsData = connections || [];
      state.connectionsExpanded = false;
      renderConnectionsList(state.allConnectionsData);
    }).withFailureHandler(function(err) {
      leftList.innerHTML = '<li class="connections-empty">Error loading connections</li>';
    }).getConnectionsForContact(contactId);
  }
  
  function toggleConnectionsExpand() {
    state.connectionsExpanded = !state.connectionsExpanded;
    renderConnectionsList(state.allConnectionsData);
  }
  
  function renderConnectionsPills(connections) {
    const container = document.getElementById('connectionsPillContainer');
    if (!container) return;
    container.innerHTML = '';
    
    if (!connections || connections.length === 0) {
      container.innerHTML = '<span class="connections-empty">No connections yet</span>';
      return;
    }
    
    const pillRoleLabels = {
      'parent': 'Parent',
      'child': 'Child',
      'sibling': 'Sibling',
      'friend': 'Friend',
      'employer of': 'Employer',
      'employee of': 'Employee',
      'referred by': 'Referrer',
      'has referred': 'Referred',
      'family': 'Family'
    };
    
    const getPillRoleLabel = (role) => {
      const r = (role || '').toLowerCase().trim();
      return pillRoleLabels[r] || role;
    };
    
    const getPillRoleClass = (role) => {
      const r = (role || '').toLowerCase();
      if (r.includes('parent') || r.includes('child')) return 'parent';
      if (r.includes('sibling')) return 'sibling';
      if (r.includes('friend')) return 'friend';
      if (r.includes('employer') || r.includes('employee')) return 'employer';
      if (r.includes('referred') || r.includes('referral')) return 'referred';
      if (r.includes('household')) return 'household';
      if (r.includes('family')) return 'family';
      return '';
    };
    
    const roleOrder = ['referred by', 'parent', 'child', 'sibling', 'employer of', 'employee of', 'family', 'friend', 'has referred'];
    connections.sort((a, b) => {
      const aRole = (a.myRole || '').toLowerCase();
      const bRole = (b.myRole || '').toLowerCase();
      const aIdx = roleOrder.findIndex(r => aRole.includes(r));
      const bIdx = roleOrder.findIndex(r => bRole.includes(r));
      if (aIdx !== bIdx) return (aIdx === -1 ? 999 : aIdx) - (bIdx === -1 ? 999 : bIdx);
      return (a.otherContactName || '').localeCompare(b.otherContactName || '');
    });
    
    connections.forEach(conn => {
      const pill = document.createElement('div');
      pill.className = 'connection-pill' + (conn.note && conn.note.trim() ? ' has-note' : '');
      pill.setAttribute('data-conn-id', conn.id);
      pill.setAttribute('data-contact-id', conn.otherContactId || '');
      
      const roleClass = getPillRoleClass(conn.myRole);
      const roleLabel = getPillRoleLabel(conn.myRole);
      
      const deceasedSuffix = conn.otherContactDeceased ? ' (DECEASED)' : '';
      pill.innerHTML = `
        <span class="pill-role ${roleClass}">${roleLabel}</span>
        <span class="pill-name">${conn.otherContactName || 'Unknown'}${deceasedSuffix}</span>
      `;
      if (conn.otherContactDeceased) pill.style.opacity = '0.6';
      
      pill.addEventListener('click', function(e) {
        openConnectionDetailsModal(conn);
      });
      
      if (conn.otherContactId && typeof attachQuickViewToElement === 'function') {
        attachQuickViewToElement(pill, conn.otherContactId);
      }
      
      container.appendChild(pill);
    });
  }
  
  function renderConnectionsList(connections) {
    const leftList = document.getElementById('connectionsListLeft');
    const rightList = document.getElementById('connectionsListRight');
    const accordionWrapper = document.getElementById('connectionsAccordionWrapper');
    const noAccordionAdd = document.getElementById('connectionsNoAccordionAdd');
    const connectionsContent = document.getElementById('connectionsContent');
    const accordionArrow = document.getElementById('connectionsAccordionArrow');
    if (!leftList || !rightList) return;
    leftList.innerHTML = '';
    rightList.innerHTML = '';
    
    let friendCount = 0;
    let refersCount = 0;
    if (connections && connections.length > 0) {
      connections.forEach(conn => {
        const role = (conn.myRole || '').toLowerCase().trim();
        if (role === 'friend') friendCount++;
        if (role === 'has referred') refersCount++;
      });
    }
    const needsAccordion = friendCount >= 6 || refersCount >= 6;
    
    if (accordionWrapper) accordionWrapper.style.display = needsAccordion ? 'block' : 'none';
    if (noAccordionAdd) noAccordionAdd.style.display = needsAccordion ? 'none' : 'flex';
    
    if (needsAccordion && connectionsContent && accordionArrow) {
      connectionsContent.style.display = 'flex';
      accordionArrow.classList.add('expanded');
    }
    
    if (!connections || connections.length === 0) {
      leftList.innerHTML = '<li class="connections-empty">No connections yet</li>';
      return;
    }
    
    const roleDisplayMap = {
      'parent': 'Parent of',
      'child': 'Child of',
      'sibling': 'Sibling of',
      'friend': 'Friend of',
      'employer of': 'Employer of',
      'employee of': 'Employee of',
      'referred by': 'Referred by',
      'has referred': 'Has Referred'
    };
    
    const getDisplayRole = (role) => {
      const r = (role || '').toLowerCase().trim();
      return roleDisplayMap[r] || role;
    };
    
    const formatConnectionDate = (dateStr) => {
      if (!dateStr) return '';
      try {
        const d = new Date(dateStr);
        if (isNaN(d.getTime())) return '';
        const day = String(d.getDate()).padStart(2, '0');
        const month = String(d.getMonth() + 1).padStart(2, '0');
        const year = d.getFullYear();
        return `${day}/${month}/${year}`;
      } catch (e) {
        return '';
      }
    };
    
    const rolesWithDate = ['referred by', 'has referred'];
    const leftRoles = ['Referred by', 'Parent', 'Child', 'Sibling', 'Employer of', 'Employee of'];
    const rightRoles = ['Friend', 'Has Referred'];
    const groupableRoles = ['friend', 'has referred'];
    
    const leftConns = [];
    const rightConns = [];
    
    connections.forEach(conn => {
      const role = (conn.myRole || '').trim();
      const isLeft = leftRoles.some(r => role.toLowerCase().includes(r.toLowerCase()));
      const isRight = rightRoles.some(r => role.toLowerCase().includes(r.toLowerCase()));
      
      if (isLeft) {
        leftConns.push(conn);
      } else if (isRight) {
        rightConns.push(conn);
      } else {
        leftConns.push(conn);
      }
    });
    
    const sortByRolePriority = (conns, roles, sortHasReferredByDate = false) => {
      return conns.sort((a, b) => {
        const aRole = (a.myRole || '').toLowerCase();
        const bRole = (b.myRole || '').toLowerCase();
        const aIdx = roles.findIndex(r => aRole.includes(r.toLowerCase()));
        const bIdx = roles.findIndex(r => bRole.includes(r.toLowerCase()));
        if (aIdx !== bIdx) return (aIdx === -1 ? 999 : aIdx) - (bIdx === -1 ? 999 : bIdx);
        
        if (sortHasReferredByDate && aRole === 'has referred' && bRole === 'has referred') {
          const dateA = a.createdOn ? new Date(a.createdOn) : new Date(0);
          const dateB = b.createdOn ? new Date(b.createdOn) : new Date(0);
          return dateB - dateA;
        }
        
        return (a.otherContactName || '').localeCompare(b.otherContactName || '');
      });
    };
    
    sortByRolePriority(leftConns, leftRoles);
    sortByRolePriority(rightConns, rightRoles, true);
    
    const renderSingleConnection = (list, conn) => {
      const li = document.createElement('li');
      li.className = 'connection-item connection-clickable';
      li.setAttribute('data-conn-id', conn.id);
      li.setAttribute('data-conn-name', conn.otherContactName || '');
      li.setAttribute('data-conn-created', conn.createdOn || '');
      li.setAttribute('data-conn-modified', conn.modifiedOn || '');
      li.setAttribute('data-conn-note', conn.note || '');
      
      const badgeClass = getRoleBadgeClass(conn.myRole);
      const displayRole = getDisplayRole(conn.myRole);
      const role = (conn.myRole || '').toLowerCase().trim();
      const showDate = rolesWithDate.includes(role);
      const dateDisplay = showDate ? formatConnectionDate(conn.createdOn) : '';
      const hasNote = conn.note && conn.note.trim();
      const deceasedSuffix = conn.otherContactDeceased ? ' (DECEASED)' : '';
      
      li.innerHTML = `
        <div class="connection-info">
          <span class="connection-role-badge ${badgeClass}">${displayRole}</span>
          <span class="connection-name" data-contact-id="${conn.otherContactId || ''}">${conn.otherContactName}${deceasedSuffix}</span>
          ${dateDisplay ? `<span class="connection-date">${dateDisplay}</span>` : ''}
        </div>
        <button type="button" class="conn-note-icon ${hasNote ? 'has-note' : ''}" data-conn-id="${conn.id}" title="Add/view note"></button>
      `;
      if (conn.otherContactDeceased) li.style.opacity = '0.6';
      
      const noteIcon = li.querySelector('.conn-note-icon');
      if (noteIcon) {
        noteIcon.addEventListener('click', function(e) {
          e.stopPropagation();
          const currentNote = li.getAttribute('data-conn-note') || '';
          openConnectionNotePopover(this, conn.id, currentNote);
        });
      }
      
      li.addEventListener('click', function(e) {
        if (e.target.classList.contains('connection-name') || e.target.classList.contains('conn-note-icon')) {
          e.stopPropagation();
        } else {
          openConnectionDetailsModal(conn);
        }
      });
      
      const nameEl = li.querySelector('.connection-name');
      if (nameEl && conn.otherContactId && typeof attachQuickViewToElement === 'function') {
        attachQuickViewToElement(nameEl, conn.otherContactId);
      }
      
      list.appendChild(li);
    };
    
    leftConns.forEach(conn => renderSingleConnection(leftList, conn));
    rightConns.forEach(conn => renderSingleConnection(rightList, conn));
  }
  
  function openConnectionDetailsModal(conn) {
    const modal = document.getElementById('deactivateConnectionModal');
    const title = modal.querySelector('.modal-title');
    const body = modal.querySelector('.modal-body-content');
    
    let currentContactName = 'Contact';
    if (state.currentContactRecord && state.currentContactRecord.fields) {
      const f = state.currentContactRecord.fields;
      currentContactName = f['Calculated Name'] || 
        `${f.FirstName || ''} ${f.MiddleName || ''} ${f.LastName || ''}`.replace(/\s+/g, ' ').trim();
    }
    
    const roleDisplayMap = {
      'parent': 'Parent of',
      'child': 'Child of',
      'sibling': 'Sibling of',
      'friend': 'Friend of',
      'employer of': 'Employer of',
      'employee of': 'Employee of',
      'referred by': 'Referred by',
      'has referred': 'Has Referred'
    };
    const myRoleLower = (conn.myRole || '').toLowerCase().trim();
    const displayRole = roleDisplayMap[myRoleLower] || conn.myRole;
    
    const formatAuditDateTime = (dateStr) => {
      if (!dateStr) return '';
      try {
        const d = new Date(dateStr);
        if (isNaN(d.getTime())) return '';
        const perthOffset = 8 * 60;
        const localOffset = d.getTimezoneOffset();
        const perthTime = new Date(d.getTime() + (perthOffset + localOffset) * 60000);
        const hours = String(perthTime.getHours()).padStart(2, '0');
        const mins = String(perthTime.getMinutes()).padStart(2, '0');
        const day = String(perthTime.getDate()).padStart(2, '0');
        const month = String(perthTime.getMonth() + 1).padStart(2, '0');
        const year = perthTime.getFullYear();
        return `${hours}:${mins} ${day}/${month}/${year}`;
      } catch (e) {
        return '';
      }
    };
    
    const createdDateTime = formatAuditDateTime(conn.createdOn);
    const modifiedDateTime = formatAuditDateTime(conn.modifiedOn);
    const createdBy = conn.createdByName || '';
    const modifiedBy = conn.modifiedByName || '';
    
    const createdText = createdDateTime ? `${createdDateTime}${createdBy ? ' by ' + createdBy : ''}` : 'Unknown';
    const modifiedText = modifiedDateTime ? `${modifiedDateTime}${modifiedBy ? ' by ' + modifiedBy : ''}` : 'Unknown';
    
    title.textContent = `${currentContactName}: ${displayRole} ${conn.otherContactName}`;
    
    body.innerHTML = `
      <div class="panel-audit-section" style="margin-bottom: 15px; text-align: left;">
        <div><span class="audit-label">Created</span> <span class="audit-value">${createdText}</span></div>
        <div><span class="audit-label">Modified</span> <span class="audit-value">${modifiedText}</span></div>
      </div>
      <div class="connection-modal-remove">
        <span class="remove-label">Remove this connection?</span>
        <button type="button" class="btn-danger conn-modal-btn" onclick="executeDeactivateConnection()">Remove</button>
      </div>
    `;
    
    const modalBody = modal.querySelector('.modal-body');
    let closeDiv = modalBody.querySelector('.connection-modal-close');
    if (!closeDiv) {
      closeDiv = document.createElement('div');
      closeDiv.className = 'connection-modal-close';
      closeDiv.innerHTML = '<button type="button" class="btn-secondary conn-modal-btn" onclick="closeDeactivateConnectionModal()">Close</button>';
      modalBody.appendChild(closeDiv);
    }
    
    modal.setAttribute('data-conn-id', conn.id);
    modal.setAttribute('data-conn-name', conn.otherContactName || '');
    
    modal.classList.add('visible');
    setTimeout(() => modal.classList.add('showing'), 10);
  }
  
  function getRoleBadgeClass(role) {
    if (!role) return '';
    const r = role.toLowerCase();
    if (r === 'parent' || r === 'child') return r;
    if (r === 'sibling') return 'sibling';
    if (r === 'friend') return 'friend';
    if (r.includes('employer')) return 'employer';
    if (r.includes('employee')) return 'employee';
    if (r.includes('referred') || r.includes('referral')) return 'referred';
    if (r.includes('household')) return 'household';
    return '';
  }
  
  function escapeHtmlForAttr(str) {
    if (!str) return '';
    return str.replace(/'/g, "&#39;").replace(/"/g, '&quot;');
  }
  
  function openAddConnectionModal() {
    if (!state.currentContactRecord) return;
    
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
    const results = document.getElementById('connectionSearchResults');
    results.innerHTML = '<div class="search-option" style="color:#999; font-style:italic;">Loading recent contacts...</div>';
    results.style.display = 'block';
    
    google.script.run.withSuccessHandler(function(contacts) {
      renderConnectionSearchResults(contacts);
    }).getRecentContacts();
  }
  
  function handleConnectionSearch(event) {
    const query = event.target.value.trim();
    if (query.length < 2) {
      loadRecentContactsForConnectionModal();
      return;
    }
    
    google.script.run.withSuccessHandler(function(contacts) {
      renderConnectionSearchResults(contacts);
    }).searchContacts(query);
  }
  
  function renderConnectionSearchResults(contacts) {
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
      div.onclick = function() { selectConnectionTarget(contact.id, name); };
      results.appendChild(div);
    });
  }
  
  function selectConnectionTarget(contactId, contactName) {
    document.getElementById('targetConnectionContactId').value = contactId;
    document.getElementById('targetContactNameConn').innerText = contactName;
    
    const f = state.currentContactRecord.fields;
    const currentName = `${f.FirstName || ''} ${f.LastName || ''}`.trim();
    document.getElementById('currentContactNameConn').innerText = currentName;
    
    populateConnectionRoleSelect();
    
    document.getElementById('connectionStep1').style.display = 'none';
    document.getElementById('connectionStep2').style.display = 'flex';
  }
  
  function populateConnectionRoleSelect() {
    const select = document.getElementById('connectionRoleSelect');
    select.innerHTML = '<option value="">-- Select relationship --</option>';
    
    state.connectionRoleTypes.forEach((pair, index) => {
      const option = document.createElement('option');
      option.value = index;
      option.textContent = `${pair.role1} of this contact`;
      select.appendChild(option);
    });
  }
  
  function backToConnectionStep1() {
    document.getElementById('connectionStep1').style.display = 'flex';
    document.getElementById('connectionStep2').style.display = 'none';
  }
  
  function executeCreateConnection() {
    const selectEl = document.getElementById('connectionRoleSelect');
    const roleIndex = selectEl.value;
    if (roleIndex === '') {
      alert('Please select a relationship type');
      return;
    }
    
    const pair = state.connectionRoleTypes[parseInt(roleIndex)];
    const contact1Id = state.currentContactRecord.id;
    const contact2Id = document.getElementById('targetConnectionContactId').value;
    
    const btn = document.getElementById('confirmConnectionBtn');
    btn.disabled = true;
    btn.innerText = 'Creating...';
    
    google.script.run.withSuccessHandler(function(result) {
      btn.disabled = false;
      btn.innerText = 'Create Connection';
      
      if (result.success) {
        closeAddConnectionModal();
        loadConnections(state.currentContactRecord.id);
      } else {
        alert('Error: ' + (result.error || 'Failed to create connection'));
      }
    }).withFailureHandler(function(err) {
      btn.disabled = false;
      btn.innerText = 'Create Connection';
      alert('Error: ' + err.message);
    }).createConnection(contact1Id, contact2Id, pair.role1, pair.role2);
  }
  
  function closeDeactivateConnectionModal() {
    const modal = document.getElementById('deactivateConnectionModal');
    modal.classList.remove('showing');
    setTimeout(() => modal.classList.remove('visible'), 250);
  }
  
  function executeDeactivateConnection() {
    const modal = document.getElementById('deactivateConnectionModal');
    const connectionId = modal.getAttribute('data-conn-id');
    
    if (!connectionId) {
      closeDeactivateConnectionModal();
      return;
    }
    
    google.script.run.withSuccessHandler(function(result) {
      closeDeactivateConnectionModal();
      if (result.success) {
        loadConnections(state.currentContactRecord.id);
      } else {
        alert('Error: ' + (result.error || 'Failed to remove connection'));
      }
    }).withFailureHandler(function(err) {
      closeDeactivateConnectionModal();
      alert('Error: ' + err.message);
    }).deactivateConnection(connectionId);
  }

  window.toggleConnectionsAccordion = toggleConnectionsAccordion;
  window.loadConnections = loadConnections;
  window.toggleConnectionsExpand = toggleConnectionsExpand;
  window.renderConnectionsPills = renderConnectionsPills;
  window.renderConnectionsList = renderConnectionsList;
  window.openConnectionDetailsModal = openConnectionDetailsModal;
  window.toggleConnectionGroup = function(groupId) {
    const list = document.getElementById(groupId);
    const toggle = document.getElementById(groupId + '-toggle');
    if (list && toggle) {
      const isHidden = list.style.display === 'none';
      list.style.display = isHidden ? 'block' : 'none';
      toggle.textContent = isHidden ? '▲' : '▼';
    }
  };
  window.getRoleBadgeClass = getRoleBadgeClass;
  window.escapeHtmlForAttr = escapeHtmlForAttr;
  window.openAddConnectionModal = openAddConnectionModal;
  window.closeAddConnectionModal = closeAddConnectionModal;
  window.loadRecentContactsForConnectionModal = loadRecentContactsForConnectionModal;
  window.handleConnectionSearch = handleConnectionSearch;
  window.renderConnectionSearchResults = renderConnectionSearchResults;
  window.selectConnectionTarget = selectConnectionTarget;
  window.populateConnectionRoleSelect = populateConnectionRoleSelect;
  window.backToConnectionStep1 = backToConnectionStep1;
  window.executeCreateConnection = executeCreateConnection;
  window.closeDeactivateConnectionModal = closeDeactivateConnectionModal;
  window.executeDeactivateConnection = executeDeactivateConnection;
  
})();
