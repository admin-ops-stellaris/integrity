/**
 * Contacts Module
 * Contact loading, searching, selection, form handling
 */
(function() {
  'use strict';
  
  const state = window.IntegrityState;
  
  // ============================================================
  // Search Dropdown
  // ============================================================
  
  window.showSearchDropdown = function() {
    document.getElementById('searchDropdown').classList.add('visible');
  };
  
  window.hideSearchDropdown = function() {
    document.getElementById('searchDropdown').classList.remove('visible');
    state.searchHighlightIndex = -1;
    updateSearchHighlight();
  };
  
  // Close dropdown when clicking outside
  document.addEventListener('click', function(e) {
    const wrapper = document.querySelector('.header-search-wrapper');
    if (wrapper && !wrapper.contains(e.target)) {
      hideSearchDropdown();
    }
  });
  
  // ============================================================
  // Load Contacts
  // ============================================================
  
  window.loadContacts = function() {
    const loadingDiv = document.getElementById('loading');
    const list = document.getElementById('contactList');
    list.innerHTML = '';
    loadingDiv.style.display = 'block';
    loadingDiv.innerHTML = 'Loading directory...';
    clearTimeout(state.loadingTimer);
    
    state.loadingTimer = setTimeout(() => {
      loadingDiv.innerHTML = `
        <div style="margin-top:15px; text-align:center;">
          <p style="color:#666; font-size:13px;">Taking a while to connect...</p>
          <button onclick="loadContacts()" style="padding:8px 16px; background:var(--color-cedar); color:white; border:none; border-radius:4px; cursor:pointer; font-size:12px; margin-top:8px;">Try Again</button>
        </div>
      `;
    }, 4000);
    
    const statusFilterToSend = state.contactStatusFilter === 'All' ? null : state.contactStatusFilter;
    google.script.run.withSuccessHandler(function(records) {
      clearTimeout(state.loadingTimer);
      document.getElementById('loading').style.display = 'none';
      state.currentSearchRecords = records;
      state.searchHighlightIndex = -1;
      renderList(records);
    }).getRecentContacts(statusFilterToSend);
  };
  
  // ============================================================
  // Search
  // ============================================================
  
  window.handleSearch = function(event) {
    if (['ArrowDown', 'ArrowUp', 'Enter', 'Escape'].includes(event.key)) return;
    
    const query = event.target.value;
    const statusEl = document.getElementById('searchStatus');
    clearTimeout(state.loadingTimer);
    state.searchHighlightIndex = -1;
    showSearchDropdown();
    
    if (query.length === 0) {
      statusEl.innerText = "";
      loadContacts();
      return;
    }
    
    clearTimeout(state.searchTimeout);
    statusEl.innerText = "Typing...";
    
    state.searchTimeout = setTimeout(() => {
      statusEl.innerText = "Searching...";
      const statusFilterToSend = state.contactStatusFilter === 'All' ? null : state.contactStatusFilter;
      google.script.run.withSuccessHandler(function(records) {
        statusEl.innerText = records.length > 0 ? `Found ${records.length} matches` : "No matches found";
        state.currentSearchRecords = records;
        state.searchHighlightIndex = -1;
        renderList(records);
      }).searchContacts(query, statusFilterToSend);
    }, 500);
  };
  
  window.handleSearchKeydown = function(e) {
    const items = document.querySelectorAll('#contactList li');
    if (!items.length) return;
    
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      state.searchHighlightIndex = Math.min(state.searchHighlightIndex + 1, items.length - 1);
      updateSearchHighlight();
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      state.searchHighlightIndex = Math.max(state.searchHighlightIndex - 1, 0);
      updateSearchHighlight();
    } else if (e.key === 'Enter') {
      e.preventDefault();
      if (state.searchHighlightIndex >= 0 && state.searchHighlightIndex < state.currentSearchRecords.length) {
        selectContact(state.currentSearchRecords[state.searchHighlightIndex]);
      }
    } else if (e.key === 'Escape') {
      e.preventDefault();
      hideSearchDropdown();
      e.target.blur();
    }
  };
  
  function updateSearchHighlight() {
    const items = document.querySelectorAll('#contactList li');
    items.forEach((item, idx) => {
      item.classList.toggle('keyboard-highlight', idx === state.searchHighlightIndex);
    });
    if (state.searchHighlightIndex >= 0 && items[state.searchHighlightIndex]) {
      items[state.searchHighlightIndex].scrollIntoView({ block: 'nearest' });
    }
  }
  
  // ============================================================
  // Render Contact List
  // ============================================================
  
  window.renderList = function(records) {
    const list = document.getElementById('contactList');
    list.innerHTML = '';
    state.currentSearchRecords = records;
    
    if (!records || records.length === 0) {
      list.innerHTML = '<li style="color:#AAA;font-size:12px;">No contacts found</li>';
      return;
    }
    
    records.forEach((record, idx) => {
      const f = record.fields;
      const li = document.createElement('li');
      li.onclick = () => selectContact(record);
      
      const initials = getInitials(f.FirstName, f.LastName);
      const avatarColor = getAvatarColor((f.FirstName || '') + (f.LastName || ''));
      const isDeceased = f.Deceased === true || f.Deceased === 'true';
      
      li.innerHTML = `
        <div class="contact-list-item${isDeceased ? ' contact-deceased' : ''}">
          <div class="contact-avatar" style="background-color:${avatarColor};">${initials}</div>
          <div class="contact-info">
            <div class="contact-name">${formatName(f)}${isDeceased ? ' <span class="deceased-badge-small">DECEASED</span>' : ''}</div>
            <div class="contact-details">${formatDetailsRow(f)}</div>
          </div>
        </div>
      `;
      
      if (idx === state.searchHighlightIndex) {
        li.classList.add('keyboard-highlight');
      }
      
      list.appendChild(li);
    });
  };
  
  // ============================================================
  // Select Contact
  // ============================================================
  
  window.selectContact = function(record) {
    toggleProfileView(true);
    hideSearchDropdown();
    state.currentContactRecord = record;
    
    const f = record.fields;
    const isDeceased = f.Deceased === true || f.Deceased === 'true';
    
    // Populate form fields
    document.getElementById('recordId').value = record.id;
    document.getElementById('firstName').value = f.FirstName || '';
    document.getElementById('middleName').value = f.MiddleName || '';
    document.getElementById('lastName').value = f.LastName || '';
    document.getElementById('preferredName').value = f.PreferredName || '';
    document.getElementById('mailingTitle').value = f.mc_mailingtitle || '';
    document.getElementById('salutation').value = f.mc_salutation || '';
    document.getElementById('email1').value = f.EmailAddress1 || '';
    document.getElementById('email2').value = f.Email2 || '';
    document.getElementById('email3').value = f.Email3 || '';
    document.getElementById('mobile').value = f.Mobile || '';
    document.getElementById('homePhone').value = f.Telephone1 || '';
    document.getElementById('workPhone').value = f.Telephone2 || '';
    document.getElementById('dob').value = f.DOB || '';
    document.getElementById('gender').value = f.Gender || '';
    document.getElementById('notes').value = f.Notes || '';
    document.getElementById('status').value = f.Status || 'Active';
    document.getElementById('unsubscribeFromMarketing').value = f.UnsubscribeFromMarketing || 'false';
    
    // Update UI elements
    document.getElementById('formTitle').innerText = formatName(f);
    document.getElementById('formSubtitle').innerHTML = formatSubtitle(f);
    
    handleGenderChange();
    updateUnsubscribeDisplay(f.UnsubscribeFromMarketing === 'true' || f.UnsubscribeFromMarketing === true);
    applyDeceasedStyling(isDeceased);
    
    // Show/hide buttons
    document.getElementById('submitBtn').innerText = "Save Changes";
    document.getElementById('cancelBtn').style.display = 'none';
    document.getElementById('editBtn').style.visibility = 'visible';
    document.getElementById('refreshBtn').style.display = 'inline';
    
    // Show actions menu
    const actionsMenu = document.getElementById('actionsMenuWrapper');
    if (actionsMenu) actionsMenu.style.display = 'block';
    
    // Disable editing
    disableAllFieldEditing();
    
    // Load related data
    renderSpouseSection(f);
    loadOpportunities(f);
    loadConnections(record.id);
    loadAddressHistory(record.id);
    populateNoteFields(record);
    updateAllNoteIcons();
    renderContactMetaBar(f);
    updateContactBackButton();
    
    // Auto-expand textareas
    setTimeout(() => {
      const textareas = document.querySelectorAll('.field-group textarea');
      textareas.forEach(autoExpandTextarea);
    }, 100);
  };
  
  // ============================================================
  // Contact History Navigation
  // ============================================================
  
  window.loadContactById = function(contactId, addToHistory) {
    if (addToHistory && state.currentContactRecord) {
      state.contactHistory.push(state.currentContactRecord.id);
    }
    
    google.script.run.withSuccessHandler(function(record) {
      if (record && record.fields) {
        selectContact(record);
      }
    }).getContactById(contactId);
  };
  
  function updateContactBackButton() {
    const btn = document.getElementById('contactBackBtn');
    if (btn) {
      btn.style.display = state.contactHistory.length > 0 ? 'block' : 'none';
    }
  }
  
  window.goBackToContact = function() {
    if (state.contactHistory.length > 0) {
      const prevId = state.contactHistory.pop();
      google.script.run.withSuccessHandler(function(record) {
        if (record && record.fields) {
          selectContact(record);
        }
      }).getContactById(prevId);
    }
  };
  
  window.navigateFromQuickView = function() {
    const quickView = document.getElementById('contactQuickView');
    const contactId = quickView?.dataset.contactId;
    if (contactId) {
      hideContactQuickView();
      loadContactById(contactId, true);
    }
  };
  
  // ============================================================
  // Form Handling
  // ============================================================
  
  window.resetForm = function() {
    toggleProfileView(true);
    document.getElementById('contactForm').reset();
    document.getElementById('recordId').value = "";
    enableNewContactMode();
    document.getElementById('formTitle').innerText = "New Contact";
    document.getElementById('formSubtitle').innerText = '';
    document.getElementById('submitBtn').innerText = "Create Contact";
    document.getElementById('cancelBtn').style.display = 'inline-block';
    document.getElementById('editBtn').style.visibility = 'hidden';
    document.getElementById('oppList').innerHTML = '<li style="color:#CCC; font-size:12px; font-style:italic;">No opportunities linked.</li>';
    document.getElementById('contactMetaBar').classList.remove('visible');
    document.getElementById('duplicateWarningBox').style.display = 'none';
    document.getElementById('spouseStatusText').innerHTML = "Single";
    document.getElementById('spouseHistoryList').innerHTML = "";
    document.getElementById('spouseEditLink').style.display = 'inline';
    document.getElementById('refreshBtn').style.display = 'none';
    
    const actionsMenu = document.getElementById('actionsMenuWrapper');
    if (actionsMenu) actionsMenu.style.display = 'none';
    const deceasedBadge = document.getElementById('deceasedBadge');
    if (deceasedBadge) deceasedBadge.style.display = 'none';
    const profileContent = document.getElementById('profileContent');
    if (profileContent) profileContent.classList.remove('contact-deceased');
    
    closeOppPanel();
  };
  
  window.cancelNewContact = function() {
    if (state.currentContactRecord) {
      selectContact(state.currentContactRecord);
    } else {
      toggleProfileView(false);
      disableAllFieldEditing();
    }
  };
  
  window.cancelEditMode = function() { cancelNewContact(); };
  
  // handleFormSubmit is defined in app.js - it uses processForm() API
  // which properly handles both create and update operations
  
  // ============================================================
  // Refresh Contact
  // ============================================================
  
  window.refreshCurrentContact = function() {
    if (!state.currentContactRecord) return;
    const refreshBtn = document.getElementById('refreshBtn');
    if (refreshBtn) {
      refreshBtn.classList.add('spinning');
      setTimeout(() => refreshBtn.classList.remove('spinning'), 1000);
    }
    google.script.run.withSuccessHandler(function(record) {
      if (record && record.fields) {
        selectContact(record);
      }
    }).getContactById(state.currentContactRecord.id);
  };
  
  // ============================================================
  // Delete Contact
  // ============================================================
  
  window.confirmDeleteContact = function() {
    openModal('deleteConfirmModal');
  };
  
  window.closeDeleteConfirmModal = function() {
    closeModal('deleteConfirmModal');
  };
  
  window.executeDeleteContact = function() {
    const recordId = state.currentContactRecord?.id;
    if (!recordId) return;
    
    closeModal('deleteConfirmModal');
    
    google.script.run.withSuccessHandler(function(result) {
      if (result.success) {
        state.currentContactRecord = null;
        toggleProfileView(false);
        loadContacts();
      } else {
        showAlert('Error', result.error || 'Failed to delete contact', 'error');
      }
    }).deleteContact(recordId);
  };
  
  // ============================================================
  // Formatting Helpers
  // ============================================================
  
  window.formatName = function(f) {
    const first = f.FirstName || '';
    const middle = f.MiddleName || '';
    const last = f.LastName || '';
    return [first, middle, last].filter(Boolean).join(' ') || 'Unnamed';
  };
  
  window.formatDetailsRow = function(f) {
    const parts = [];
    if (f.EmailAddress1) parts.push(f.EmailAddress1);
    if (f.Mobile) parts.push(f.Mobile);
    return parts.join(' | ') || '';
  };
  
  window.formatSubtitle = function(f) {
    const parts = [];
    if (f.PreferredName) parts.push(`"${f.PreferredName}"`);
    const tenure = formatTenureText(f);
    if (tenure) parts.push(tenure);
    return parts.join(' · ');
  };
  
  window.formatTenureText = function(f) {
    const createdOn = f['Created On (pre-Stellaris)'] || f['Created On'];
    if (!createdOn) return '';
    const tenure = calculateTenure(createdOn);
    return tenure ? `Client since ${tenure}` : '';
  };
  
  window.calculateTenure = function(createdStr) {
    if (!createdStr) return '';
    const created = new Date(createdStr);
    if (isNaN(created.getTime())) return '';
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    return `${months[created.getMonth()]} ${created.getFullYear()}`;
  };
  
  window.parseModifiedFormula = function(modifiedStr) {
    if (!modifiedStr) return null;
    const match = modifiedStr.match(/^(.+?), ([\d\/]+) at ([\d:]+[ap]m)$/i);
    if (match) {
      return { name: match[1].trim(), date: match[2], time: match[3] };
    }
    return null;
  };
  
  window.formatModifiedTooltip = function(f) {
    const modifiedFormula = f['Modified'];
    const parsed = parseModifiedFormula(modifiedFormula);
    if (parsed) {
      return `${parsed.name} on ${parsed.date} at ${parsed.time}`;
    }
    const modifiedOn = f['Modified On'];
    const modifiedBy = f['Modified By'];
    if (modifiedOn) {
      const date = formatAuditDate(modifiedOn);
      return modifiedBy ? `${modifiedBy} on ${date}` : date;
    }
    return '';
  };
  
  window.formatModifiedShort = function(f) {
    const modifiedFormula = f['Modified'];
    const parsed = parseModifiedFormula(modifiedFormula);
    if (parsed) {
      return `${parsed.date}`;
    }
    const modifiedOn = f['Modified On'];
    if (modifiedOn) {
      return formatAuditDate(modifiedOn).split(',')[0];
    }
    return '';
  };
  
  // ============================================================
  // Contact Meta Bar
  // ============================================================
  
  window.renderContactMetaBar = function(f) {
    // Render dossier meta (right side of header)
    const dossierMeta = document.getElementById('dossierMeta');
    if (dossierMeta) {
      const createdOn = f['Created On (pre-Stellaris)'] || f['Created On'];
      const createdBy = f['Created By (pre-Stellaris)'] || f['Created By'];
      const modifiedTooltip = formatModifiedTooltip(f);
      const modifiedShort = formatModifiedShort(f);
      
      let html = '';
      if (createdOn) {
        const createdDate = formatAuditDate(createdOn);
        html += `<span class="meta-item" title="Created by ${createdBy || 'Unknown'}">Created ${createdDate}</span>`;
      }
      if (modifiedShort) {
        html += `<span class="meta-item" title="${modifiedTooltip}">Modified ${modifiedShort}</span>`;
      }
      dossierMeta.innerHTML = html;
    }
    
    // Render status badge
    const statusBadge = document.getElementById('statusBadge');
    if (statusBadge) {
      const status = f.Status || 'Active';
      statusBadge.textContent = status;
      statusBadge.className = 'status-badge clickable-badge ' + (status === 'Active' ? 'status-active' : 'status-inactive');
      statusBadge.style.display = 'inline-block';
    }
    
    // Render marketing badge
    const marketingBadge = document.getElementById('marketingBadge');
    if (marketingBadge) {
      const isUnsubscribed = f['Unsubscribed from Marketing'] || false;
      const marketingText = isUnsubscribed ? 'UNSUBSCRIBED' : 'SUBSCRIBED';
      marketingBadge.textContent = marketingText;
      marketingBadge.style.display = 'inline-block';
    }
  };
  
})();
