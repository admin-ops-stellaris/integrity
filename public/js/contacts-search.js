/**
 * Contacts Search Module
 * Handles contact search, display, and keyboard navigation
 * 
 * IMPORTANT: This must be loaded AFTER foundation modules (shared-state, shared-utils, modal-utils)
 * Uses window.IntegrityState for state management
 */
(function() {
  'use strict';
  
  const state = window.IntegrityState;
  
  // ============================================================
  // Search Dropdown Show/Hide
  // ============================================================
  
  window.showSearchDropdown = function() {
    const dropdown = document.getElementById('searchDropdown');
    if (dropdown) dropdown.classList.add('open');
  };
  
  window.hideSearchDropdown = function() {
    const dropdown = document.getElementById('searchDropdown');
    if (dropdown) dropdown.classList.remove('open');
  };
  
  // ============================================================
  // Avatar Helpers
  // ============================================================
  
  window.getInitials = function(firstName, lastName) {
    const f = (firstName || '').charAt(0).toUpperCase();
    const l = (lastName || '').charAt(0).toUpperCase();
    return f + l || '?';
  };
  
  window.getAvatarColor = function(name) {
    const colors = ['#19414C', '#7B8B64', '#BB9934', '#2C2622', '#6B5B4F'];
    let hash = 0;
    for (let i = 0; i < name.length; i++) hash = name.charCodeAt(i) + ((hash << 5) - hash);
    return colors[Math.abs(hash) % colors.length];
  };
  
  // ============================================================
  // Name Formatting
  // ============================================================
  
  window.formatName = function(f) {
    return `${f.FirstName || ''} ${f.MiddleName || ''} ${f.LastName || ''}`.replace(/\s+/g, ' ').trim();
  };
  
  window.formatDetailsRow = function(f) {
    const parts = [];
    if (f.EmailAddress1) parts.push(`<span>${f.EmailAddress1}</span>`);
    if (f.Mobile) parts.push(`<span>${f.Mobile}</span>`);
    return parts.join('');
  };
  
  // ============================================================
  // Modified Date Parsing & Formatting (Perth timezone)
  // ============================================================
  
  window.parseModifiedFormula = function(modifiedStr) {
    if (!modifiedStr) return null;
    const match = modifiedStr.match(/(\d{2}):(\d{2})\s+(\d{2})\/(\d{2})\/(\d{4})/);
    if (!match) return null;
    const [, hours, mins, day, month, year] = match;
    const utcMs = Date.UTC(parseInt(year), parseInt(month) - 1, parseInt(day), parseInt(hours), parseInt(mins));
    const perthOffsetMs = 8 * 60 * 60 * 1000;
    return new Date(utcMs - perthOffsetMs);
  };
  
  window.formatModifiedTooltip = function(f) {
    const modified = f.Modified;
    if (!modified) return null;
    
    const modDate = window.parseModifiedFormula(modified);
    if (!modDate) return null;
    
    const byMatch = modified.match(/by\s+(.+)$/);
    const modifiedBy = byMatch ? byMatch[1] : null;
    
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
  };
  
  window.formatModifiedShort = function(f) {
    const modified = f.Modified;
    if (!modified) return null;
    
    const modDate = window.parseModifiedFormula(modified);
    if (!modDate) return null;
    
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
  };
  
  // ============================================================
  // Contact Search Handler
  // ============================================================
  
  window.handleSearch = function(event) {
    if (['ArrowDown', 'ArrowUp', 'Enter', 'Escape'].includes(event.key)) return;
    
    const query = event.target.value;
    const statusEl = document.getElementById('searchStatus');
    
    clearTimeout(state.loadingTimer);
    state.searchHighlightIndex = -1;
    window.showSearchDropdown();
    
    if (query.length === 0) {
      statusEl.innerText = "";
      window.loadContacts();
      return;
    }
    
    clearTimeout(state.searchTimeout);
    statusEl.innerText = "Typing...";
    
    state.searchTimeout = setTimeout(() => {
      statusEl.innerText = "Searching...";
      const statusFilterToSend = state.contactStatusFilter === 'All' ? null : state.contactStatusFilter;
      google.script.run.withSuccessHandler(function(records) {
        statusEl.innerText = records.length > 0 ? `Found ${records.length} matches` : "No matches found";
        window.renderList(records);
      }).searchContacts(query, statusFilterToSend);
    }, 500);
  };
  
  // ============================================================
  // Load Contacts (Recent/Default)
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
      if (!records || records.length === 0) {
        list.innerHTML = '<li style="padding:20px; color:#999; text-align:center; font-size:13px;">No contacts found</li>';
        return;
      }
      window.renderList(records);
    }).getRecentContacts(statusFilterToSend);
  };
  
  // ============================================================
  // Render Contact List
  // ============================================================
  
  window.renderList = function(records) {
    const list = document.getElementById('contactList');
    document.getElementById('loading').style.display = 'none';
    list.innerHTML = '';
    
    state.currentSearchRecords = records;
    
    if (state.searchHighlightIndex >= records.length) {
      state.searchHighlightIndex = records.length > 0 ? records.length - 1 : -1;
    }
    
    records.forEach((record, index) => {
      const f = record.fields;
      const item = document.createElement('li');
      item.className = 'contact-item';
      item.dataset.index = index;
      
      const fullName = window.formatName(f);
      const initials = window.getInitials(f.FirstName, f.LastName);
      const avatarColor = window.getAvatarColor(fullName);
      const modifiedTooltip = window.formatModifiedTooltip(f);
      const modifiedShort = window.formatModifiedShort(f);
      const isDeceased = f.Deceased === true;
      const deceasedBadge = isDeceased ? '<span class="deceased-badge-small">DECEASED</span>' : '';
      
      item.innerHTML = `<div class="contact-avatar" style="background-color:${avatarColor}">${initials}</div><div class="contact-info"><span class="contact-name">${fullName}${deceasedBadge}</span><div class="contact-details-row">${window.formatDetailsRow(f)}</div></div>${modifiedShort ? `<span class="contact-modified" title="${modifiedTooltip || ''}">${modifiedShort}</span>` : ''}`;
      
      if (isDeceased) item.style.opacity = '0.6';
      
      item.onclick = function() { window.selectContact(record); };
      list.appendChild(item);
    });
    
    if (state.searchHighlightIndex >= 0) {
      window.updateSearchHighlight();
    }
  };
  
  // ============================================================
  // Search Highlight Navigation
  // ============================================================
  
  window.updateSearchHighlight = function() {
    const items = document.querySelectorAll('#contactList .contact-item');
    items.forEach((item, i) => {
      item.classList.toggle('keyboard-highlight', i === state.searchHighlightIndex);
    });
    if (state.searchHighlightIndex >= 0 && items[state.searchHighlightIndex]) {
      items[state.searchHighlightIndex].scrollIntoView({ block: 'nearest', behavior: 'smooth' });
    }
  };
  
  window.handleSearchKeydown = function(e) {
    const dropdown = document.getElementById('searchDropdown');
    if (!dropdown || !dropdown.classList.contains('open')) return;
    
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      if (state.currentSearchRecords.length > 0) {
        state.searchHighlightIndex = Math.min(state.searchHighlightIndex + 1, state.currentSearchRecords.length - 1);
        window.updateSearchHighlight();
      }
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      if (state.currentSearchRecords.length > 0) {
        state.searchHighlightIndex = Math.max(state.searchHighlightIndex - 1, 0);
        window.updateSearchHighlight();
      }
    } else if (e.key === 'Enter') {
      e.preventDefault();
      if (state.searchHighlightIndex >= 0 && state.currentSearchRecords[state.searchHighlightIndex]) {
        window.selectContact(state.currentSearchRecords[state.searchHighlightIndex]);
      } else {
        const searchInput = document.getElementById('searchInput');
        if (searchInput && searchInput.value.trim()) {
          clearTimeout(state.searchTimeout);
          const statusEl = document.getElementById('searchStatus');
          statusEl.innerText = "Searching...";
          const statusFilterToSend = state.contactStatusFilter === 'All' ? null : state.contactStatusFilter;
          google.script.run.withSuccessHandler(function(records) {
            statusEl.innerText = records.length > 0 ? `Found ${records.length} matches` : "No matches found";
            window.renderList(records);
          }).searchContacts(searchInput.value.trim(), statusFilterToSend);
        }
      }
    }
  };
  
  // ============================================================
  // Close dropdown when clicking outside (event listener)
  // ============================================================
  
  document.addEventListener('click', function(e) {
    const wrapper = document.querySelector('.header-search-wrapper');
    if (wrapper && !wrapper.contains(e.target)) {
      window.hideSearchDropdown();
    }
  });
  
})();
