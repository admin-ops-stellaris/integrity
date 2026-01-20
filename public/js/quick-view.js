/**
 * quick-view.js - Contact Quick View Module
 * 
 * Displays contact summary cards on hover with mobile/email details.
 * Includes positioning, event delegation, and navigation.
 * 
 * Dependencies: shared-state.js, contacts-search.js (getInitials, getAvatarColor)
 * 
 * Functions exposed to window:
 * - showContactQuickView, hideContactQuickView
 * - attachQuickViewToElement, navigateFromQuickView
 */
(function() {
  'use strict';
  
  const state = window.IntegrityState;

  state.quickViewContactId = null;
  state.quickViewHoverTimeout = null;
  state.isQuickViewHovered = false;

  function showContactQuickView(contactId, triggerElement) {
    if (!contactId) return;
    state.quickViewContactId = contactId;
    
    const card = document.getElementById('contactQuickView');
    // Store contact ID on card element for button click access
    card.dataset.contactId = contactId;
    const rect = triggerElement.getBoundingClientRect();
    
    let left = rect.left + (rect.width / 2) - 150;
    let top = rect.bottom + 8;
    
    if (left < 10) left = 10;
    if (left + 320 > window.innerWidth - 10) left = window.innerWidth - 330;
    if (top + 300 > window.innerHeight) {
      top = rect.top - 8;
      card.style.transform = 'translateY(-100%)';
    } else {
      card.style.transform = 'translateY(0)';
    }
    
    card.style.left = left + 'px';
    card.style.top = top + 'px';
    
    document.getElementById('quickViewName').textContent = 'Loading...';
    document.getElementById('quickViewPreferred').textContent = '';
    document.getElementById('quickViewAvatar').textContent = '...';
    document.getElementById('quickViewAvatar').style.backgroundColor = '#999';
    document.getElementById('quickViewDetails').innerHTML = '';
    document.getElementById('quickViewFooter').textContent = '';
    
    card.classList.add('visible');
    
    google.script.run.withSuccessHandler(function(contact) {
      if (!contact || !contact.fields || state.quickViewContactId !== contactId) return;
      
      const f = contact.fields;
      const fullName = f['Calculated Name'] || 
        `${f.FirstName || ''} ${f.MiddleName || ''} ${f.LastName || ''}`.replace(/\s+/g, ' ').trim();
      const initials = getInitials(f.FirstName, f.LastName);
      const avatarColor = getAvatarColor(fullName);
      
      document.getElementById('quickViewName').textContent = fullName;
      document.getElementById('quickViewAvatar').textContent = initials;
      document.getElementById('quickViewAvatar').style.backgroundColor = avatarColor;
      
      const preferred = f.PreferredName;
      if (preferred && preferred.toLowerCase() !== (f.FirstName || '').toLowerCase()) {
        document.getElementById('quickViewPreferred').textContent = `Preferred: ${preferred}`;
      } else {
        document.getElementById('quickViewPreferred').textContent = '';
      }
      
      let detailsHtml = '';
      if (f.Mobile) {
        detailsHtml += `<div class="quick-view-detail-row"><span class="quick-view-detail-icon">📱</span><span class="quick-view-detail-value">${f.Mobile}</span></div>`;
      }
      if (f.EmailAddress1) {
        detailsHtml += `<div class="quick-view-detail-row"><span class="quick-view-detail-icon">✉️</span><span class="quick-view-detail-value">${f.EmailAddress1}</span></div>`;
      }
      if (!f.Mobile && !f.EmailAddress1) {
        detailsHtml = '<div class="quick-view-no-details">No contact details available</div>';
      }
      document.getElementById('quickViewDetails').innerHTML = detailsHtml;
      
      const modifiedOn = f['Modified On (Web App)'] || f['Last Modified'];
      if (modifiedOn) {
        const modDate = new Date(modifiedOn);
        const now = new Date();
        const diffMs = now - modDate;
        const diffMins = Math.floor(diffMs / 60000);
        const diffHours = Math.floor(diffMs / 3600000);
        const diffDays = Math.floor(diffMs / 86400000);
        
        let relativeTime;
        if (diffMins < 1) relativeTime = 'Just now';
        else if (diffMins < 60) relativeTime = `${diffMins} minute${diffMins > 1 ? 's' : ''} ago`;
        else if (diffHours < 24) relativeTime = `${diffHours} hour${diffHours > 1 ? 's' : ''} ago`;
        else if (diffDays === 1) relativeTime = 'Yesterday';
        else if (diffDays < 7) relativeTime = `${diffDays} days ago`;
        else {
          const day = String(modDate.getDate()).padStart(2, '0');
          const month = modDate.toLocaleString('en', { month: 'short' });
          const year = modDate.getFullYear();
          relativeTime = `${day} ${month} ${year}`;
        }
        document.getElementById('quickViewFooter').textContent = `Last modified: ${relativeTime}`;
      } else {
        document.getElementById('quickViewFooter').textContent = '';
      }
    }).getContactById(contactId);
  }

  function hideContactQuickView() {
    if (state.isQuickViewHovered) return;
    const card = document.getElementById('contactQuickView');
    card.classList.remove('visible');
    state.quickViewContactId = null;
  }

  function attachQuickViewToElement(element, contactId) {
    element.setAttribute('data-quick-view-contact', contactId);
    element.addEventListener('mouseenter', function(e) {
      if (state.quickViewHoverTimeout) clearTimeout(state.quickViewHoverTimeout);
      state.quickViewHoverTimeout = setTimeout(function() {
        showContactQuickView(contactId, element);
      }, 300);
    });
    element.addEventListener('mouseleave', function(e) {
      if (state.quickViewHoverTimeout) clearTimeout(state.quickViewHoverTimeout);
      state.quickViewHoverTimeout = setTimeout(hideContactQuickView, 200);
    });
  }

  function navigateFromQuickView(id) {
    // Accept ID as parameter (from button click) or fall back to state
    const contactId = id || state.quickViewContactId;
    console.log('Navigate requested for:', contactId);
    
    if (!contactId) {
      console.error('No contact ID available for navigation');
      return;
    }
    
    state.isQuickViewHovered = false;
    hideContactQuickView();
    
    // Use window.loadContactById which is exposed from app.js
    if (typeof window.loadContactById === 'function') {
      window.loadContactById(contactId, true);
    } else {
      console.error('loadContactById not available - app.js may not be loaded');
    }
  }

  document.addEventListener('DOMContentLoaded', function() {
    const card = document.getElementById('contactQuickView');
    if (card) {
      card.addEventListener('mouseenter', function() {
        state.isQuickViewHovered = true;
        if (state.quickViewHoverTimeout) clearTimeout(state.quickViewHoverTimeout);
      });
      card.addEventListener('mouseleave', function() {
        state.isQuickViewHovered = false;
        state.quickViewHoverTimeout = setTimeout(hideContactQuickView, 200);
      });
    }
  });

  document.addEventListener('keydown', function(e) {
    if (e.key === 'Escape') {
      const card = document.getElementById('contactQuickView');
      if (card && card.classList.contains('visible')) {
        state.isQuickViewHovered = false;
        hideContactQuickView();
      }
    }
  });

  document.addEventListener('click', function(e) {
    const card = document.getElementById('contactQuickView');
    if (card && card.classList.contains('visible')) {
      if (!card.contains(e.target) && !e.target.hasAttribute('data-quick-view-contact')) {
        state.isQuickViewHovered = false;
        hideContactQuickView();
      }
    }
  });

  document.addEventListener('mouseover', function(e) {
    const link = e.target.closest('.panel-contact-link');
    if (link && !link.hasAttribute('data-quick-view-contact')) {
      const contactId = link.getAttribute('data-contact-id');
      if (contactId) {
        attachQuickViewToElement(link, contactId);
        if (state.quickViewHoverTimeout) clearTimeout(state.quickViewHoverTimeout);
        state.quickViewHoverTimeout = setTimeout(function() {
          showContactQuickView(contactId, link);
        }, 300);
      }
    }
  });

  window.showContactQuickView = showContactQuickView;
  window.hideContactQuickView = hideContactQuickView;
  window.attachQuickViewToElement = attachQuickViewToElement;
  window.navigateFromQuickView = navigateFromQuickView;

})();
