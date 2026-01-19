/**
 * Quick View Module
 * Contact quick view hover functionality
 */
(function() {
  'use strict';
  
  const state = window.IntegrityState;
  
  let quickViewTimeout = null;
  let hideQuickViewTimeout = null;
  
  // ============================================================
  // Show Contact Quick View
  // ============================================================
  
  window.showContactQuickView = function(contactId, triggerElement) {
    clearTimeout(hideQuickViewTimeout);
    
    // Delay showing to prevent flicker
    quickViewTimeout = setTimeout(() => {
      google.script.run.withSuccessHandler(function(record) {
        if (!record || !record.fields) return;
        
        const f = record.fields;
        const quickView = document.getElementById('contactQuickView');
        
        if (!quickView) return;
        
        quickView.dataset.contactId = contactId;
        
        const initials = getInitials(f.FirstName, f.LastName);
        const avatarColor = getAvatarColor((f.FirstName || '') + (f.LastName || ''));
        
        quickView.innerHTML = `
          <div class="quick-view-header">
            <div class="quick-view-avatar" style="background-color:${avatarColor};">${initials}</div>
            <div class="quick-view-name">${escapeHtml(formatName(f))}</div>
          </div>
          <div class="quick-view-details">
            ${f.EmailAddress1 ? `<div class="quick-view-row"><span class="quick-view-label">Email:</span> ${escapeHtml(f.EmailAddress1)}</div>` : ''}
            ${f.Mobile ? `<div class="quick-view-row"><span class="quick-view-label">Mobile:</span> ${escapeHtml(f.Mobile)}</div>` : ''}
            ${f.Telephone1 ? `<div class="quick-view-row"><span class="quick-view-label">Home:</span> ${escapeHtml(f.Telephone1)}</div>` : ''}
          </div>
          <div class="quick-view-actions">
            <button class="quick-view-btn" onclick="navigateFromQuickView()">View Contact</button>
          </div>
        `;
        
        // Position quick view
        const rect = triggerElement.getBoundingClientRect();
        quickView.style.top = (rect.bottom + 5) + 'px';
        quickView.style.left = Math.min(rect.left, window.innerWidth - 280) + 'px';
        quickView.classList.add('visible');
        
      }).getContactById(contactId);
    }, 300);
  };
  
  // ============================================================
  // Hide Contact Quick View
  // ============================================================
  
  window.hideContactQuickView = function() {
    clearTimeout(quickViewTimeout);
    
    hideQuickViewTimeout = setTimeout(() => {
      const quickView = document.getElementById('contactQuickView');
      if (quickView) {
        quickView.classList.remove('visible');
      }
    }, 200);
  };
  
  // ============================================================
  // Attach Quick View to Element
  // ============================================================
  
  window.attachQuickViewToElement = function(element, contactId) {
    element.addEventListener('mouseenter', function(e) {
      showContactQuickView(contactId, element);
    });
    
    element.addEventListener('mouseleave', function(e) {
      hideContactQuickView();
    });
  };
  
  // Keep quick view visible when hovering over it
  document.addEventListener('DOMContentLoaded', function() {
    const quickView = document.getElementById('contactQuickView');
    if (quickView) {
      quickView.addEventListener('mouseenter', function() {
        clearTimeout(hideQuickViewTimeout);
      });
      
      quickView.addEventListener('mouseleave', function() {
        hideContactQuickView();
      });
    }
  });
  
})();
