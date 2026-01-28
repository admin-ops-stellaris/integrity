/**
 * Modal Utilities Module
 * Functions for opening/closing modals, alerts, and confirmations
 * 
 * IMPORTANT: This must be loaded THIRD, after shared-state.js and shared-utils.js
 * These modal functions are used by many other modules
 */
(function() {
  'use strict';
  
  // ============================================================
  // Generic Modal Open/Close
  // ============================================================
  
  window.openModal = function(modalId) {
    const modal = document.getElementById(modalId);
    if (!modal) return;
    modal.classList.add('visible');
    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        modal.classList.add('showing');
      });
    });
  };
  
  window.closeModal = function(modalId, callback) {
    const modal = document.getElementById(modalId);
    if (!modal) {
      if (callback) callback();
      return;
    }
    modal.classList.remove('showing');
    setTimeout(() => {
      modal.classList.remove('visible');
      if (callback) callback();
    }, 250);
  };
  
  // ============================================================
  // Confirmation Modal (generic yes/no)
  // ============================================================
  
  window.showConfirmModal = function(message, onConfirm) {
    const modal = document.getElementById('confirmModal');
    const msgEl = document.getElementById('confirmModalMessage');
    const okBtn = document.getElementById('confirmModalOk');
    const cancelBtn = document.getElementById('confirmModalCancel');
    
    if (!modal || !msgEl) {
      if (confirm(message) && onConfirm) onConfirm();
      return;
    }
    
    msgEl.textContent = message;
    modal.style.display = 'flex';
    
    const cleanup = () => {
      modal.style.display = 'none';
      if (okBtn) okBtn.onclick = null;
      if (cancelBtn) cancelBtn.onclick = null;
    };
    
    if (okBtn) okBtn.onclick = () => { cleanup(); if (onConfirm) onConfirm(); };
    if (cancelBtn) cancelBtn.onclick = cleanup;
  };
  
  window.closeConfirmModal = function() {
    const modal = document.getElementById('confirmModal');
    if (modal) modal.style.display = 'none';
  };
  
  // Alias for consistency - use this name going forward
  window.showCustomConfirm = window.showConfirmModal;
  
  // ============================================================
  // Alert Modal
  // ============================================================
  
  window.showAlert = function(title, message, type) {
    const modal = document.getElementById('alertModal');
    const sidebar = document.getElementById('alertModalSidebar');
    const icon = document.getElementById('alertModalIcon');
    
    if (!modal) {
      alert(`${title}: ${message}`);
      return;
    }
    
    const titleEl = document.getElementById('alertModalTitle');
    const msgEl = document.getElementById('alertModalMessage');
    
    if (titleEl) titleEl.innerText = title;
    if (msgEl) msgEl.innerText = message;
    
    if (sidebar && icon) {
      if (type === 'success') {
        sidebar.style.background = 'var(--color-cedar)';
        icon.innerText = '✓';
      } else if (type === 'error') {
        sidebar.style.background = '#A00';
        icon.innerText = '✕';
      } else {
        sidebar.style.background = 'var(--color-star)';
        icon.innerText = 'ℹ';
      }
    }
    
    window.openModal('alertModal');
  };
  
  window.closeAlertModal = function() {
    window.closeModal('alertModal');
  };
  
  // ============================================================
  // Delete Confirmation Modal (Contact)
  // ============================================================
  
  window.closeDeleteConfirmModal = function() {
    window.closeModal('deleteConfirmModal');
  };
  
  // ============================================================
  // Delete Confirmation Modal (Opportunity)
  // ============================================================
  
  window.closeDeleteOppConfirmModal = function() {
    window.closeModal('deleteOppConfirmModal');
    const state = window.IntegrityState;
    if (state) state.currentOppToDelete = null;
  };
  
})();
