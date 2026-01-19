/**
 * UI Utilities Module
 * Modal management, alerts, confirmations, and HTML escape functions
 */
(function() {
  'use strict';
  
  // ============================================================
  // HTML Escape/Unescape Functions
  // ============================================================
  
  window.escapeHtml = function(str) {
    if (!str) return '';
    return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  };
  
  window.escapeHtmlForAttr = function(str) {
    if (!str) return '';
    return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;').replace(/\n/g, '&#10;');
  };
  
  window.unescapeHtml = function(str) {
    if (!str) return '';
    const doc = new DOMParser().parseFromString(str, 'text/html');
    return doc.documentElement.textContent;
  };
  
  // ============================================================
  // Modal Management
  // ============================================================
  
  window.openModal = function(modalId) {
    const modal = document.getElementById(modalId);
    if (modal) {
      if (modal.classList.contains('modal-overlay')) {
        modal.style.display = 'flex';
        setTimeout(() => modal.classList.add('showing'), 10);
      } else {
        modal.classList.add('visible');
      }
    }
  };
  
  window.closeModal = function(modalId, callback) {
    const modal = document.getElementById(modalId);
    if (modal) {
      if (modal.classList.contains('modal-overlay')) {
        modal.classList.remove('showing');
        setTimeout(() => { modal.style.display = 'none'; if (callback) callback(); }, 250);
      } else {
        modal.classList.remove('visible');
        if (callback) callback();
      }
    }
  };
  
  // ============================================================
  // Confirmation Modal
  // ============================================================
  
  window.showConfirmModal = function(message, onConfirm) {
    const modal = document.getElementById('confirmModal');
    const msgEl = document.getElementById('confirmModalMessage');
    const yesBtn = document.getElementById('confirmModalYes');
    
    if (!modal || !msgEl || !yesBtn) {
      if (confirm(message)) onConfirm();
      return;
    }
    
    msgEl.textContent = message;
    
    const newYesBtn = yesBtn.cloneNode(true);
    yesBtn.parentNode.replaceChild(newYesBtn, yesBtn);
    newYesBtn.id = 'confirmModalYes';
    
    newYesBtn.addEventListener('click', function() {
      closeConfirmModal();
      onConfirm();
    });
    
    modal.style.display = 'flex';
    setTimeout(() => modal.classList.add('showing'), 10);
  };
  
  window.closeConfirmModal = function() {
    const modal = document.getElementById('confirmModal');
    if (modal) {
      modal.classList.remove('showing');
      setTimeout(() => modal.style.display = 'none', 250);
    }
  };
  
  // ============================================================
  // Alert Modal
  // ============================================================
  
  window.showAlert = function(title, message, type) {
    const modal = document.getElementById('alertModal');
    const titleEl = document.getElementById('alertModalTitle');
    const msgEl = document.getElementById('alertModalMessage');
    
    if (!modal) {
      alert(message);
      return;
    }
    
    if (titleEl) titleEl.textContent = title || 'Alert';
    if (msgEl) msgEl.textContent = message;
    
    modal.classList.remove('alert-success', 'alert-error', 'alert-warning', 'alert-info');
    if (type) modal.classList.add('alert-' + type);
    
    modal.style.display = 'flex';
    setTimeout(() => modal.classList.add('showing'), 10);
  };
  
  window.closeAlertModal = function() {
    const modal = document.getElementById('alertModal');
    if (modal) {
      modal.classList.remove('showing');
      setTimeout(() => modal.style.display = 'none', 250);
    }
  };
  
  // ============================================================
  // Utility Functions
  // ============================================================
  
  window.getInitials = function(firstName, lastName) {
    const f = (firstName || '').charAt(0).toUpperCase();
    const l = (lastName || '').charAt(0).toUpperCase();
    return f + l || '?';
  };
  
  window.getAvatarColor = function(name) {
    if (!name) return '#888';
    const colors = ['#2C2622', '#19414C', '#BB9934', '#7B8B64', '#5D7A8C', '#8B6B5D', '#6B7B5D', '#7B5D8B'];
    let hash = 0;
    for (let i = 0; i < name.length; i++) {
      hash = name.charCodeAt(i) + ((hash << 5) - hash);
    }
    return colors[Math.abs(hash) % colors.length];
  };
  
  window.autoExpandTextarea = function(el) {
    if (!el) return;
    el.style.height = 'auto';
    el.style.height = (el.scrollHeight) + 'px';
  };
  
  window.autoResizeTextarea = function(textarea) {
    if (!textarea) return;
    const originalDisplay = textarea.style.display;
    if (originalDisplay === 'none') {
      textarea.style.visibility = 'hidden';
      textarea.style.display = 'block';
    }
    textarea.style.height = 'auto';
    const newHeight = Math.max(textarea.scrollHeight, 60);
    textarea.style.height = newHeight + 'px';
    if (originalDisplay === 'none') {
      textarea.style.display = 'none';
      textarea.style.visibility = '';
    }
  };
  
  // ============================================================
  // Date Formatting
  // ============================================================
  
  window.formatDateDisplay = function(isoDate) {
    if (!isoDate) return '';
    const d = new Date(isoDate);
    if (isNaN(d.getTime())) return '';
    const day = String(d.getDate()).padStart(2, '0');
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const year = d.getFullYear();
    return `${day}/${month}/${year}`;
  };
  
  window.parseDateInput = function(value) {
    if (!value) return null;
    const parts = value.split('/');
    if (parts.length !== 3) return null;
    const [day, month, year] = parts;
    return `${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`;
  };
  
  window.formatAuditDate = function(dateStr) {
    if (!dateStr) return '';
    const d = new Date(dateStr);
    if (isNaN(d.getTime())) return dateStr;
    
    const pad = n => String(n).padStart(2, '0');
    const day = pad(d.getDate());
    const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const month = monthNames[d.getMonth()];
    const year = d.getFullYear();
    
    let hours = d.getHours();
    const mins = pad(d.getMinutes());
    const ampm = hours >= 12 ? 'pm' : 'am';
    hours = hours % 12 || 12;
    
    return `${day} ${month} ${year}, ${hours}:${mins}${ampm}`;
  };
  
  window.formatDatetimeForInput = function(dateStr) {
    if (!dateStr) return '';
    const d = new Date(dateStr);
    if (isNaN(d.getTime())) return '';
    const pad = n => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
  };
  
  window.formatDatetimeForDisplay = function(dateStr) {
    if (!dateStr) return '';
    const d = new Date(dateStr);
    if (isNaN(d.getTime())) return '';
    const pad = n => String(n).padStart(2, '0');
    const day = pad(d.getDate());
    const month = pad(d.getMonth() + 1);
    const year = d.getFullYear();
    let hours = d.getHours();
    const mins = pad(d.getMinutes());
    const ampm = hours >= 12 ? 'pm' : 'am';
    hours = hours % 12 || 12;
    return `${day}/${month}/${year} ${hours}:${mins}${ampm}`;
  };
  
})();
