/**
 * Inline Editing Module
 * InlineEditingManager for click-to-edit form fields
 */
(function() {
  'use strict';
  
  const state = window.IntegrityState;
  
  // ============================================================
  // InlineEditingManager IIFE
  // ============================================================
  
  const InlineEditingManager = (function() {
    const editState = {};
    const sessionIds = {};
    let currentField = null;
    let isBulkEditMode = false;
    
    function generateSessionId() {
      return Date.now().toString(36) + Math.random().toString(36).substr(2);
    }
    
    function getCompositeKey(fieldKey) {
      return `${fieldKey}_${sessionIds[fieldKey] || 'default'}`;
    }
    
    function startEdit(fieldKey) {
      sessionIds[fieldKey] = generateSessionId();
      editState[getCompositeKey(fieldKey)] = { editing: true, saving: false };
      currentField = fieldKey;
    }
    
    function isEditing(fieldKey) {
      const key = getCompositeKey(fieldKey);
      return editState[key]?.editing || false;
    }
    
    function isSaving(fieldKey) {
      const key = getCompositeKey(fieldKey);
      return editState[key]?.saving || false;
    }
    
    function setSaving(fieldKey, value) {
      const key = getCompositeKey(fieldKey);
      if (editState[key]) {
        editState[key].saving = value;
      }
    }
    
    function endEdit(fieldKey) {
      const key = getCompositeKey(fieldKey);
      delete editState[key];
      delete sessionIds[fieldKey];
      if (currentField === fieldKey) {
        currentField = null;
      }
    }
    
    function setBulkEditMode(enabled) {
      isBulkEditMode = enabled;
    }
    
    function isInBulkEditMode() {
      return isBulkEditMode;
    }
    
    function getCurrentField() {
      return currentField;
    }
    
    return {
      startEdit,
      isEditing,
      isSaving,
      setSaving,
      endEdit,
      setBulkEditMode,
      isInBulkEditMode,
      getCurrentField,
      getSessionId: (fieldKey) => sessionIds[fieldKey]
    };
  })();
  
  window.InlineEditingManager = InlineEditingManager;
  
  // ============================================================
  // Initialization
  // ============================================================
  
  window.initInlineEditing = function() {
    const profileColumns = document.getElementById('profileColumns');
    if (profileColumns) {
      profileColumns.addEventListener('click', handleFormClick);
    }
  };
  
  function handleFormClick(event) {
    const fieldGroup = event.target.closest('.field-group');
    if (fieldGroup && InlineEditingManager.isInBulkEditMode()) {
      return;
    }
  }
  
  // ============================================================
  // Enable/Disable Editing
  // ============================================================
  
  window.enableNewContactMode = function() {
    InlineEditingManager.setBulkEditMode(true);
    
    const inputs = document.querySelectorAll('#contactForm input, #contactForm select, #contactForm textarea');
    inputs.forEach(input => {
      if (input.id === 'recordId' || input.id === 'status') return;
      input.classList.remove('locked');
      input.removeAttribute('readonly');
      if (input.tagName === 'SELECT') {
        input.removeAttribute('disabled');
      }
    });
    
    document.getElementById('cancelBtn').style.display = 'inline-block';
    document.getElementById('submitBtn').style.display = 'inline-block';
    updateHeaderTitle(true);
  };
  
  window.disableAllFieldEditing = function() {
    InlineEditingManager.setBulkEditMode(false);
    
    const inputs = document.querySelectorAll('#contactForm input, #contactForm select, #contactForm textarea');
    inputs.forEach(input => {
      if (input.id === 'recordId') return;
      input.classList.add('locked');
      input.setAttribute('readonly', 'readonly');
      if (input.tagName === 'SELECT') {
        input.setAttribute('disabled', 'disabled');
      }
    });
    
    document.getElementById('cancelBtn').style.display = 'none';
    document.getElementById('submitBtn').style.display = 'none';
    document.getElementById('editBtn').style.visibility = 'visible';
    toggleProfileView(true);
    updateHeaderTitle(false);
  };
  
  window.enableEditMode = function() { enableNewContactMode(); };
  window.disableEditMode = function() { disableAllFieldEditing(); };
  
  // ============================================================
  // Contact Status Toggle
  // ============================================================
  
  window.toggleContactStatus = function() {
    const statusField = document.getElementById('status');
    if (!statusField || !state.currentContactRecord) return;
    
    const currentStatus = statusField.value;
    const newStatus = currentStatus === 'Active' ? 'Inactive' : 'Active';
    
    statusField.value = newStatus;
    
    google.script.run.withSuccessHandler(function(result) {
      if (result && result.id) {
        selectContact(result);
        loadContacts();
      }
    }).updateContact(state.currentContactRecord.id, { Status: newStatus });
  };
  
  // ============================================================
  // Gender Field Handling
  // ============================================================
  
  window.handleGenderChange = function() {
    const genderSelect = document.getElementById('gender');
    const genderOtherWrapper = document.getElementById('genderOtherWrapper');
    
    if (genderSelect && genderOtherWrapper) {
      const showOther = genderSelect.value === 'Other' || genderSelect.value === 'Prefer not to say';
      genderOtherWrapper.style.display = showOther ? 'block' : 'none';
    }
  };
  
  // ============================================================
  // Unsubscribe Display
  // ============================================================
  
  window.updateUnsubscribeDisplay = function(isUnsubscribed) {
    const wrapper = document.getElementById('unsubscribeWrapper');
    const displaySpan = wrapper?.querySelector('.unsubscribe-display');
    
    if (displaySpan) {
      if (isUnsubscribed) {
        displaySpan.innerHTML = '<span class="unsubscribed-badge">UNSUBSCRIBED</span>';
      } else {
        displaySpan.innerHTML = '<span class="subscribed-text">Subscribed to marketing</span>';
      }
    }
  };
  
  window.openUnsubscribeEdit = function() {
    const modal = document.getElementById('unsubscribeModal');
    const select = document.getElementById('unsubscribeModalSelect');
    if (modal && select) {
      select.value = document.getElementById('unsubscribeFromMarketing').value;
      openModal('unsubscribeModal');
    }
  };
  
  window.closeUnsubscribeModal = function() {
    closeModal('unsubscribeModal');
  };
  
  window.saveUnsubscribePreference = function() {
    const newValue = document.getElementById('unsubscribeModalSelect').value;
    const hiddenField = document.getElementById('unsubscribeFromMarketing');
    
    if (hiddenField) {
      hiddenField.value = newValue;
    }
    
    closeUnsubscribeModal();
    updateUnsubscribeDisplay(newValue === 'true');
    
    if (state.currentContactRecord) {
      google.script.run.withSuccessHandler(function(result) {
        if (result && result.id) {
          state.currentContactRecord = result;
        }
      }).updateContact(state.currentContactRecord.id, { UnsubscribeFromMarketing: newValue });
    }
  };
  
  // ============================================================
  // Deceased Styling
  // ============================================================
  
  window.applyDeceasedStyling = function(isDeceased) {
    const profileContent = document.getElementById('profileContent');
    const deceasedBadge = document.getElementById('deceasedBadge');
    
    if (profileContent) {
      profileContent.classList.toggle('contact-deceased', isDeceased);
    }
    if (deceasedBadge) {
      deceasedBadge.style.display = isDeceased ? 'inline-block' : 'none';
    }
  };
  
  // ============================================================
  // Actions Menu
  // ============================================================
  
  window.toggleActionsMenu = function() {
    const dropdown = document.getElementById('actionsDropdown');
    if (dropdown) {
      dropdown.classList.toggle('visible');
    }
  };
  
  document.addEventListener('click', function(e) {
    const wrapper = document.getElementById('actionsMenuWrapper');
    const dropdown = document.getElementById('actionsDropdown');
    if (wrapper && dropdown && !wrapper.contains(e.target)) {
      dropdown.classList.remove('visible');
    }
  });
  
  // ============================================================
  // Deceased Workflow
  // ============================================================
  
  window.markAsDeceased = function() {
    document.getElementById('actionsDropdown')?.classList.remove('visible');
    const modal = document.getElementById('deceasedConfirmModal');
    if (modal) {
      modal.style.display = 'flex';
      setTimeout(() => modal.classList.add('showing'), 10);
    }
  };
  
  window.undoDeceased = function() {
    if (!state.currentContactRecord) return;
    
    google.script.run.withSuccessHandler(function(result) {
      if (result && result.id) {
        selectContact(result);
      }
    }).updateContact(state.currentContactRecord.id, { Deceased: false });
  };
  
  window.closeDeceasedConfirmModal = function() {
    const modal = document.getElementById('deceasedConfirmModal');
    if (modal) {
      modal.classList.remove('showing');
      setTimeout(() => modal.style.display = 'none', 250);
    }
  };
  
  window.executeMarkDeceased = function() {
    if (!state.currentContactRecord) return;
    
    closeDeceasedConfirmModal();
    
    google.script.run.withSuccessHandler(function(result) {
      if (result && result.id) {
        selectContact(result);
      }
    }).updateContact(state.currentContactRecord.id, { 
      Deceased: true,
      UnsubscribeFromMarketing: 'true'
    });
  };
  
})();
