/**
 * Notes Module
 * Field note popovers and connection note popovers
 */
(function() {
  'use strict';
  
  const state = window.IntegrityState;
  
  // Note field configuration
  const NOTE_FIELDS = [
    { fieldId: 'email1Comment', airtableField: 'Email1Comment', inputId: 'email1' },
    { fieldId: 'email2Comment', airtableField: 'Email2Comment', inputId: 'email2' },
    { fieldId: 'email3Comment', airtableField: 'Email3Comment', inputId: 'email3' },
    { fieldId: 'genderOther', airtableField: 'GenderOther', inputId: 'gender' }
  ];
  
  // Popover state
  let activePopover = null;
  let activeNoteFieldId = null;
  let noteDebounceTimer = null;
  let activeConnectionPopover = null;
  let activeConnectionId = null;
  let connNoteDebounceTimer = null;
  
  // ============================================================
  // Initialize Note Icons
  // ============================================================
  
  window.initAllNoteFields = function() {
    NOTE_FIELDS.forEach(config => {
      const input = document.getElementById(config.inputId);
      if (!input) return;
      
      const wrapper = input.closest('.field-group');
      if (!wrapper) return;
      
      // Check if icon already exists
      if (wrapper.querySelector('.note-icon-btn')) return;
      
      initNoteIcon(wrapper, config.fieldId);
    });
  };
  
  function initNoteIcon(wrapper, fieldId) {
    const iconBtn = document.createElement('span');
    iconBtn.className = 'note-icon-btn';
    iconBtn.innerHTML = '&#9998;';
    iconBtn.title = 'Add note';
    iconBtn.dataset.fieldId = fieldId;
    iconBtn.onclick = function(e) {
      e.stopPropagation();
      openNotePopover(this, fieldId);
    };
    
    wrapper.style.position = 'relative';
    wrapper.appendChild(iconBtn);
  }
  
  // ============================================================
  // Populate Note Fields from Contact
  // ============================================================
  
  window.populateNoteFields = function(contact) {
    if (!contact || !contact.fields) return;
    
    NOTE_FIELDS.forEach(config => {
      const hiddenField = document.getElementById(config.fieldId);
      if (hiddenField) {
        hiddenField.value = contact.fields[config.airtableField] || '';
      }
    });
  };
  
  window.getNoteFieldValues = function() {
    const values = {};
    NOTE_FIELDS.forEach(config => {
      const hiddenField = document.getElementById(config.fieldId);
      if (hiddenField) {
        values[config.airtableField] = hiddenField.value;
      }
    });
    return values;
  };
  
  // ============================================================
  // Open Field Note Popover
  // ============================================================
  
  window.openNotePopover = function(iconBtn, fieldId) {
    // Close any existing popover
    closeNotePopover();
    
    const config = NOTE_FIELDS.find(c => c.fieldId === fieldId);
    if (!config) return;
    
    const hiddenField = document.getElementById(fieldId);
    const currentValue = hiddenField?.value || '';
    
    // Create popover
    const popover = document.createElement('div');
    popover.className = 'note-popover';
    popover.id = 'activeNotePopover';
    popover.innerHTML = `
      <div class="note-popover-header">
        <span class="note-popover-title">Note</span>
        <span class="note-popover-close" onclick="closeNotePopover()">&times;</span>
      </div>
      <textarea class="note-popover-textarea" id="notePopoverTextarea">${escapeHtml(currentValue)}</textarea>
      <div class="note-popover-footer">
        <span class="note-popover-status" id="notePopoverStatus"></span>
      </div>
    `;
    
    // Position popover
    const rect = iconBtn.getBoundingClientRect();
    popover.style.position = 'fixed';
    popover.style.top = (rect.bottom + 5) + 'px';
    popover.style.left = Math.min(rect.left, window.innerWidth - 260) + 'px';
    popover.style.zIndex = '10000';
    
    document.body.appendChild(popover);
    
    activePopover = popover;
    activeNoteFieldId = fieldId;
    
    // Focus textarea
    const textarea = document.getElementById('notePopoverTextarea');
    textarea.focus();
    textarea.selectionStart = textarea.value.length;
    
    // Auto-save on input
    textarea.addEventListener('input', function() {
      clearTimeout(noteDebounceTimer);
      document.getElementById('notePopoverStatus').textContent = 'Typing...';
      noteDebounceTimer = setTimeout(() => saveNoteFromPopover(false), 800);
    });
    
    // Keyboard shortcuts
    textarea.addEventListener('keydown', function(e) {
      if (e.key === 'Escape') {
        e.preventDefault();
        closeNotePopover();
      } else if (e.key === 'Tab' || e.key === 'Enter') {
        if (!e.shiftKey) {
          e.preventDefault();
          saveNoteFromPopover(true);
        }
      }
    });
    
    // Prevent text selection from closing
    textarea.addEventListener('mousedown', function(e) {
      e.stopPropagation();
    });
    
    // Close on outside click
    setTimeout(() => {
      document.addEventListener('mousedown', handleNotePopoverOutsideClick);
    }, 100);
  };
  
  function handleNotePopoverOutsideClick(e) {
    if (activePopover && !activePopover.contains(e.target)) {
      saveNoteFromPopover(true);
    }
  }
  
  // ============================================================
  // Save Field Note
  // ============================================================
  
  function saveNoteFromPopover(andClose) {
    if (!activeNoteFieldId) return;
    
    const textarea = document.getElementById('notePopoverTextarea');
    const statusEl = document.getElementById('notePopoverStatus');
    const hiddenField = document.getElementById(activeNoteFieldId);
    const config = NOTE_FIELDS.find(c => c.fieldId === activeNoteFieldId);
    
    if (!textarea || !hiddenField || !config) {
      if (andClose) closeNotePopover();
      return;
    }
    
    const newValue = textarea.value;
    const oldValue = hiddenField.value;
    
    // Update hidden field
    hiddenField.value = newValue;
    
    // Update icon state
    const iconBtn = document.querySelector(`.note-icon-btn[data-field-id="${activeNoteFieldId}"]`);
    if (iconBtn) {
      updateNoteIconState(iconBtn, newValue);
    }
    
    // Save to server if changed and we have a contact
    if (newValue !== oldValue && state.currentContactRecord) {
      if (statusEl) statusEl.textContent = 'Saving...';
      
      const updateData = {};
      updateData[config.airtableField] = newValue;
      
      google.script.run
        .withSuccessHandler(function(result) {
          if (statusEl && activePopover) {
            statusEl.textContent = 'Saved';
            setTimeout(() => {
              if (statusEl && activePopover) statusEl.textContent = '';
            }, 1500);
          }
        })
        .withFailureHandler(function(err) {
          console.error('Error saving note:', err);
          if (statusEl && activePopover) statusEl.textContent = 'Error saving';
        })
        .updateContact(state.currentContactRecord.id, updateData);
    }
    
    if (andClose) {
      closeNotePopover();
    }
  }
  
  // ============================================================
  // Close Field Note Popover
  // ============================================================
  
  window.closeNotePopover = function() {
    document.removeEventListener('mousedown', handleNotePopoverOutsideClick);
    clearTimeout(noteDebounceTimer);
    
    if (activePopover) {
      activePopover.remove();
      activePopover = null;
    }
    activeNoteFieldId = null;
  };
  
  // ============================================================
  // Connection Note Popover
  // ============================================================
  
  window.openConnectionNotePopover = function(iconBtn, connectionId, currentNote) {
    closeConnectionNotePopover();
    
    const popover = document.createElement('div');
    popover.className = 'note-popover connection-note-popover';
    popover.id = 'activeConnectionNotePopover';
    popover.innerHTML = `
      <div class="note-popover-header">
        <span class="note-popover-title">Connection Note</span>
        <span class="note-popover-close" onclick="closeConnectionNotePopover()">&times;</span>
      </div>
      <textarea class="note-popover-textarea" id="connNotePopoverTextarea">${escapeHtml(currentNote || '')}</textarea>
      <div class="note-popover-footer">
        <span class="note-popover-status" id="connNotePopoverStatus"></span>
      </div>
    `;
    
    const rect = iconBtn.getBoundingClientRect();
    popover.style.position = 'fixed';
    popover.style.top = (rect.bottom + 5) + 'px';
    popover.style.left = Math.min(rect.left, window.innerWidth - 260) + 'px';
    popover.style.zIndex = '10000';
    
    document.body.appendChild(popover);
    
    activeConnectionPopover = popover;
    activeConnectionId = connectionId;
    
    const textarea = document.getElementById('connNotePopoverTextarea');
    textarea.focus();
    textarea.selectionStart = textarea.value.length;
    
    textarea.addEventListener('input', function() {
      clearTimeout(connNoteDebounceTimer);
      document.getElementById('connNotePopoverStatus').textContent = 'Typing...';
      connNoteDebounceTimer = setTimeout(() => saveConnNoteFromPopover(false), 800);
    });
    
    textarea.addEventListener('keydown', function(e) {
      if (e.key === 'Escape') {
        e.preventDefault();
        closeConnectionNotePopover();
      } else if (e.key === 'Tab' || e.key === 'Enter') {
        if (!e.shiftKey) {
          e.preventDefault();
          saveConnNoteFromPopover(true);
        }
      }
    });
    
    textarea.addEventListener('mousedown', function(e) {
      e.stopPropagation();
    });
    
    setTimeout(() => {
      document.addEventListener('mousedown', handleConnNotePopoverOutsideClick);
    }, 100);
  };
  
  function handleConnNotePopoverOutsideClick(e) {
    if (activeConnectionPopover && !activeConnectionPopover.contains(e.target)) {
      saveConnNoteFromPopover(true);
    }
  }
  
  function saveConnNoteFromPopover(andClose) {
    if (!activeConnectionId) {
      if (andClose) closeConnectionNotePopover();
      return;
    }
    
    const textarea = document.getElementById('connNotePopoverTextarea');
    const statusEl = document.getElementById('connNotePopoverStatus');
    
    if (!textarea) {
      if (andClose) closeConnectionNotePopover();
      return;
    }
    
    const newNote = textarea.value;
    
    if (statusEl) statusEl.textContent = 'Saving...';
    
    google.script.run
      .withSuccessHandler(function(result) {
        if (result.success) {
          if (statusEl && activeConnectionPopover) {
            statusEl.textContent = 'Saved';
            setTimeout(() => {
              if (statusEl && activeConnectionPopover) statusEl.textContent = '';
            }, 1500);
          }
          // Reload connections to update icon state
          if (state.currentContactRecord) {
            loadConnections(state.currentContactRecord.id);
          }
        } else {
          if (statusEl) statusEl.textContent = 'Error';
        }
        if (andClose) closeConnectionNotePopover();
      })
      .withFailureHandler(function(err) {
        console.error('Error saving connection note:', err);
        if (statusEl) statusEl.textContent = 'Error';
        if (andClose) closeConnectionNotePopover();
      })
      .updateConnectionNote(activeConnectionId, newNote);
  }
  
  window.closeConnectionNotePopover = function() {
    document.removeEventListener('mousedown', handleConnNotePopoverOutsideClick);
    clearTimeout(connNoteDebounceTimer);
    
    if (activeConnectionPopover) {
      activeConnectionPopover.remove();
      activeConnectionPopover = null;
    }
    activeConnectionId = null;
  };
  
  // ============================================================
  // Update Note Icon State
  // ============================================================
  
  function updateNoteIconState(iconBtn, value) {
    if (value && value.trim().length > 0) {
      iconBtn.classList.add('has-note');
      iconBtn.title = 'View/edit note';
    } else {
      iconBtn.classList.remove('has-note');
      iconBtn.title = 'Add note';
    }
  }
  
  window.updateAllNoteIcons = function() {
    NOTE_FIELDS.forEach(config => {
      const hiddenField = document.getElementById(config.fieldId);
      const iconBtn = document.querySelector(`.note-icon-btn[data-field-id="${config.fieldId}"]`);
      
      if (hiddenField && iconBtn) {
        updateNoteIconState(iconBtn, hiddenField.value);
      }
    });
  };
  
})();
