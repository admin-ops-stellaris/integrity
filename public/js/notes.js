/**
 * notes.js - Note Popover System Module
 * 
 * Handles note popovers for contact fields and connection notes.
 * Includes NOTE_FIELDS configuration, icon rendering, popover management,
 * auto-save with debounce, and keyboard handling.
 * 
 * Dependencies: shared-state.js (for IntegrityState)
 * 
 * Functions exposed to window:
 * - initNoteIcon, initAllNoteFields, populateNoteFields, getNoteFieldValues
 * - openNotePopover, closeNotePopover, saveNoteFromPopover
 * - openConnectionNotePopover, closeConnectionNotePopover, saveConnNoteFromPopover
 * - updateNoteIconState, updateAllNoteIcons
 * - NOTE_FIELDS, NOTE_FIELD_MAP (config arrays)
 */
(function() {
  'use strict';
  
  const state = window.IntegrityState;
  
  state.activeNotePopover = null;
  state.noteSaveTimeout = null;
  state.activeConnNotePopover = null;
  state.connNoteSaveTimeout = null;

  const NOTE_FIELDS = [
    { fieldId: 'email1Comment', airtableField: 'EmailAddress1Comment', inputId: 'email1' },
    { fieldId: 'email2Comment', airtableField: 'EmailAddress2Comment', inputId: 'email2' },
    { fieldId: 'email3Comment', airtableField: 'EmailAddress3Comment', inputId: 'email3' },
    { fieldId: 'genderOther', airtableField: 'Gender - Other', inputId: 'gender' }
  ];

  const NOTE_FIELD_MAP = NOTE_FIELDS.reduce((map, f) => {
    map[f.fieldId] = f.airtableField;
    return map;
  }, {});

  function initNoteIcon(wrapper, fieldId) {
    if (wrapper.querySelector('.note-icon')) return;
    
    const iconBtn = document.createElement('button');
    iconBtn.type = 'button';
    iconBtn.className = 'note-icon';
    iconBtn.dataset.noteField = fieldId;
    iconBtn.onclick = function(e) {
      e.preventDefault();
      e.stopPropagation();
      openNotePopover(this, fieldId);
    };
    
    wrapper.appendChild(iconBtn);
    
    const hiddenField = document.getElementById(fieldId);
    if (hiddenField) {
      updateNoteIconState(iconBtn, hiddenField.value);
    }
  }

  function initAllNoteFields() {
    NOTE_FIELDS.forEach(config => {
      const input = document.getElementById(config.inputId);
      if (input) {
        const wrapper = input.closest('.input-with-note');
        if (wrapper) {
          initNoteIcon(wrapper, config.fieldId);
        }
      }
    });
  }

  function populateNoteFields(contact) {
    NOTE_FIELDS.forEach(config => {
      const field = document.getElementById(config.fieldId);
      if (field) {
        field.value = contact[config.airtableField] || '';
      }
    });
    updateAllNoteIcons();
  }

  function getNoteFieldValues() {
    const values = {};
    NOTE_FIELDS.forEach(config => {
      const field = document.getElementById(config.fieldId);
      values[config.fieldId] = field ? field.value : '';
    });
    return values;
  }

  function updateNoteIconState(iconBtn, value) {
    if (!iconBtn) return;
    if (value && value.trim()) {
      iconBtn.classList.add('has-note');
    } else {
      iconBtn.classList.remove('has-note');
    }
  }

  function updateAllNoteIcons() {
    document.querySelectorAll('.note-icon').forEach(icon => {
      const fieldId = icon.dataset.noteField;
      if (fieldId) {
        const field = document.getElementById(fieldId);
        if (field) {
          updateNoteIconState(icon, field.value);
        }
      }
    });
  }

  function openNotePopover(iconBtn, fieldId) {
    closeNotePopover();
    
    const hiddenField = document.getElementById(fieldId);
    if (!hiddenField) return;
    
    const currentValue = hiddenField.value || '';
    const rect = iconBtn.getBoundingClientRect();
    
    const popover = document.createElement('div');
    popover.className = 'note-popover';
    popover.id = 'activeNotePopover';
    popover.innerHTML = `
      <div class="note-popover-header">
        <span class="note-popover-title">Note</span>
        <button type="button" class="note-popover-close" onclick="closeNotePopover()">×</button>
      </div>
      <textarea id="notePopoverTextarea" placeholder="Add a note...">${currentValue}</textarea>
      <div class="note-popover-footer">
        <span class="note-popover-status" id="notePopoverStatus"></span>
        <button type="button" class="note-popover-done" id="notePopoverDone">Done</button>
      </div>
    `;
    
    document.body.appendChild(popover);
    
    const popoverRect = popover.getBoundingClientRect();
    let top = rect.bottom + 5;
    let left = rect.right - popoverRect.width;
    
    if (left < 10) left = 10;
    if (top + popoverRect.height > window.innerHeight - 10) {
      top = rect.top - popoverRect.height - 5;
    }
    
    popover.style.position = 'fixed';
    popover.style.top = top + 'px';
    popover.style.left = left + 'px';
    
    state.activeNotePopover = {
      element: popover,
      fieldId: fieldId,
      iconBtn: iconBtn,
      originalValue: currentValue
    };
    
    const textarea = popover.querySelector('textarea');
    textarea.focus();
    textarea.setSelectionRange(textarea.value.length, textarea.value.length);
    
    textarea.addEventListener('input', function() {
      if (state.noteSaveTimeout) clearTimeout(state.noteSaveTimeout);
      const status = document.getElementById('notePopoverStatus');
      status.textContent = '';
      status.className = 'note-popover-status';
      
      state.noteSaveTimeout = setTimeout(() => {
        saveNoteFromPopover();
      }, 800);
    });
    
    textarea.addEventListener('blur', function(e) {
      setTimeout(() => {
        if (state.activeNotePopover && !state.activeNotePopover.element.contains(document.activeElement)) {
          saveNoteFromPopover(true);
        }
      }, 100);
    });
    
    textarea.addEventListener('keydown', function(e) {
      if (e.key === 'Escape') {
        saveNoteFromPopover(true);
      }
    });
    
    const doneBtn = popover.querySelector('#notePopoverDone');
    doneBtn.addEventListener('click', function() {
      saveNoteFromPopover(true);
    });
    doneBtn.addEventListener('keydown', function(e) {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        saveNoteFromPopover(true);
      } else if (e.key === 'Escape') {
        saveNoteFromPopover(true);
      }
    });
  }

  function saveNoteFromPopover(andClose = false) {
    if (!state.activeNotePopover) return;
    
    const textarea = document.getElementById('notePopoverTextarea');
    const status = document.getElementById('notePopoverStatus');
    const hiddenField = document.getElementById(state.activeNotePopover.fieldId);
    
    if (!textarea || !hiddenField) return;
    
    const newValue = textarea.value;
    const recordId = state.currentContactRecord?.id;
    const originalValue = state.activeNotePopover.originalValue;
    const airtableField = NOTE_FIELD_MAP[state.activeNotePopover.fieldId];
    const iconBtn = state.activeNotePopover.iconBtn;
    
    hiddenField.value = newValue;
    updateNoteIconState(iconBtn, newValue);
    
    if (andClose) {
      if (state.noteSaveTimeout) {
        clearTimeout(state.noteSaveTimeout);
        state.noteSaveTimeout = null;
      }
      closeNotePopover();
    }
    
    if (!recordId || newValue === originalValue || !airtableField) {
      return;
    }
    
    if (status && !andClose) {
      status.textContent = 'Saving...';
      status.className = 'note-popover-status saving';
    }
    
    google.script.run
      .withSuccessHandler(function() {
        console.log('Note saved successfully');
      })
      .withFailureHandler(function(err) {
        console.error('Error saving note:', err);
      })
      .updateRecord('Contacts', recordId, airtableField, newValue);
  }

  function closeNotePopover() {
    if (state.noteSaveTimeout) {
      clearTimeout(state.noteSaveTimeout);
      state.noteSaveTimeout = null;
    }
    
    if (state.activeNotePopover) {
      state.activeNotePopover.element.remove();
      state.activeNotePopover = null;
    }
  }

  function openConnectionNotePopover(iconBtn, connectionId, currentNote) {
    closeConnectionNotePopover();
    closeNotePopover();
    
    const rect = iconBtn.getBoundingClientRect();
    
    const popover = document.createElement('div');
    popover.className = 'note-popover conn-note-popover';
    popover.id = 'activeConnNotePopover';
    popover.innerHTML = `
      <div class="note-popover-header">
        <span class="note-popover-title">Connection Note</span>
        <button type="button" class="note-popover-close" onclick="closeConnectionNotePopover()">×</button>
      </div>
      <textarea id="connNotePopoverTextarea" placeholder="Add a note about this connection...">${currentNote || ''}</textarea>
      <div class="note-popover-footer">
        <span class="note-popover-status" id="connNotePopoverStatus"></span>
        <button type="button" class="note-popover-done" id="connNotePopoverDone">Done</button>
      </div>
    `;
    
    document.body.appendChild(popover);
    
    const popoverRect = popover.getBoundingClientRect();
    let top = rect.bottom + 5;
    let left = rect.right - popoverRect.width;
    
    if (left < 10) left = 10;
    if (top + popoverRect.height > window.innerHeight - 10) {
      top = rect.top - popoverRect.height - 5;
    }
    
    popover.style.position = 'fixed';
    popover.style.top = top + 'px';
    popover.style.left = left + 'px';
    popover.style.zIndex = '10000';
    
    state.activeConnNotePopover = {
      element: popover,
      connectionId: connectionId,
      originalValue: currentNote || '',
      iconBtn: iconBtn
    };
    
    const textarea = document.getElementById('connNotePopoverTextarea');
    textarea.focus();
    
    textarea.addEventListener('input', function() {
      if (state.connNoteSaveTimeout) clearTimeout(state.connNoteSaveTimeout);
      state.connNoteSaveTimeout = setTimeout(() => saveConnNoteFromPopover(false), 800);
    });
    
    const doneBtn = document.getElementById('connNotePopoverDone');
    doneBtn.addEventListener('click', function() {
      saveConnNoteFromPopover(true);
    });
    
    textarea.addEventListener('keydown', function(e) {
      if (e.key === 'Tab' && !e.shiftKey) {
        e.preventDefault();
        doneBtn.focus();
      } else if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        saveConnNoteFromPopover(true);
      } else if (e.key === 'Escape') {
        saveConnNoteFromPopover(true);
      }
    });
    
    doneBtn.addEventListener('keydown', function(e) {
      if (e.key === 'Enter' || e.key === ' ' || e.key === 'Escape') {
        e.preventDefault();
        saveConnNoteFromPopover(true);
      }
    });
  }

  function saveConnNoteFromPopover(andClose = false) {
    if (!state.activeConnNotePopover) return;
    
    const textarea = document.getElementById('connNotePopoverTextarea');
    const status = document.getElementById('connNotePopoverStatus');
    
    if (!textarea) return;
    
    const newValue = textarea.value;
    const connectionId = state.activeConnNotePopover.connectionId;
    const originalValue = state.activeConnNotePopover.originalValue;
    const iconBtn = state.activeConnNotePopover.iconBtn;
    
    if (iconBtn) {
      if (newValue && newValue.trim()) {
        iconBtn.classList.add('has-note');
      } else {
        iconBtn.classList.remove('has-note');
      }
    }
    
    const parentEl = iconBtn?.closest('[data-conn-note]');
    if (parentEl) {
      parentEl.setAttribute('data-conn-note', newValue);
    }
    
    if (andClose) {
      if (state.connNoteSaveTimeout) {
        clearTimeout(state.connNoteSaveTimeout);
        state.connNoteSaveTimeout = null;
      }
      closeConnectionNotePopover();
    }
    
    if (newValue === originalValue) {
      return;
    }
    
    if (state.activeConnNotePopover) {
      state.activeConnNotePopover.originalValue = newValue;
    }
    
    if (status && !andClose) {
      status.textContent = 'Saving...';
      status.className = 'note-popover-status saving';
    }
    
    google.script.run
      .withSuccessHandler(function() {
        console.log('Connection note saved successfully');
      })
      .withFailureHandler(function(err) {
        console.error('Error saving connection note:', err);
      })
      .updateConnectionNote(connectionId, newValue);
  }

  function closeConnectionNotePopover() {
    if (state.connNoteSaveTimeout) {
      clearTimeout(state.connNoteSaveTimeout);
      state.connNoteSaveTimeout = null;
    }
    
    if (state.activeConnNotePopover) {
      state.activeConnNotePopover.element.remove();
      state.activeConnNotePopover = null;
    }
  }

  document.addEventListener('click', function(e) {
    if (state.activeNotePopover && !state.activeNotePopover.element.contains(e.target) && !e.target.classList.contains('note-icon')) {
      const selection = window.getSelection();
      if (selection && selection.toString().length > 0) {
        const textarea = document.getElementById('notePopoverTextarea');
        if (textarea && (document.activeElement === textarea || selection.anchorNode?.parentElement?.closest('.note-popover'))) {
          return;
        }
      }
      saveNoteFromPopover(true);
    }
  });

  document.addEventListener('click', function(e) {
    if (state.activeConnNotePopover && !state.activeConnNotePopover.element.contains(e.target) && !e.target.classList.contains('conn-note-icon')) {
      const selection = window.getSelection();
      if (selection && selection.toString().length > 0) {
        const textarea = document.getElementById('connNotePopoverTextarea');
        if (textarea && (document.activeElement === textarea || selection.anchorNode?.parentElement?.closest('.conn-note-popover'))) {
          return;
        }
      }
      saveConnNoteFromPopover(true);
    }
  });

  window.NOTE_FIELDS = NOTE_FIELDS;
  window.NOTE_FIELD_MAP = NOTE_FIELD_MAP;
  window.initNoteIcon = initNoteIcon;
  window.initAllNoteFields = initAllNoteFields;
  window.populateNoteFields = populateNoteFields;
  window.getNoteFieldValues = getNoteFieldValues;
  window.updateNoteIconState = updateNoteIconState;
  window.updateAllNoteIcons = updateAllNoteIcons;
  window.openNotePopover = openNotePopover;
  window.saveNoteFromPopover = saveNoteFromPopover;
  window.closeNotePopover = closeNotePopover;
  window.openConnectionNotePopover = openConnectionNotePopover;
  window.saveConnNoteFromPopover = saveConnNoteFromPopover;
  window.closeConnectionNotePopover = closeConnectionNotePopover;

})();
