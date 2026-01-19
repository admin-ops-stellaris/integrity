/**
 * Inline Editing Module
 * Reusable IIFE module for click-to-edit form fields
 * Features: per-field state tracking, Tab/Shift+Tab navigation, select support
 */
(function() {
  'use strict';
  
  const state = window.IntegrityState;
  
  // ============================================================
  // InlineEditingManager - Core inline editing functionality
  // ============================================================
  // Usage: Call InlineEditingManager.init(containerSelector, config) where:
  //   containerSelector: CSS selector for the container (e.g., '#profileTop')
  //   config: { fieldMap: {fieldId: 'AirtableField'}, getRecordId: fn, saveCallback: fn, onFieldSave: fn }
  
  const InlineEditingManager = (function() {
    const instances = new Map();
    
    const normalizers = {
      dateOfBirth: function(value) {
        if (!value) return value;
        if (/^\d{4}-\d{2}-\d{2}/.test(value)) return value;
        const match = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
        if (match) {
          const [, day, month, year] = match;
          return `${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`;
        }
        return value;
      }
    };
    
    function createInstance(container, config) {
      const instanceState = {
        container: container,
        fields: [],
        currentField: null,
        pendingSaves: new Map(),
        fieldMap: config.fieldMap || {},
        getRecordId: config.getRecordId || (() => null),
        saveCallback: config.saveCallback || null,
        onFieldSave: config.onFieldSave || null,
        allowEditing: false,
        sessionCounter: 0
      };
      
      function init() {
        const selectors = 'input:not([type="hidden"]), textarea, select';
        instanceState.fields = Array.from(container.querySelectorAll(selectors))
          .filter(f => !f.id.match(/^(recordId|submitBtn|cancelBtn)$/));
        
        instanceState.fields.forEach((field, index) => {
          field.dataset.inlineIndex = index;
          
          if (field.tagName === 'SELECT') {
            const parent = field.parentElement;
            parent.style.cursor = 'pointer';
            parent.addEventListener('click', function(e) {
              if (!instanceState.allowEditing && field.classList.contains('locked') && canEdit()) {
                e.preventDefault();
                e.stopPropagation();
                enableField(field);
              }
            });
            
            field.addEventListener('change', function() {
              if (instanceState.allowEditing) return;
              if (instanceState.currentField === this) {
                this.dataset.selectSaved = 'true';
                saveField(this, false);
              }
            });
          } else {
            field.addEventListener('click', function(e) {
              if (!instanceState.allowEditing && this.classList.contains('locked') && canEdit()) {
                enableField(this);
                e.stopPropagation();
              }
            });
          }
          
          field.addEventListener('blur', function(e) {
            if (instanceState.allowEditing) return;
            if (instanceState.currentField === this) {
              if (this.dataset.selectSaved === 'true') {
                delete this.dataset.selectSaved;
                return;
              }
              setTimeout(() => {
                if (instanceState.currentField === this) {
                  saveField(this, false);
                }
              }, 10);
            }
          });
          
          field.addEventListener('keydown', function(e) {
            if (instanceState.allowEditing) return;
            if (!instanceState.currentField) return;
            
            if (e.key === 'Tab') {
              e.preventDefault();
              saveField(this, true, e.shiftKey ? -1 : 1);
            } else if (e.key === 'Enter' && this.tagName !== 'TEXTAREA') {
              e.preventDefault();
              saveField(this, false);
            } else if (e.key === 'Escape') {
              cancelEdit(this);
            }
          });
        });
      }
      
      function canEdit() {
        return !!instanceState.getRecordId() || instanceState.allowEditing;
      }
      
      function enableField(field) {
        if (instanceState.currentField && instanceState.currentField !== field) {
          saveFieldSync(instanceState.currentField);
        }
        
        instanceState.currentField = field;
        instanceState.sessionCounter++;
        field.dataset.editSession = instanceState.sessionCounter;
        field.dataset.originalValue = field.value;
        
        field.classList.remove('locked');
        field.classList.add('inline-editing');
        
        if (field.tagName === 'SELECT') {
          field.disabled = false;
        } else {
          field.readOnly = false;
        }
        
        field.focus();
        if (field.tagName === 'INPUT') {
          field.select();
        }
      }
      
      function disableField(field, savedValue) {
        if (instanceState.allowEditing) {
          field.classList.remove('inline-editing', 'saving');
          delete field.dataset.selectSaved;
          field.dataset.originalValue = field.value;
          if (instanceState.currentField === field) {
            instanceState.currentField = null;
          }
          return;
        }
        
        field.classList.add('locked');
        field.classList.remove('inline-editing', 'saving');
        delete field.dataset.selectSaved;
        field.dataset.originalValue = savedValue !== undefined ? savedValue : field.value;
        
        if (field.tagName === 'SELECT') {
          field.disabled = true;
        } else {
          field.readOnly = true;
        }
        
        if (instanceState.currentField === field) {
          instanceState.currentField = null;
        }
      }
      
      function saveFieldSync(field) {
        saveField(field, false);
      }
      
      function saveField(field, moveToNext, direction) {
        const fieldName = field.id || field.name;
        let newValue = normalizeValue(field.value, fieldName);
        const originalVal = field.dataset.originalValue || '';
        
        if (newValue === originalVal) {
          disableField(field);
          if (moveToNext) focusNextField(field, direction);
          return;
        }
        
        field.classList.add('saving');
        
        const saveSessionId = parseInt(field.dataset.editSession) || 0;
        const saveKey = `${field.id || field.name}_${saveSessionId}`;
        instanceState.pendingSaves.set(saveKey, { 
          field: field,
          originalValue: originalVal, 
          fieldName: fieldName, 
          sessionId: saveSessionId 
        });
        
        performSave(field, fieldName, newValue, function(success) {
          const currentSessionId = parseInt(field.dataset.editSession) || 0;
          const isStale = saveSessionId !== currentSessionId;
          
          const saveContext = instanceState.pendingSaves.get(saveKey);
          instanceState.pendingSaves.delete(saveKey);
          
          if (isStale) {
            if (!success) {
              console.warn('Save failed for previous edit session');
            }
            return;
          }
          
          const revertVal = saveContext ? saveContext.originalValue : originalVal;
          
          if (success) {
            disableField(field, newValue);
            field.classList.add('save-success');
            setTimeout(() => field.classList.remove('save-success'), 500);
            if (moveToNext) focusNextField(field, direction);
          } else {
            const userMovedOn = instanceState.currentField !== null && instanceState.currentField !== field;
            
            field.value = revertVal;
            field.classList.remove('saving');
            delete field.dataset.selectSaved;
            
            if (userMovedOn) {
              field.classList.add('locked');
              field.classList.remove('inline-editing');
              field.dataset.originalValue = revertVal;
              if (field.tagName === 'SELECT') {
                field.disabled = true;
              } else {
                field.readOnly = true;
              }
            } else {
              field.classList.add('inline-editing');
              field.classList.remove('locked');
              field.dataset.originalValue = revertVal;
              if (field.tagName === 'SELECT') {
                field.disabled = false;
              } else {
                field.readOnly = false;
              }
              instanceState.currentField = field;
              field.focus();
            }
          }
        }, originalVal);
      }
      
      function performSave(field, fieldName, newValue, callback, originalVal) {
        const airtableField = instanceState.fieldMap[fieldName];
        if (!airtableField) {
          console.warn('No Airtable mapping for field:', fieldName);
          if (callback) callback(true);
          return;
        }
        
        const recordId = instanceState.getRecordId();
        
        if (!recordId || instanceState.allowEditing) {
          if (callback) callback(true);
          return;
        }
        
        const revertValue = originalVal !== undefined ? originalVal : (field.dataset.originalValue || '');
        
        if (instanceState.saveCallback) {
          instanceState.saveCallback(recordId, airtableField, newValue, field, fieldName, function(success) {
            if (success && instanceState.onFieldSave) {
              instanceState.onFieldSave(fieldName, newValue);
            }
            if (callback) callback(success);
          });
        } else {
          google.script.run
            .withSuccessHandler(function() {
              if (instanceState.onFieldSave) {
                instanceState.onFieldSave(fieldName, newValue);
              }
              if (callback) callback(true);
            })
            .withFailureHandler(function(err) {
              console.error('Failed to save:', err);
              if (revertValue !== undefined) {
                field.value = revertValue;
              }
              showAlert('Error', 'Failed to save: ' + err.message, 'error');
              if (callback) callback(false);
            })
            .updateRecord('Contacts', recordId, airtableField, newValue);
        }
      }
      
      function normalizeValue(value, fieldName) {
        if (normalizers[fieldName]) {
          return normalizers[fieldName](value);
        }
        return value;
      }
      
      function focusNextField(currentField, direction) {
        const currentIndex = parseInt(currentField.dataset.inlineIndex);
        let nextIndex = currentIndex + direction;
        
        if (nextIndex >= instanceState.fields.length) nextIndex = 0;
        if (nextIndex < 0) nextIndex = instanceState.fields.length - 1;
        
        let attempts = 0;
        while (attempts < instanceState.fields.length) {
          const nextField = instanceState.fields[nextIndex];
          if (nextField && nextField.offsetParent !== null) {
            enableField(nextField);
            return;
          }
          nextIndex += direction;
          if (nextIndex >= instanceState.fields.length) nextIndex = 0;
          if (nextIndex < 0) nextIndex = instanceState.fields.length - 1;
          attempts++;
        }
        
        currentField.blur();
      }
      
      function cancelEdit(field) {
        if (instanceState.currentField !== field) return;
        const originalVal = field.dataset.originalValue || '';
        if (field.value !== originalVal) {
          field.value = originalVal;
        }
        disableField(field);
      }
      
      function lockAllFields() {
        instanceState.allowEditing = false;
        instanceState.fields.forEach(field => {
          field.classList.add('locked');
          field.classList.remove('inline-editing', 'saving');
          delete field.dataset.originalValue;
          delete field.dataset.selectSaved;
          if (field.tagName === 'SELECT') {
            field.disabled = true;
          } else {
            field.readOnly = true;
          }
        });
        instanceState.currentField = null;
        instanceState.pendingSaves.clear();
      }
      
      function unlockAllFields() {
        instanceState.allowEditing = true;
        instanceState.fields.forEach(field => {
          field.classList.remove('locked');
          delete field.dataset.originalValue;
          if (field.tagName === 'SELECT') {
            field.disabled = false;
          } else {
            field.readOnly = false;
          }
        });
      }
      
      init();
      
      return {
        lockAll: lockAllFields,
        unlockAll: unlockAllFields,
        enableField: enableField,
        getFields: () => instanceState.fields
      };
    }
    
    return {
      init: function(containerSelector, config) {
        const container = document.querySelector(containerSelector);
        if (!container) {
          console.warn('InlineEditingManager: Container not found:', containerSelector);
          return null;
        }
        
        const instance = createInstance(container, config);
        instances.set(containerSelector, instance);
        return instance;
      },
      
      get: function(containerSelector) {
        return instances.get(containerSelector);
      },
      
      addNormalizer: function(fieldName, fn) {
        normalizers[fieldName] = fn;
      }
    };
  })();
  
  // ============================================================
  // Contact Field Mapping
  // ============================================================
  
  const CONTACT_FIELD_MAP = {
    'firstName': 'FirstName',
    'middleName': 'MiddleName', 
    'lastName': 'LastName',
    'preferredName': 'PreferredName',
    'doesNotLike': 'Does Not Like Being Called',
    'mothersMaidenName': "Mother's Maiden Name",
    'previousNames': 'Previous Names',
    'mobilePhone': 'Mobile',
    'dateOfBirth': 'Date of Birth',
    'email1': 'EmailAddress1',
    'email1Comment': 'EmailAddress1Comment',
    'email2': 'EmailAddress2',
    'email2Comment': 'EmailAddress2Comment',
    'email3': 'EmailAddress3',
    'email3Comment': 'EmailAddress3Comment',
    'notes': 'Notes',
    'gender': 'Gender',
    'genderOther': 'Gender - Other',
    'maritalStatus': 'Marital Status'
  };
  
  // ============================================================
  // Contact Inline Editor Initialization
  // ============================================================
  
  function initInlineEditing() {
    state.contactInlineEditor = InlineEditingManager.init('#profileColumns', {
      fieldMap: CONTACT_FIELD_MAP,
      getRecordId: () => document.getElementById('recordId').value,
      onFieldSave: function(fieldName, newValue) {
        if (state.currentContactRecord) {
          const airtableField = CONTACT_FIELD_MAP[fieldName];
          if (airtableField) {
            state.currentContactRecord.fields[airtableField] = newValue;
          }
        }
        
        if (['firstName', 'middleName', 'lastName'].includes(fieldName)) {
          updateHeaderTitle(false);
        }
        
        if (fieldName === 'gender') {
          handleGenderChange();
        }
      }
    });
  }
  
  // ============================================================
  // Edit Mode Functions
  // ============================================================
  
  function enableNewContactMode() {
    if (state.contactInlineEditor) {
      state.contactInlineEditor.unlockAll();
    } else {
      const inputs = document.querySelectorAll('#profileColumns input, #profileColumns textarea');
      inputs.forEach(input => { input.classList.remove('locked'); input.readOnly = false; });
      const selects = document.querySelectorAll('#profileColumns select');
      selects.forEach(select => { select.classList.remove('locked'); select.disabled = false; });
    }
    document.getElementById('actionRow').style.display = 'flex';
    document.getElementById('cancelBtn').style.display = 'inline-block';
    document.getElementById('submitBtn').textContent = 'Create Contact';
    document.getElementById('profileColumns').classList.add('editing');
    updateHeaderTitle(true);
    document.getElementById('firstName').focus();
  }
  
  function disableAllFieldEditing() {
    if (state.contactInlineEditor) {
      state.contactInlineEditor.lockAll();
    } else {
      const inputs = document.querySelectorAll('#profileColumns input, #profileColumns textarea');
      inputs.forEach(input => { input.classList.add('locked'); input.readOnly = true; });
      const selects = document.querySelectorAll('#profileColumns select');
      selects.forEach(select => { select.classList.add('locked'); select.disabled = true; });
    }
    document.getElementById('actionRow').style.display = 'none';
    document.getElementById('cancelBtn').style.display = 'none';
    document.getElementById('profileColumns').classList.remove('editing');
    updateHeaderTitle(false);
  }
  
  function enableEditMode() { enableNewContactMode(); }
  function disableEditMode() { disableAllFieldEditing(); }
  
  function cancelNewContact() {
    document.getElementById('contactForm').reset();
    toggleProfileView(false);
    disableAllFieldEditing();
  }
  
  function cancelEditMode() { cancelNewContact(); }
  
  function handleFormClick(event) {
    const recordId = document.getElementById('recordId').value;
    if (!recordId) {
      enableNewContactMode();
    }
  }
  
  // ============================================================
  // Panel Inline Edit Helpers (for Opportunity panel fields)
  // ============================================================
  
  function refreshPanelAudit(table, id) {
    google.script.run.withSuccessHandler(function(response) {
      if (!response || !response.audit) return;
      const auditSection = document.querySelector('.panel-audit-section');
      if (auditSection && response.audit.Modified) {
        const modifiedDiv = auditSection.querySelectorAll('div')[1];
        if (modifiedDiv) modifiedDiv.innerText = response.audit.Modified;
      }
    }).getRecordDetail(table, id);
  }
  
  function toggleFieldEdit(fieldKey) {
    document.getElementById('view_' + fieldKey).style.display = 'none';
    document.getElementById('edit_' + fieldKey).style.display = 'block';
  }
  
  function cancelFieldEdit(fieldKey) {
    document.getElementById('view_' + fieldKey).style.display = 'block';
    document.getElementById('edit_' + fieldKey).style.display = 'none';
  }
  
  function saveFieldEdit(table, id, fieldKey) {
    const input = document.getElementById('input_' + fieldKey);
    const val = input.value;
    const btn = document.getElementById('btn_save_' + fieldKey);
    const originalText = btn.innerText;
    btn.innerText = "Saving..."; btn.disabled = true;
    google.script.run.withSuccessHandler(function(res) {
      const displayEl = document.getElementById('display_' + fieldKey);
      if (displayEl) displayEl.innerText = val || 'Not set';
      cancelFieldEdit(fieldKey);
      btn.innerText = originalText; btn.disabled = false;
      refreshPanelAudit(table, id);
      if(fieldKey === 'Opportunity Name') {
        document.getElementById('panelTitle').innerText = val;
        if(state.currentContactRecord) { loadOpportunities(state.currentContactRecord.fields); }
      }
      if(fieldKey === 'Status' && val === 'Won') {
        triggerWonCelebration();
      }
      if(fieldKey === 'Status' && state.currentContactRecord) {
        loadOpportunities(state.currentContactRecord.fields);
      }
      if(fieldKey === 'Taco: Type of Appointment') {
        const phoneWrap = document.getElementById('field_wrap_Taco: Appt Phone Number');
        const videoWrap = document.getElementById('field_wrap_Taco: Appt Meet URL');
        if (phoneWrap) phoneWrap.style.display = val === 'Phone' ? '' : 'none';
        if (videoWrap) videoWrap.style.display = val === 'Video' ? '' : 'none';
      }
      if(fieldKey === 'Taco: How appt booked') {
        const otherWrap = document.getElementById('field_wrap_Taco: How Appt Booked Other');
        if (otherWrap) otherWrap.style.display = val === 'Other' ? '' : 'none';
        const reminderLabel = document.querySelector('label[for="input_Taco: Need Appt Reminder"]');
        const reminderInput = document.getElementById('input_Taco: Need Appt Reminder');
        if (reminderLabel && reminderInput && !reminderInput.checked) {
          reminderLabel.innerText = val === 'Calendly' ? 'Not required as Calendly will do it automatically' : 'No';
        }
      }
    }).updateRecord(table, id, fieldKey, val);
  }
  
  function saveDateField(table, id, fieldKey) {
    const input = document.getElementById('input_' + fieldKey);
    const isoVal = input.value;
    let displayVal = '';
    let saveVal = '';
    if (isoVal) {
      const parts = isoVal.split('-');
      if (parts.length === 3) {
        displayVal = `${parts[2]}/${parts[1]}/${parts[0].slice(-2)}`;
        saveVal = `${parts[2]}/${parts[1]}/${parts[0]}`;
      } else {
        displayVal = isoVal;
        saveVal = isoVal;
      }
    }
    const btn = document.getElementById('btn_save_' + fieldKey);
    const originalText = btn.innerText;
    btn.innerText = "Saving..."; btn.disabled = true;
    google.script.run.withSuccessHandler(function(res) {
      document.getElementById('display_' + fieldKey).innerText = displayVal || 'Not set';
      cancelFieldEdit(fieldKey);
      btn.innerText = originalText; btn.disabled = false;
      refreshPanelAudit(table, id);
    }).updateRecord(table, id, fieldKey, saveVal);
  }
  
  // ============================================================
  // Expose functions globally
  // ============================================================
  
  window.InlineEditingManager = InlineEditingManager;
  window.CONTACT_FIELD_MAP = CONTACT_FIELD_MAP;
  window.initInlineEditing = initInlineEditing;
  window.enableNewContactMode = enableNewContactMode;
  window.disableAllFieldEditing = disableAllFieldEditing;
  window.enableEditMode = enableEditMode;
  window.disableEditMode = disableEditMode;
  window.cancelNewContact = cancelNewContact;
  window.cancelEditMode = cancelEditMode;
  window.handleFormClick = handleFormClick;
  window.refreshPanelAudit = refreshPanelAudit;
  window.toggleFieldEdit = toggleFieldEdit;
  window.cancelFieldEdit = cancelFieldEdit;
  window.saveFieldEdit = saveFieldEdit;
  window.saveDateField = saveDateField;
  
})();
