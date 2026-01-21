// Timeouts - now using IntegrityState but keeping local refs for backward compat
let searchTimeout;
let spouseSearchTimeout;
let linkedSearchTimeout;
let loadingTimer;
// contactStatusFilter MOVED to IntegrityState - use window.IntegrityState.contactStatusFilter

// Initialize status toggle on page load
document.addEventListener('DOMContentLoaded', function() {
  const saved = localStorage.getItem('contactStatusFilter') || 'Active';
  window.IntegrityState.contactStatusFilter = saved;
  updateStatusToggleUI(saved);
});

function setContactStatusFilter(status) {
  window.IntegrityState.contactStatusFilter = status;
  localStorage.setItem('contactStatusFilter', status);
  updateStatusToggleUI(status);
  // Re-trigger search or reload
  const query = document.getElementById('searchInput')?.value?.trim();
  if (query && query.length > 0) {
    handleSearch({ target: { value: query } });
  } else {
    loadContacts();
  }
}

function updateStatusToggleUI(status) {
  document.querySelectorAll('.status-toggle-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.status === status);
  });
}
// State aliases - these reference IntegrityState for module compatibility
const state = window.IntegrityState;

// Legacy local variables - TODO: migrate to IntegrityState
let pollInterval;
let pollAttempts = 0;
let panelHistory = []; 
let contactHistory = [];
// currentContactRecord MOVED to IntegrityState - use state.currentContactRecord
let currentOppRecords = []; 
let currentOppSortDirection = 'desc'; 
let pendingLinkedEdits = {}; 
let currentPanelData = {}; 
let pendingRemovals = {};
let searchHighlightIndex = -1;
let currentSearchRecords = [];

// Getter/setter for backwards compatibility with local code
Object.defineProperty(window, 'currentContactRecord', {
  get: function() { return state.currentContactRecord; },
  set: function(val) { state.currentContactRecord = val; }
}); 

window.onload = function() { 
  loadContacts(); 
  checkUserIdentity(); 
  initKeyboardShortcuts();
  initDarkMode();
  initScreensaver();
  initInlineEditing();
  initAllNoteFields();
  initScrollHeader();
  initSmartDateListener();
  initSmartTimeListener();
  initSmartFieldEnterGuard();
  
  if (window.IntegrityRouter) {
    window.IntegrityRouter.init();
  }
};

// --- KEYBOARD SHORTCUTS ---
function initKeyboardShortcuts() {
  document.addEventListener('keydown', function(e) {
    const activeEl = document.activeElement;
    const isTyping = activeEl.tagName === 'INPUT' || activeEl.tagName === 'TEXTAREA' || activeEl.isContentEditable;
    
    if (e.key === 'Escape') {
      hideSearchDropdown();
      closeOppPanel();
      closeSpouseModal();
      closeNewOppModal();
      closeShortcutsModal();
      closeDeleteConfirmModal();
      closeDeceasedConfirmModal();
      closeAlertModal();
      if (document.getElementById('actionRow').style.display === 'flex') disableEditMode();
      return;
    }
    
    if (isTyping) return;
    
    if (e.key === '?') {
      e.preventDefault();
      showShortcutsHelp();
      return;
    }
    
    if (e.key === '/') {
      e.preventDefault();
      document.getElementById('searchInput').focus();
      showSearchDropdown();
    } else if (e.key === 'n' || e.key === 'N') {
      e.preventDefault();
      resetForm();
    } else if (e.key === 'e' || e.key === 'E') {
      if (currentContactRecord && document.getElementById('editBtn').style.visibility !== 'hidden') {
        e.preventDefault();
        enableEditMode();
      }
    }
  });
}

// --- AUTO-EXPANDING TEXTAREA ---
function autoExpandTextarea(el) {
  el.style.height = 'auto';
  el.style.height = el.scrollHeight + 'px';
}

function showShortcutsHelp() {
  openModal('shortcutsModal');
}

function closeShortcutsModal() {
  closeModal('shortcutsModal');
}

function closeNewOppModal() {
  closeModal('newOppModal');
}

document.addEventListener('click', function(e) {
  const shortcutsModal = document.getElementById('shortcutsModal');
  if (shortcutsModal && e.target === shortcutsModal) closeShortcutsModal();
  const deleteConfirmModal = document.getElementById('deleteConfirmModal');
  if (deleteConfirmModal && e.target === deleteConfirmModal) closeDeleteConfirmModal();
  const alertModal = document.getElementById('alertModal');
  if (alertModal && e.target === alertModal) closeAlertModal();
});

function checkUserIdentity() {
  google.script.run.withSuccessHandler(function(email) {
     const display = email ? email : "Unknown";
     document.getElementById('debugUser').innerText = display;
     document.getElementById('userEmail').innerText = email || "Not signed in";
     if (!email) alert("Warning: The system cannot detect your email address.");
  }).getEffectiveUserEmail();
}

function updateHeaderTitle(isEditing) {
  const fName = document.getElementById('firstName').value || "";
  const mName = document.getElementById('middleName').value || "";
  const lName = document.getElementById('lastName').value || "";
  let fullName = [fName, mName, lName].filter(Boolean).join(" ");
  if (!fullName.trim()) { document.getElementById('formTitle').innerText = "New Contact"; return; }
  document.getElementById('formTitle').innerText = isEditing ? `Editing ${fullName}` : fullName;
}

function toggleProfileView(show) {
  if(show) {
    document.getElementById('emptyState').style.display = 'none';
    document.getElementById('profileContent').style.display = 'flex';
  } else {
    document.getElementById('emptyState').style.display = 'flex';
    document.getElementById('profileContent').style.display = 'none';
    document.getElementById('formTitle').innerText = "Contact";
    document.getElementById('formSubtitle').innerText = '';
    document.getElementById('refreshBtn').style.display = 'none'; 
    document.getElementById('contactMetaBar').classList.remove('visible');
    document.getElementById('duplicateWarningBox').style.display = 'none'; 
  }
}

function selectContact(record) {
  // If this is a partial record (from optimized list), fetch full record first
  if (record._isPartial) {
    toggleProfileView(true);
    hideSearchDropdown();
    // Show loading state
    document.getElementById('formTitle').innerText = 'Loading...';
    document.getElementById('formSubtitle').innerHTML = '';
    document.getElementById('emptyState').style.display = 'none';
    document.getElementById('profileContent').style.display = 'flex';
    
    google.script.run.withSuccessHandler(function(fullRecord) {
      if (fullRecord && fullRecord.fields) {
        selectContact(fullRecord); // Call again with full record
      } else {
        showAlert('Failed to load contact details');
      }
    }).withFailureHandler(function(err) {
      showAlert('Error loading contact: ' + err.message);
    }).getContactById(record.id);
    return;
  }
  
  document.getElementById('cancelBtn').style.display = 'none';
  toggleProfileView(true);
  hideSearchDropdown(); // Close search dropdown when contact selected
  currentContactRecord = record; 
  const f = record.fields;

  document.getElementById('recordId').value = record.id;
  document.getElementById('firstName').value = f.FirstName || "";
  document.getElementById('middleName').value = f.MiddleName || "";
  document.getElementById('lastName').value = f.LastName || "";
  document.getElementById('preferredName').value = f.PreferredName || "";
  document.getElementById('doesNotLike').value = f["Does Not Like Being Called"] || "";
  document.getElementById('mothersMaidenName').value = f["Mother's Maiden Name"] || "";
  document.getElementById('previousNames').value = f["Previous Names"] || "";
  document.getElementById('mobilePhone').value = f.Mobile || "";
  
  // Email fields
  document.getElementById('email1').value = f.EmailAddress1 || "";
  document.getElementById('email2').value = f.EmailAddress2 || "";
  document.getElementById('email3').value = f.EmailAddress3 || "";
  
  // Note fields (uses NOTE_FIELDS config for automatic handling)
  populateNoteFields(f);
  
  // Notes field (renamed from Description in Airtable)
  document.getElementById('notes').value = f.Notes || "";
  
  // Gender field
  document.getElementById('gender').value = f.Gender || "";
  // Ensure Gender - Other note value is set (populateNoteFields should handle this, but be explicit)
  document.getElementById('genderOther').value = f["Gender - Other"] || "";
  document.getElementById('maritalStatus').value = f["Marital Status"] || "";
  updateAllNoteIcons();
  handleGenderChange(); // Show/hide note icon based on gender value
  
  // Date of Birth - convert from ISO to DD/MM/YYYY display format
  const dob = f["Date of Birth"] || "";
  if (dob) {
    const dobDate = new Date(dob);
    if (!isNaN(dobDate.getTime())) {
      const day = String(dobDate.getDate()).padStart(2, '0');
      const month = String(dobDate.getMonth() + 1).padStart(2, '0');
      const year = dobDate.getFullYear();
      document.getElementById('dateOfBirth').value = `${day}/${month}/${year}`;
    } else {
      document.getElementById('dateOfBirth').value = dob;
    }
  } else {
    document.getElementById('dateOfBirth').value = "";
  }

  disableEditMode(); 
  document.getElementById('editBtn').style.visibility = 'visible';
  document.getElementById('refreshBtn').style.display = 'inline';

  const warnBox = document.getElementById('duplicateWarningBox');
  if (f['Duplicate Warning']) {
     document.getElementById('duplicateWarningText').innerText = f['Duplicate Warning'];
     warnBox.style.display = 'flex'; 
  } else {
     warnBox.style.display = 'none';
  }

  document.getElementById('formSubtitle').innerHTML = formatSubtitle(f);
  
  // Set title with tenure info
  const fullName = formatName(f);
  const tenureText = formatTenureText(f);
  if (tenureText) {
    document.getElementById('formTitle').innerHTML = `${fullName} <span class="tenure-text">${tenureText}</span>`;
  } else {
    document.getElementById('formTitle').innerText = fullName;
  }
  
  renderHistory(f);
  loadOpportunities(f);
  renderSpouseSection(f);
  loadConnections(record.id);
  loadAddressHistory(record.id);
  closeOppPanel();
  
  // Show Actions menu for existing contacts
  const actionsMenu = document.getElementById('actionsMenuWrapper');
  if (actionsMenu) actionsMenu.style.display = 'inline-block';
  
  // Apply deceased styling if contact is deceased
  const isDeceased = f.Deceased === true;
  applyDeceasedStyling(isDeceased);
  
  // Update URL (only if not already navigating from router)
  if (window.IntegrityRouter && !window._routerNavigating) {
    window.IntegrityRouter.navigateTo(record.id, null);
  }
}

// Callback-based version for router (daisy chain pattern)
window.selectContactFromRouter = function(record, callback) {
  window._routerNavigating = true;
  selectContact(record);
  window._routerNavigating = false;
  if (callback) callback();
};

// Collapsible section pattern - reusable for any collapsible field groups
window.toggleCollapsible = function(sectionId) {
  const section = document.getElementById(sectionId);
  if (section) {
    section.classList.toggle('expanded');
  }
};

// Gender field handling (Gender - Other is now a note popover)
function handleGenderChange() {
  // Only show the Gender - Other note icon when "Other (specify)" is selected
  const genderSelect = document.getElementById('gender');
  const genderWrapper = genderSelect?.closest('.input-with-note');
  const noteIcon = genderWrapper?.querySelector('.note-icon');
  
  if (noteIcon) {
    const isOther = genderSelect.value === 'Other (specify)';
    if (isOther) {
      // Delay showing icon until dropdown has closed
      setTimeout(() => {
        noteIcon.style.display = '';
      }, 150);
    } else {
      noteIcon.style.display = 'none';
    }
  }
}

// Unsubscribe handling
function updateUnsubscribeDisplay(isUnsubscribed) {
  const statusEl = document.getElementById('unsubscribeStatus');
  if (!statusEl) return;
  if (isUnsubscribed) {
    statusEl.textContent = 'Unsubscribed';
    statusEl.className = 'unsubscribe-status-unsubscribed';
  } else {
    statusEl.textContent = 'Subscribed';
    statusEl.className = 'unsubscribe-status-subscribed';
  }
}

function openUnsubscribeEdit() {
  if (!currentContactRecord) return;
  const currentValue = currentContactRecord.fields["Unsubscribed from Marketing"] || false;
  document.getElementById('unsubscribeCheckbox').checked = currentValue;
  openModal('unsubscribeModal');
}

function closeUnsubscribeModal() {
  closeModal('unsubscribeModal');
}

function saveUnsubscribePreference() {
  if (!currentContactRecord) return;
  
  const newValue = document.getElementById('unsubscribeCheckbox').checked;
  const currentValue = currentContactRecord.fields["Unsubscribed from Marketing"] || false;
  
  if (newValue === currentValue) {
    closeUnsubscribeModal();
    return;
  }
  
  // Different warnings for subscribe vs unsubscribe
  let confirmMsg;
  if (newValue) {
    // Unsubscribing someone who is subscribed
    confirmMsg = "If you proceed, this person won't receive any marketing communications from us again. Are you sure you want that?";
  } else {
    // Subscribing someone who is unsubscribed - SPAM Act warning
    confirmMsg = "Are you sure you want to change this client's marketing preferences? If they expressed an opinion and you're going against that, it's a breach of the SPAM Act among other things.";
  }
  if (!confirm(confirmMsg)) {
    return;
  }
  
  closeUnsubscribeModal();
  
  google.script.run
    .withSuccessHandler(function(result) {
      if (result && result.fields) {
        currentContactRecord = result;
        updateUnsubscribeDisplay(result.fields["Unsubscribed from Marketing"] || false);
        showAlert('Success', 'Marketing preference updated.', 'success');
      }
    })
    .withFailureHandler(function(err) {
      showAlert('Error', err.message, 'error');
    })
    .updateContact(currentContactRecord.id, "Unsubscribed from Marketing", newValue);
}

function refreshCurrentContact() {
   if (!currentContactRecord) return;
   const btn = document.getElementById('refreshBtn');
   btn.classList.add('spin-anim'); 
   setTimeout(() => { btn.classList.remove('spin-anim'); }, 1000); 

   const id = currentContactRecord.id;
   google.script.run.withSuccessHandler(function(r) {
      if (r && r.fields) selectContact(r);
   }).getContactById(id);
}

function confirmDeleteContact() {
  if (!currentContactRecord) return;
  const f = currentContactRecord.fields;
  const name = formatName(f);
  
  document.getElementById('deleteConfirmMessage').innerText = `Are you sure you want to delete "${name}"? This action cannot be undone.`;
  openModal('deleteConfirmModal');
}

function closeDeleteConfirmModal() {
  closeModal('deleteConfirmModal');
}

function executeDeleteContact() {
  if (!currentContactRecord) return;
  const f = currentContactRecord.fields;
  const name = formatName(f);
  const contactId = currentContactRecord.id;
  
  closeModal('deleteConfirmModal', function() {
    google.script.run
      .withSuccessHandler(function(result) {
        if (result.success) {
          showAlert('Success', `"${name}" has been deleted.`, 'success');
          currentContactRecord = null;
          toggleProfileView(false);
          loadContacts();
        } else {
          showAlert('Cannot Delete', result.error || 'Failed to delete contact.', 'error');
        }
      })
      .withFailureHandler(function(err) {
        showAlert('Error', err.message, 'error');
      })
      .deleteContact(contactId);
  });
}

function openModal(modalId) {
  const modal = document.getElementById(modalId);
  modal.classList.add('visible');
  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      modal.classList.add('showing');
    });
  });
}

function closeModal(modalId, callback) {
  const modal = document.getElementById(modalId);
  modal.classList.remove('showing');
  setTimeout(() => {
    modal.classList.remove('visible');
    if (callback) callback();
  }, 250);
}

function showAlert(title, message, type) {
  const modal = document.getElementById('alertModal');
  const sidebar = document.getElementById('alertModalSidebar');
  const icon = document.getElementById('alertModalIcon');
  
  document.getElementById('alertModalTitle').innerText = title;
  document.getElementById('alertModalMessage').innerText = message;
  
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
  
  openModal('alertModal');
}

function closeAlertModal() {
  closeModal('alertModal');
}

// --- ACTIONS MENU & DECEASED WORKFLOW ---

window.toggleActionsMenu = function() {
  const dropdown = document.getElementById('actionsDropdown');
  if (dropdown) dropdown.classList.toggle('open');
};

document.addEventListener('click', function(e) {
  const wrapper = document.getElementById('actionsMenuWrapper');
  if (wrapper && !wrapper.contains(e.target)) {
    const dropdown = document.getElementById('actionsDropdown');
    if (dropdown) dropdown.classList.remove('open');
  }
});

let pendingDeceasedAction = null; // 'mark' or 'undo'

window.markAsDeceased = function() {
  if (!currentContactRecord) return;
  const f = currentContactRecord.fields;
  const name = formatName(f);
  
  document.getElementById('actionsDropdown').classList.remove('open');
  pendingDeceasedAction = 'mark';
  
  document.getElementById('deceasedModalTitle').innerText = 'Mark as Deceased';
  document.getElementById('deceasedConfirmMessage').innerHTML = 
    `Are you sure you want to mark <strong>${name}</strong> as deceased?<br><br>` +
    `This will:<br>` +
    `&bull; Set the Deceased flag<br>` +
    `&bull; Unsubscribe them from marketing communications<br><br>` +
    `<em>This action can be undone later.</em>`;
  document.getElementById('deceasedConfirmBtn').innerText = 'Mark as Deceased';
  document.getElementById('deceasedConfirmBtn').className = 'btn-primary';
  
  openModal('deceasedConfirmModal');
};

window.undoDeceased = function() {
  if (!currentContactRecord) return;
  const f = currentContactRecord.fields;
  const name = formatName(f);
  
  document.getElementById('actionsDropdown').classList.remove('open');
  pendingDeceasedAction = 'undo';
  
  document.getElementById('deceasedModalTitle').innerText = 'Undo Deceased Status';
  document.getElementById('deceasedConfirmMessage').innerHTML = 
    `Are you sure you want to undo the deceased status for <strong>${name}</strong>?<br><br>` +
    `<strong style="color:#A00;">Important:</strong> This will NOT re-subscribe them to marketing communications.<br><br>` +
    `If you need to send marketing to this contact, you will need to manually change their marketing preferences afterward.`;
  document.getElementById('deceasedConfirmBtn').innerText = 'Undo Status';
  document.getElementById('deceasedConfirmBtn').className = 'btn-primary';
  
  openModal('deceasedConfirmModal');
};

window.closeDeceasedConfirmModal = function() {
  closeModal('deceasedConfirmModal');
  pendingDeceasedAction = null;
};

window.executeMarkDeceased = function() {
  if (!currentContactRecord || !pendingDeceasedAction) return;
  const f = currentContactRecord.fields;
  const name = formatName(f);
  const isMarking = pendingDeceasedAction === 'mark';
  
  closeModal('deceasedConfirmModal', function() {
    google.script.run
      .withSuccessHandler(function(result) {
        if (result.success && result.record) {
          currentContactRecord = result.record;
          applyDeceasedStyling(isMarking);
          if (isMarking) {
            showAlert('Contact Marked as Deceased', `${name} has been marked as deceased and unsubscribed from marketing.`, 'success');
          } else {
            showAlert('Status Updated', `${name} is no longer marked as deceased.`, 'success');
          }
        }
      })
      .withFailureHandler(function(err) {
        showAlert('Error', err.message, 'error');
      })
      .markContactDeceased(currentContactRecord.id, isMarking);
  });
  
  pendingDeceasedAction = null;
};

function applyDeceasedStyling(isDeceased) {
  const badge = document.getElementById('deceasedBadge');
  const profileContent = document.getElementById('profileContent');
  const undoBtn = document.getElementById('undoDeceasedBtn');
  const markBtn = document.querySelector('#actionsDropdown button:first-child');
  
  if (isDeceased) {
    if (badge) badge.style.display = 'inline-block';
    if (profileContent) profileContent.classList.add('contact-deceased');
    if (undoBtn) undoBtn.style.display = 'block';
    if (markBtn) markBtn.style.display = 'none';
  } else {
    if (badge) badge.style.display = 'none';
    if (profileContent) profileContent.classList.remove('contact-deceased');
    if (undoBtn) undoBtn.style.display = 'none';
    if (markBtn) markBtn.style.display = 'block';
  }
  
  // Update unsubscribe display if needed
  if (currentContactRecord && currentContactRecord.fields) {
    updateUnsubscribeDisplay(currentContactRecord.fields["Unsubscribed from Marketing"] || false);
  }
}

// --- CHECKBOX FIELD SAVING (Opportunity Panel) ---

function saveCheckboxField(table, id, fieldKey, isChecked) {
  const label = document.querySelector(`label[for="input_${fieldKey}"]`);
  // For "Converted to Appt", always use that as the label (not Yes/No)
  const labelText = fieldKey === 'Taco: Converted to Appt' ? 'Converted to Appt' : (label ? label.innerText : '');
  if (label) label.innerText = 'Saving...';
  google.script.run.withSuccessHandler(function(res) {
    // Restore the proper label text (not Yes/No)
    if (label) label.innerText = labelText;
    // Toggle appointment fields visibility
    refreshPanelAudit(table, id);
    if (fieldKey === 'Taco: Converted to Appt') {
      const section = document.getElementById('tacoApptFieldsSection');
      if (section) section.style.display = isChecked ? '' : 'none';
      
      // If checked and this is an Opportunities record, create an appointment in the Appointments table
      if (isChecked && table === 'Opportunities') {
        // Get raw primitive values from currentPanelData 
        const getRawValue = (key) => {
          const item = currentPanelData[key];
          // Handle both primitive values and objects with .value property
          if (item === undefined || item === null) return null;
          if (typeof item === 'object' && item.value !== undefined) return item.value;
          return item;
        };
        
        // Get boolean values - check currentPanelData first, then DOM
        const getBoolValue = (key) => {
          const item = currentPanelData[key];
          if (item !== undefined && item !== null) {
            if (typeof item === 'object' && item.value !== undefined) return item.value === true;
            return item === true;
          }
          // Fallback to DOM checkbox
          const input = document.getElementById('input_' + key);
          return input ? input.checked : false;
        };
        
        // Build appointment fields with Airtable field names
        const apptFields = {
          "Appointment Time": getRawValue('Taco: Appointment Time'),
          "Type of Appointment": getRawValue('Taco: Type of Appointment'),
          "How Booked": getRawValue('Taco: How appt booked'),
          "How Booked Other": getRawValue('Taco: How Appt Booked Other'),
          "Phone Number": getRawValue('Taco: Appt Phone Number'),
          "Video Meet URL": getRawValue('Taco: Appt Meet URL'),
          "Need Evidence in Advance": getBoolValue('Taco: Need Evidence in Advance'),
          "Need Appt Reminder": getBoolValue('Taco: Need Appt Reminder'),
          "Notes": getRawValue('Taco: Appt Notes')
        };
        
        console.log('Creating appointment from Converted to Appt toggle with fields:', apptFields);
        
        // Create appointment record in Appointments table
        google.script.run
          .withSuccessHandler(function() {
            console.log('Appointment record created from Converted to Appt toggle');
            loadAppointmentsForOpportunity(id);
          })
          .withFailureHandler(function(err) {
            console.error('Failed to create appointment record:', err);
          })
          .createAppointment(id, apptFields);
      }
    }
  }).withFailureHandler(function(err) {
    const input = document.getElementById('input_' + fieldKey);
    if (input) input.checked = !isChecked;
    if (label) label.innerText = labelText;
    console.error('Failed to save checkbox:', err);
  }).updateRecord(table, id, fieldKey, isChecked);
}

function togglePastApptFields() {
  const section = document.getElementById('apptFieldsSection');
  const notice = document.getElementById('apptCollapsedNotice');
  const noticeText = document.getElementById('apptNoticeText');
  if (!section || !notice) return;
  const isHidden = section.style.display === 'none';
  section.style.display = isHidden ? '' : 'none';
  notice.classList.toggle('expanded', isHidden);
  if (noticeText) {
    noticeText.innerText = isHidden ? 'Hide appointment details' : noticeText.dataset.collapsedText || 'Show appointment details';
  }
}

// --- LINKED RECORD EDITOR (TAGS) ---
function toggleLinkedEdit(key) {
   document.getElementById('view_' + key).style.display = 'none';
   document.getElementById('edit_' + key).style.display = 'block';

   const currentLinks = currentPanelData[key] || [];
   pendingLinkedEdits[key] = currentLinks.map(link => ({...link}));
   pendingRemovals = {}; 
   renderLinkedEditorState(key);
   document.getElementById('error_' + key).innerText = ''; 
}

function cancelLinkedEdit(key) {
   // Just close the editor - pendingLinkedEdits will be repopulated from currentPanelData on next open
   document.getElementById('view_' + key).style.display = 'block';
   document.getElementById('edit_' + key).style.display = 'none';
}

function closeLinkedEdit(key) {
   document.getElementById('view_' + key).style.display = 'block';
   document.getElementById('edit_' + key).style.display = 'none';
}


function renderLinkedEditorState(key) {
   const container = document.getElementById('chip_container_' + key);
   container.innerHTML = '';
   const links = pendingLinkedEdits[key];

   if(links.length === 0) {
      container.innerHTML = '<span style="font-size:11px; color:#999; font-style:italic;">No links selected</span>';
   } else {
      links.forEach(link => {
         const chip = document.createElement('div');
         chip.className = 'link-chip';
         chip.innerHTML = `<span>${link.name}</span><span class="link-chip-remove" onclick="removePendingLink('${key}', '${link.id}')">✕</span>`;
         container.appendChild(chip);
      });
   }
}

function removePendingLink(key, id) {
   pendingLinkedEdits[key] = pendingLinkedEdits[key].filter(l => l.id !== id);
   renderLinkedEditorState(key);
   document.getElementById('error_' + key).innerText = ''; 
}

function handleLinkedSearch(event, key) {
   const query = event.target.value;
   const resultsDiv = document.getElementById('results_' + key);

   document.getElementById('error_' + key).innerText = '';

   clearTimeout(linkedSearchTimeout);
   if(query.length < 2) { resultsDiv.style.display = 'none'; return; }

   resultsDiv.style.display = 'block';
   resultsDiv.innerHTML = '<div style="padding:6px; color:#999; font-style:italic;">Searching...</div>';

   linkedSearchTimeout = setTimeout(() => {
      google.script.run.withSuccessHandler(function(records) {
         resultsDiv.innerHTML = '';
         if(records.length === 0) {
            resultsDiv.innerHTML = '<div style="padding:6px; color:#999; font-style:italic;">No results</div>';
         } else {
            records.forEach(r => {
               const name = formatName(r.fields);
               const details = formatDetails(r.fields);
               const div = document.createElement('div');
               div.className = 'link-result-item';
               div.innerHTML = `<strong>${name}</strong> <span style="color:#888;">${details}</span>`;
               div.onclick = function() {
                  addPendingLink(key, {id: r.id, name: name});
                  resultsDiv.style.display = 'none';
                  event.target.value = '';
               };
               resultsDiv.appendChild(div);
            });
         }
      }).searchContacts(query);
   }, 400);
}

function addPendingLink(key, newLink) {
   const errorEl = document.getElementById('error_' + key);
   errorEl.innerText = ''; 

   // 1. Enforce Primary Single
   if(key === 'Primary Applicant') {
      pendingLinkedEdits[key] = [newLink]; 
   } else {
      // 2. Check Duplicates in current list
      if(pendingLinkedEdits[key].some(l => l.id === newLink.id)) {
         errorEl.innerText = "Already added.";
         return; 
      }
      // 3. Enforce Mutually Exclusive Logic WITH PROMPT
      const exclusiveKeys = ['Primary Applicant', 'Applicants', 'Guarantors'];
      let conflictFound = false;

      exclusiveKeys.forEach(otherKey => {
         if(otherKey === key) return;
         const otherLinks = currentPanelData[otherKey] || [];
         if(otherLinks.some(l => l.id === newLink.id)) {
            conflictFound = true;
            if(confirm(`${newLink.name} is currently a '${otherKey}'.\n\nDo you want to move them to '${key}'?`)) {
               if(!pendingRemovals[otherKey]) pendingRemovals[otherKey] = [];
               pendingRemovals[otherKey].push(newLink.id);
               pendingLinkedEdits[key].push(newLink);
            } else {
               return; 
            }
         }
      });

      if(conflictFound) {
         renderLinkedEditorState(key);
         return; 
      }

      pendingLinkedEdits[key].push(newLink);
   }
   renderLinkedEditorState(key);
}

function saveLinkedEdit(table, id, key) {
   const btn = document.getElementById('btn_save_' + key);
   const originalText = btn.innerText;
   btn.innerText = "Saving..."; btn.disabled = true;

   const operations = [];

   for (const [otherKey, idsToRemove] of Object.entries(pendingRemovals)) {
       const current = currentPanelData[otherKey] || [];
       const kept = current.filter(l => !idsToRemove.includes(l.id)).map(l => l.id);
       if(idsToRemove.length > 0) {
          operations.push({ field: otherKey, val: kept });
       }
   }

   const finalIds = pendingLinkedEdits[key].map(l => l.id);
   operations.push({ field: key, val: finalIds });

   executeQueue(table, id, operations, function() {
       loadPanelRecord(table, id); 
   });
}

function executeQueue(table, id, ops, callback) {
   if(ops.length === 0) {
      callback();
      return;
   }
   const currentOp = ops.shift(); 
   google.script.run.withSuccessHandler(function() {
      executeQueue(table, id, ops, callback); 
   }).updateRecord(table, id, currentOp.field, currentOp.val);
}

// --- END LINKED EDITOR ---

// ... (Spouse Modal Logic & General Logic remains the same) ...
function openSpouseModal() {
   const f = currentContactRecord.fields;
   const spouseName = (f['Spouse Name'] && f['Spouse Name'].length > 0) ? f['Spouse Name'][0] : null;
   const spouseId = (f['Spouse'] && f['Spouse'].length > 0) ? f['Spouse'][0] : null;
   openModal('spouseModal');
   document.getElementById('connectForm').style.display = 'none';
   document.getElementById('confirmConnectForm').style.display = 'none';
   document.getElementById('disconnectForm').style.display = 'none';
   if (spouseId) {
      document.getElementById('disconnectForm').style.display = 'flex';
      document.getElementById('currentSpouseName').innerText = spouseName;
      document.getElementById('currentSpouseId').value = spouseId;
   } else {
      document.getElementById('connectForm').style.display = 'flex';
      document.getElementById('spouseSearchInput').value = '';
      document.getElementById('spouseSearchResults').innerHTML = '';
      document.getElementById('spouseSearchResults').style.display = 'none';
      loadRecentContactsForModal();
   }
}
function closeSpouseModal() { closeModal('spouseModal'); }
function backToSearch() {
   document.getElementById('confirmConnectForm').style.display = 'none';
   document.getElementById('connectForm').style.display = 'flex';
}
function loadRecentContactsForModal() {
   const resultsDiv = document.getElementById('spouseSearchResults');
   const inputVal = document.getElementById('spouseSearchInput').value;
   if(inputVal.length > 0) return; 
   resultsDiv.style.display = 'block';
   resultsDiv.innerHTML = '<div style="padding:10px; color:#999; font-style:italic;">Loading recent...</div>';
   google.script.run.withSuccessHandler(function(records) {
       resultsDiv.innerHTML = '<div style="padding:5px 10px; font-size:10px; color:#999; text-transform:uppercase; font-weight:700;">Recently Modified</div>';
       if (!records || records.length === 0) { resultsDiv.innerHTML += '<div style="padding:8px; font-style:italic; color:#999;">No recent contacts</div>'; } else {
          records.forEach(r => {
             if(r.id === currentContactRecord.id) return;
             renderSearchResultItem(r, resultsDiv);
          });
       }
   }).getRecentContacts();
}
function handleSpouseSearch(event) {
   const query = event.target.value;
   const resultsDiv = document.getElementById('spouseSearchResults');
   clearTimeout(spouseSearchTimeout);
   if(query.length === 0) { loadRecentContactsForModal(); return; }
   resultsDiv.style.display = 'block';
   resultsDiv.innerHTML = '<div style="padding:10px; color:#999; font-style:italic;">Searching...</div>';
   spouseSearchTimeout = setTimeout(() => {
      google.script.run.withSuccessHandler(function(records) {
         resultsDiv.innerHTML = '';
         if (records.length === 0) { resultsDiv.innerHTML = '<div style="padding:8px; font-style:italic; color:#999;">No results</div>'; } else {
            records.forEach(r => {
               if(r.id === currentContactRecord.id) return;
               renderSearchResultItem(r, resultsDiv);
            });
         }
      }).searchContacts(query);
   }, 500);
}
function renderSearchResultItem(r, container) {
   const name = formatName(r.fields);
   const details = formatDetailsRow(r.fields); 
   const div = document.createElement('div');
   div.className = 'search-option';
   div.innerHTML = `<span style="font-weight:700; display:block;">${name}</span><span style="font-size:11px; color:#666;">${details}</span>`;
   div.onclick = function() {
      document.getElementById('targetSpouseName').innerText = name;
      document.getElementById('targetSpouseId').value = r.id;
      document.getElementById('connectForm').style.display = 'none';
      document.getElementById('confirmConnectForm').style.display = 'flex';
   };
   container.appendChild(div);
}
function executeSpouseChange(action) {
   const myId = currentContactRecord.id;
   let statusStr = ""; let otherId = ""; let expectHasSpouse = false; 
   if (action === 'disconnect') {
      statusStr = "disconnected as spouse from";
      otherId = document.getElementById('currentSpouseId').value;
      expectHasSpouse = false;
   } else {
      statusStr = "connected as spouse to";
      otherId = document.getElementById('targetSpouseId').value;
      expectHasSpouse = true;
   }
   closeSpouseModal();
   const statusEl = document.getElementById('spouseStatusText');
   statusEl.innerHTML = `<span style="color:var(--color-star); font-style:italic; font-weight:700; display:inline-flex; align-items:center;">Updating <span class="pulse-dot"></span><span class="pulse-dot"></span><span class="pulse-dot"></span></span>`;
   document.getElementById('spouseEditLink').style.display = 'none'; 
   google.script.run.withSuccessHandler(function(res) {
      pollAttempts = 0; startPolling(myId, expectHasSpouse);
   }).setSpouseStatus(myId, otherId, statusStr);
}
function startPolling(contactId, expectHasSpouse) {
   if(pollInterval) clearInterval(pollInterval);
   pollInterval = setInterval(() => {
      pollAttempts++;
      if (pollAttempts > 20) { 
         clearInterval(pollInterval);
         const statusEl = document.getElementById('spouseStatusText');
         if(statusEl) { statusEl.innerHTML = `<span style="color:#A00;">Update delayed.</span> <a class="data-link" onclick="forceReload('${contactId}')">Refresh</a>`; }
         return;
      }
      google.script.run.withSuccessHandler(function(r) {
         if(r && r.fields) {
            const currentSpouseId = (r.fields['Spouse'] && r.fields['Spouse'].length > 0) ? r.fields['Spouse'][0] : null;
            const match = expectHasSpouse ? (currentSpouseId !== null) : (currentSpouseId === null);
            if (match) {
               clearInterval(pollInterval);
               if(currentContactRecord && currentContactRecord.id === contactId) { selectContact(r); }
            }
         }
      }).getContactById(contactId); 
   }, 2000); 
}
function forceReload(id) {
   clearInterval(pollInterval);
   google.script.run.withSuccessHandler(function(r) { if(r && r.fields) selectContact(r); }).getContactById(id);
}

window.goHome = function() {
  // Clear current contact and show initial empty state
  currentContactRecord = null;
  currentOppRecords = [];
  currentContactAddresses = [];
  toggleProfileView(false);
  
  // Close opp panel without URL update (we'll update URL separately)
  document.getElementById('oppDetailPanel').classList.remove('open');
  state.panelHistory = [];
  
  // Update URL to home
  if (window.IntegrityRouter) {
    window.IntegrityRouter.navigateToHome();
  }
  
  // Clear search and reload contact list
  document.getElementById('searchInput').value = '';
  loadContacts();
  
  // Reset any modals that might be open
  const modals = document.querySelectorAll('.modal-overlay, .modal');
  modals.forEach(m => {
    m.style.display = 'none';
    m.classList.remove('showing', 'visible');
  });
};

function resetForm() {
  // Clear current contact first (new contact mode)
  currentContactRecord = null;
  currentOppRecords = [];
  
  toggleProfileView(true); document.getElementById('contactForm').reset();
  document.getElementById('recordId').value = "";
  // Explicitly enable new contact mode
  enableNewContactMode();
  
  // Update URL to home (no contact selected)
  if (window.IntegrityRouter) {
    window.IntegrityRouter.navigateToHome();
  }
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
  // Hide actions menu and reset deceased styling for new contacts
  const actionsMenu = document.getElementById('actionsMenuWrapper');
  if (actionsMenu) actionsMenu.style.display = 'none';
  const deceasedBadge = document.getElementById('deceasedBadge');
  if (deceasedBadge) deceasedBadge.style.display = 'none';
  const profileContent = document.getElementById('profileContent');
  if (profileContent) profileContent.classList.remove('contact-deceased');
  closeOppPanel();
}

function formatSubtitle(f) {
  const preferredName = f.PreferredName || '';
  const doesNotLike = f['Does Not Like Being Called'] || '';
  if (!preferredName && !doesNotLike) return '';
  const parts = [];
  if (preferredName) parts.push(`<span class="name-pill pill-prefers">prefers "${preferredName}"</span>`);
  if (doesNotLike) parts.push(`<span class="name-pill pill-doesnt-like">doesn't like "${doesNotLike}"</span>`);
  return parts.join(' ');
}
function formatTenureText(f) {
  const tenure = calculateTenure(f.Created);
  if (!tenure) return '';
  if (tenure === 'just added today') return tenure;
  return `in our database for ${tenure}`;
}
function calculateTenure(createdStr) {
  if (!createdStr) return null;
  const dateMatch = createdStr.match(/(\d{2}):(\d{2})\s+(\d{2})\/(\d{2})\/(\d{4})/);
  if (!dateMatch) return null;
  const day = parseInt(dateMatch[3], 10);
  const month = parseInt(dateMatch[4], 10) - 1;
  const year = parseInt(dateMatch[5], 10);
  const createdDate = new Date(year, month, day);
  const diffMs = new Date() - createdDate;
  const diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));
  if (diffDays > 730) return `${Math.floor(diffDays / 365)}+ years`;
  if (diffDays > 60) return `${Math.floor(diffDays / 30)}+ months`;
  if (diffDays >= 1) return diffDays === 1 ? "1 day" : `${diffDays} days`;
  return "just added today";
}
function formatAuditDate(dateStr) {
  if (!dateStr) return null;
  let date = null;
  const localMatch = dateStr.match(/(\d{2}):(\d{2})\s+(\d{2})\/(\d{2})\/(\d{4})/);
  if (localMatch) {
    const day = parseInt(localMatch[3], 10);
    const month = parseInt(localMatch[4], 10) - 1;
    const year = parseInt(localMatch[5], 10);
    const hours = parseInt(localMatch[1], 10);
    const mins = parseInt(localMatch[2], 10);
    date = new Date(year, month, day, hours, mins);
  } else {
    const isoMatch = dateStr.match(/(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/);
    if (isoMatch) {
      date = new Date(dateStr);
    }
  }
  if (!date || isNaN(date.getTime())) return dateStr;
  const now = new Date();
  const diffMs = now - date;
  const diffMins = Math.floor(diffMs / (1000 * 60));
  const diffHours = Math.floor(diffMs / (1000 * 60 * 60));
  const diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));
  if (diffMins < 1) return 'just now';
  if (diffMins < 60) return `${diffMins} min${diffMins > 1 ? 's' : ''} ago`;
  if (diffHours < 24) return `${diffHours} hr${diffHours > 1 ? 's' : ''} ago`;
  if (diffDays === 1) return 'yesterday';
  if (diffDays < 7) return `${diffDays} days ago`;
  return date.toLocaleDateString('en-AU');
}

// Render contact meta bar with created/modified/marketing info
function renderHistory(f) {
  renderContactMetaBar(f);
}

function renderContactMetaBar(f) {
  const bar = document.getElementById('contactMetaBar');
  const isUnsubscribed = f["Unsubscribed from Marketing"] || false;
  const marketingText = isUnsubscribed ? "Unsubscribed" : "Subscribed";
  const marketingClass = isUnsubscribed ? "marketing-status-unsubscribed" : "marketing-status-subscribed";
  const status = f.Status || "Active";
  const statusClass = status === "Inactive" ? "status-inactive" : "status-active";
  
  // Parse Created/Modified - strip "Created:" or "Modified:" prefix if present
  let createdDisplay = f.Created || '';
  let modifiedDisplay = f.Modified || '';
  if (createdDisplay.toLowerCase().startsWith('created:')) {
    createdDisplay = createdDisplay.substring(8).trim();
  }
  if (modifiedDisplay.toLowerCase().startsWith('modified:')) {
    modifiedDisplay = modifiedDisplay.substring(9).trim();
  }
  
  let html = '';
  // Status badge
  html += `<div class="meta-status ${statusClass}" onclick="toggleContactStatus()" title="Click to toggle status">${status}</div>`;
  if (createdDisplay) {
    html += '<div class="meta-divider"></div>';
    html += `<div class="meta-item"><span class="meta-label">Created:</span> <span class="meta-value">${createdDisplay}</span></div>`;
  }
  if (modifiedDisplay) {
    html += '<div class="meta-divider"></div>';
    html += `<div class="meta-item"><span class="meta-label">Modified:</span> <span class="meta-value">${modifiedDisplay}</span></div>`;
  }
  html += `<div class="meta-marketing" onclick="openUnsubscribeEdit()"><span class="meta-label">Marketing:</span><span class="${marketingClass}">${marketingText}</span></div>`;
  
  bar.innerHTML = html;
  bar.classList.add('visible');
}

function handleFormSubmit(formObject) {
  if (window.event) window.event.preventDefault();
  const btn = document.getElementById('submitBtn'); const status = document.getElementById('status');
  btn.disabled = true; btn.innerText = "Saving...";
  const formData = {
    recordId: formObject.recordId.value,
    firstName: formObject.firstName.value,
    middleName: formObject.middleName.value,
    lastName: formObject.lastName.value,
    preferredName: formObject.preferredName.value,
    doesNotLike: formObject.doesNotLike.value,
    mobilePhone: formObject.mobilePhone.value,
    email1: formObject.email1.value,
    email1Comment: formObject.email1Comment.value,
    email2: formObject.email2.value,
    email2Comment: formObject.email2Comment.value,
    email3: formObject.email3.value,
    email3Comment: formObject.email3Comment.value,
    notes: formObject.notes.value,
    gender: formObject.gender.value,
    genderOther: formObject.genderOther.value,
    dateOfBirth: formObject.dateOfBirth.value
  };
  google.script.run.withSuccessHandler(function(response) {
       loadContacts();
       btn.disabled = false;
       btn.innerText = "Update Contact";
       disableEditMode();
       if (response.type === 'create' && response.record) {
         // New contact created - navigate to view it
         selectContact(response.record);
       } else if (formData.recordId) {
         // For updates, refresh the contact to show updated Modified date
         google.script.run.withSuccessHandler(function(r) {
           if (r && r.fields) selectContact(r);
         }).getContactById(formData.recordId);
       }
    }).withFailureHandler(function(err) { status.innerText = "❌ " + err.message; status.className = "status-error"; btn.disabled = false; btn.innerText = "Try Again"; }).processForm(formData);
}
