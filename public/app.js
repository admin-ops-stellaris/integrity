  let searchTimeout;
  let spouseSearchTimeout;
  let linkedSearchTimeout;
  let loadingTimer;
  let contactStatusFilter = localStorage.getItem('contactStatusFilter') || 'Active';
  
  // Initialize status toggle on page load
  document.addEventListener('DOMContentLoaded', function() {
    const saved = localStorage.getItem('contactStatusFilter') || 'Active';
    contactStatusFilter = saved;
    updateStatusToggleUI(saved);
  });
  
  function setContactStatusFilter(status) {
    contactStatusFilter = status;
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
  let pollInterval;
  let pollAttempts = 0;
  let panelHistory = []; 
  let contactHistory = [];
  let currentContactRecord = null; 
  let currentOppRecords = []; 
  let currentOppSortDirection = 'desc'; 
  let pendingLinkedEdits = {}; 
  let currentPanelData = {}; 
  let pendingRemovals = {};
  let searchHighlightIndex = -1;
  let currentSearchRecords = []; 

  window.onload = function() { 
    loadContacts(); 
    checkUserIdentity(); 
    initKeyboardShortcuts();
    initDarkMode();
    initScreensaver();
    initInlineEditing();
    initAllNoteFields();
    initScrollHeader();
  };

  // --- SCROLL-HIDE HEADER (Mobile/Tablet) ---
  function initScrollHeader() {
    let lastScrollY = 0;
    let ticking = false;
    const header = document.querySelector('.app-header');
    
    if (!header) return;
    
    function handleScroll(scrollTop) {
      if (!ticking) {
        window.requestAnimationFrame(function() {
          const isMobileOrTablet = window.innerWidth <= 1024;
          
          if (isMobileOrTablet) {
            if (scrollTop > lastScrollY && scrollTop > 50) {
              header.classList.add('header-hidden');
            } else {
              header.classList.remove('header-hidden');
            }
          } else {
            header.classList.remove('header-hidden');
          }
          
          lastScrollY = scrollTop;
          ticking = false;
        });
        ticking = true;
      }
    }
    
    // Listen to all scrollable elements
    const container = document.querySelector('.container');
    const columns = document.querySelectorAll('.column');
    
    if (container) {
      container.addEventListener('scroll', function() {
        handleScroll(this.scrollTop);
      });
    }
    
    columns.forEach(function(col) {
      col.addEventListener('scroll', function() {
        handleScroll(this.scrollTop);
      });
    });
    
    // Reset on resize to desktop
    window.addEventListener('resize', function() {
      if (window.innerWidth > 1024) {
        header.classList.remove('header-hidden');
      }
    });
  }

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

  // --- DARK MODE ---
  function initDarkMode() {
    const savedTheme = localStorage.getItem('integrity-theme');
    if (savedTheme === 'dark') document.body.classList.add('dark-mode');
  }
  function toggleDarkMode() {
    document.body.classList.toggle('dark-mode');
    const isDark = document.body.classList.contains('dark-mode');
    localStorage.setItem('integrity-theme', isDark ? 'dark' : 'light');
  }

  // --- AUTO-EXPANDING TEXTAREA ---
  function autoExpandTextarea(el) {
    el.style.height = 'auto';
    el.style.height = el.scrollHeight + 'px';
  }

  // --- SCREENSAVER ---
  let screensaverTimer = null;
  const SCREENSAVER_DELAY = 120000; // 2 minutes
  
  function initScreensaver() {
    function resetScreensaverTimer() {
      if (document.body.classList.contains('screensaver-active')) {
        document.body.classList.remove('screensaver-active');
      }
      clearTimeout(screensaverTimer);
      screensaverTimer = setTimeout(() => {
        document.body.classList.add('screensaver-active');
      }, SCREENSAVER_DELAY);
    }
    
    ['mousemove', 'mousedown', 'keydown', 'touchstart', 'scroll'].forEach(event => {
      document.addEventListener(event, resetScreensaverTimer, { passive: true });
    });
    
    resetScreensaverTimer();
  }

  // --- QUICK ADD OPPORTUNITY (COMPOSER) ---
  function quickAddOpportunity() {
    if (!currentContactRecord) { alert('Please select a contact first.'); return; }
    openOppComposer();
  }
  
  function openOppComposer() {
    const f = currentContactRecord.fields;
    const contactName = formatName(f);
    const today = new Date().toLocaleDateString('en-AU', { day: '2-digit', month: '2-digit', year: 'numeric' });
    const defaultName = `${contactName} - ${today}`;
    
    document.getElementById('composerOppName').value = defaultName;
    document.getElementById('composerOppType').value = 'Home Loans';
    document.getElementById('composerContactInfo').innerText = `Creating for ${contactName}`;
    document.getElementById('composerPrimaryName').innerText = contactName;
    
    const spouseName = (f['Spouse Name'] && f['Spouse Name'].length > 0) ? f['Spouse Name'][0] : null;
    const spouseId = (f['Spouse'] && f['Spouse'].length > 0) ? f['Spouse'][0] : null;
    
    const spouseSection = document.getElementById('composerSpouseSection');
    if (spouseName && spouseId) {
      document.getElementById('composerSpouseName').innerText = spouseName;
      document.getElementById('composerAddSpouse').checked = false;
      document.getElementById('composerSpouseLabelPrefix').innerText = 'Also add ';
      document.getElementById('composerSpouseLabelSuffix').innerText = ' as Applicant?';
      spouseSection.style.display = 'block';
    } else {
      spouseSection.style.display = 'none';
    }
    
    document.getElementById('oppComposer').classList.add('open');
    setTimeout(() => {
      document.getElementById('composerOppName').focus();
      document.getElementById('composerOppName').select();
    }, 100);
  }
  
  function closeOppComposer() {
    document.getElementById('oppComposer').classList.remove('open');
    clearTacoImport();
  }
  
  let parsedTacoFields = {};
  
  function parseTacoData() {
    const rawText = document.getElementById('tacoRawInput').value;
    if (!rawText.trim()) {
      document.getElementById('tacoPreview').style.display = 'none';
      parsedTacoFields = {};
      return;
    }
    
    google.script.run.withSuccessHandler(function(result) {
      parsedTacoFields = result.parsed || {};
      const display = result.display || [];
      const unmapped = result.unmapped || [];
      
      let html = '';
      if (display.length > 0) {
        display.forEach(item => {
          const displayValue = item.value.length > 50 ? item.value.substring(0, 50) + '...' : item.value;
          html += `<div style="margin-bottom:6px; display:flex; gap:8px;"><span style="color:var(--color-cedar);">&#10003;</span><span style="color:#666;">${item.airtableField}:</span> <span style="color:var(--color-midnight);">${displayValue}</span></div>`;
        });
      }
      if (unmapped.length > 0) {
        html += '<div style="margin-top:10px; padding-top:8px; border-top:1px solid #EEE;"><div style="color:#999; font-size:11px; margin-bottom:6px;">Unrecognized fields:</div>';
        unmapped.forEach(item => {
          const displayValue = item.value.length > 40 ? item.value.substring(0, 40) + '...' : item.value;
          html += `<div style="margin-bottom:4px; color:#999;"><span style="color:#CCC;">?</span> ${item.tacoField}: ${displayValue}</div>`;
        });
        html += '</div>';
      }
      if (display.length === 0 && unmapped.length === 0) {
        html = '<div style="color:#999; font-style:italic;">No valid fields found. Use format: field_name: value</div>';
      }
      
      document.getElementById('tacoPreviewContent').innerHTML = html;
      document.getElementById('tacoPreview').style.display = 'block';
      document.getElementById('tacoImportArea').style.display = 'none';
    }).parseTacoData(rawText);
  }
  
  function clearTacoImport() {
    document.getElementById('tacoRawInput').value = '';
    document.getElementById('tacoPreview').style.display = 'none';
    document.getElementById('tacoImportArea').style.display = 'block';
    parsedTacoFields = {};
  }
  
  function updateComposerSpouseLabel() {
    const checkbox = document.getElementById('composerAddSpouse');
    const prefix = document.getElementById('composerSpouseLabelPrefix');
    const suffix = document.getElementById('composerSpouseLabelSuffix');
    if (checkbox.checked) {
      prefix.innerText = 'Adding ';
      suffix.innerText = ' as Applicant';
    } else {
      prefix.innerText = 'Also add ';
      suffix.innerText = ' as Applicant?';
    }
  }
  
  function submitFromComposer() {
    const oppName = document.getElementById('composerOppName').value.trim();
    if (!oppName) { alert('Please enter an opportunity name.'); return; }
    
    const oppType = document.getElementById('composerOppType').value;
    const f = currentContactRecord.fields;
    const spouseId = (f['Spouse'] && f['Spouse'].length > 0) ? f['Spouse'][0] : null;
    const addSpouse = document.getElementById('composerAddSpouse')?.checked && spouseId;
    
    const tacoFieldsCopy = { ...parsedTacoFields };
    document.getElementById('oppComposer').classList.remove('open');
    
    google.script.run.withSuccessHandler(function(res) {
      clearTacoImport();
      if (res && res.id) {
        const finishUp = () => {
          google.script.run.withSuccessHandler(function(updatedContact) {
            if (updatedContact) {
              currentContactRecord = updatedContact;
              loadOpportunities(updatedContact.fields);
            }
            setTimeout(() => loadPanelRecord('Opportunities', res.id), 300);
          }).getContactById(currentContactRecord.id);
        };
        
        if (addSpouse) {
          google.script.run.withSuccessHandler(finishUp).updateOpportunity(res.id, 'Applicants', [spouseId]);
        } else {
          finishUp();
        }
      } else {
        showAlert('Error', 'Failed to create opportunity. Check that all Taco field names match Airtable exactly.', 'error');
      }
    }).withFailureHandler(function(err) {
      showAlert('Error', err.message || 'Failed to create opportunity', 'error');
    }).createOpportunity(oppName, currentContactRecord.id, oppType, tacoFieldsCopy);
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
  
  // --- CELEBRATION ---
  function triggerWonCelebration() {
    const container = document.createElement('div');
    container.className = 'confetti-container';
    document.body.appendChild(container);
    const colors = ['#BB9934', '#7B8B64', '#19414C', '#D0DFE6', '#F2F0E9'];
    for (let i = 0; i < 50; i++) {
      const confetti = document.createElement('div');
      confetti.className = 'confetti';
      confetti.style.left = Math.random() * 100 + '%';
      confetti.style.backgroundColor = colors[Math.floor(Math.random() * colors.length)];
      confetti.style.animationDelay = Math.random() * 0.5 + 's';
      confetti.style.animationDuration = (Math.random() * 1 + 1.5) + 's';
      container.appendChild(confetti);
    }
    setTimeout(() => container.remove(), 3000);
  }

  // --- AVATAR HELPERS ---
  function getInitials(firstName, lastName) {
    const f = (firstName || '').charAt(0).toUpperCase();
    const l = (lastName || '').charAt(0).toUpperCase();
    return f + l || '?';
  }
  function getAvatarColor(name) {
    const colors = ['#19414C', '#7B8B64', '#BB9934', '#2C2622', '#6B5B4F'];
    let hash = 0;
    for (let i = 0; i < name.length; i++) hash = name.charCodeAt(i) + ((hash << 5) - hash);
    return colors[Math.abs(hash) % colors.length];
  }

  // --- CONTACT QUICK-VIEW CARD ---
  let quickViewContactId = null;
  let quickViewHoverTimeout = null;
  let isQuickViewHovered = false;

  function showContactQuickView(contactId, triggerElement) {
    if (!contactId) return;
    quickViewContactId = contactId;
    
    const card = document.getElementById('contactQuickView');
    const rect = triggerElement.getBoundingClientRect();
    
    // Position card below trigger, centered if possible
    let left = rect.left + (rect.width / 2) - 150;
    let top = rect.bottom + 8;
    
    // Keep within viewport bounds
    if (left < 10) left = 10;
    if (left + 320 > window.innerWidth - 10) left = window.innerWidth - 330;
    if (top + 300 > window.innerHeight) {
      // Show above instead
      top = rect.top - 8;
      card.style.transform = 'translateY(-100%)';
    } else {
      card.style.transform = 'translateY(0)';
    }
    
    card.style.left = left + 'px';
    card.style.top = top + 'px';
    
    // Show loading state
    document.getElementById('quickViewName').textContent = 'Loading...';
    document.getElementById('quickViewPreferred').textContent = '';
    document.getElementById('quickViewAvatar').textContent = '...';
    document.getElementById('quickViewAvatar').style.backgroundColor = '#999';
    document.getElementById('quickViewDetails').innerHTML = '';
    document.getElementById('quickViewFooter').textContent = '';
    
    card.classList.add('visible');
    
    // Fetch contact data
    google.script.run.withSuccessHandler(function(contact) {
      if (!contact || !contact.fields || quickViewContactId !== contactId) return;
      
      const f = contact.fields;
      const fullName = f['Calculated Name'] || 
        `${f.FirstName || ''} ${f.MiddleName || ''} ${f.LastName || ''}`.replace(/\s+/g, ' ').trim();
      const initials = getInitials(f.FirstName, f.LastName);
      const avatarColor = getAvatarColor(fullName);
      
      // Update header
      document.getElementById('quickViewName').textContent = fullName;
      document.getElementById('quickViewAvatar').textContent = initials;
      document.getElementById('quickViewAvatar').style.backgroundColor = avatarColor;
      
      // Show preferred name if different from first name
      const preferred = f.PreferredName;
      if (preferred && preferred.toLowerCase() !== (f.FirstName || '').toLowerCase()) {
        document.getElementById('quickViewPreferred').textContent = `Preferred: ${preferred}`;
      } else {
        document.getElementById('quickViewPreferred').textContent = '';
      }
      
      // Build details
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
      
      // Show last modified
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
    if (isQuickViewHovered) return;
    const card = document.getElementById('contactQuickView');
    card.classList.remove('visible');
    quickViewContactId = null;
  }

  // Load a contact by ID into the main view
  function loadContactById(contactId, addToHistory = false) {
    if (addToHistory && currentContactRecord && currentContactRecord.id !== contactId) {
      contactHistory.push(currentContactRecord);
      updateContactBackButton();
    }
    google.script.run.withSuccessHandler(function(record) {
      if (record && record.fields) {
        selectContact(record);
        updateContactBackButton();
      }
    }).getContactById(contactId);
  }

  function updateContactBackButton() {
    const backBtn = document.getElementById('contactBackBtn');
    if (backBtn) {
      if (contactHistory.length > 0) {
        const prevContact = contactHistory[contactHistory.length - 1];
        const prevName = `${prevContact.fields.FirstName || ''} ${prevContact.fields.LastName || ''}`.trim();
        backBtn.textContent = `← Back to ${prevName}`;
        backBtn.style.display = 'inline-block';
      } else {
        backBtn.style.display = 'none';
      }
    }
  }

  window.goBackToContact = function() {
    if (contactHistory.length > 0) {
      const prevContact = contactHistory.pop();
      selectContact(prevContact);
      updateContactBackButton();
    }
  };

  window.navigateFromQuickView = function() {
    if (!quickViewContactId) return;
    const contactId = quickViewContactId;
    hideContactQuickView();
    // Navigate to full contact view in main panel, saving current contact to history
    loadContactById(contactId, true);
  };

  // Setup hover behavior for quick-view card itself
  document.addEventListener('DOMContentLoaded', function() {
    const card = document.getElementById('contactQuickView');
    if (card) {
      card.addEventListener('mouseenter', function() {
        isQuickViewHovered = true;
        if (quickViewHoverTimeout) clearTimeout(quickViewHoverTimeout);
      });
      card.addEventListener('mouseleave', function() {
        isQuickViewHovered = false;
        quickViewHoverTimeout = setTimeout(hideContactQuickView, 200);
      });
    }
  });

  // Attach quick-view to elements with data-quick-view-contact attribute
  function attachQuickViewToElement(element, contactId) {
    element.setAttribute('data-quick-view-contact', contactId);
    element.addEventListener('mouseenter', function(e) {
      if (quickViewHoverTimeout) clearTimeout(quickViewHoverTimeout);
      quickViewHoverTimeout = setTimeout(function() {
        showContactQuickView(contactId, element);
      }, 300);
    });
    element.addEventListener('mouseleave', function(e) {
      if (quickViewHoverTimeout) clearTimeout(quickViewHoverTimeout);
      quickViewHoverTimeout = setTimeout(hideContactQuickView, 200);
    });
  }

  // Close quick-view on escape key
  document.addEventListener('keydown', function(e) {
    if (e.key === 'Escape') {
      const card = document.getElementById('contactQuickView');
      if (card && card.classList.contains('visible')) {
        isQuickViewHovered = false;
        hideContactQuickView();
      }
    }
  });

  // Close quick-view on click outside
  document.addEventListener('click', function(e) {
    const card = document.getElementById('contactQuickView');
    if (card && card.classList.contains('visible')) {
      if (!card.contains(e.target) && !e.target.hasAttribute('data-quick-view-contact')) {
        isQuickViewHovered = false;
        hideContactQuickView();
      }
    }
  });

  // Event delegation for panel contact links (Primary Applicant, Applicants, Guarantors)
  document.addEventListener('mouseover', function(e) {
    const link = e.target.closest('.panel-contact-link');
    if (link && !link.hasAttribute('data-quick-view-contact')) {
      const contactId = link.getAttribute('data-contact-id');
      if (contactId) {
        attachQuickViewToElement(link, contactId);
        // Trigger the hover immediately
        if (quickViewHoverTimeout) clearTimeout(quickViewHoverTimeout);
        quickViewHoverTimeout = setTimeout(function() {
          showContactQuickView(contactId, link);
        }, 300);
      }
    }
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

  // ============================================================
  // INLINE EDITING MANAGER - Reusable module for click-to-edit fields
  // ============================================================
  // Usage: Call InlineEditingManager.init(containerSelector, config) where:
  //   containerSelector: CSS selector for the container (e.g., '#profileTop')
  //   config: { fieldMap: {fieldId: 'AirtableField'}, getRecordId: fn, saveCallback: fn, onFieldSave: fn }
  // ============================================================
  
  const InlineEditingManager = (function() {
    const instances = new Map();
    
    // Value normalizers for specific field types
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
      const state = {
        container: container,
        fields: [],
        currentField: null,
        pendingSaves: new Map(),  // Per-field save context: field -> {originalValue, fieldName, sessionId}
        fieldMap: config.fieldMap || {},
        getRecordId: config.getRecordId || (() => null),
        saveCallback: config.saveCallback || null,
        onFieldSave: config.onFieldSave || null,
        allowEditing: false,  // For bulk edit mode (new contacts)
        sessionCounter: 0     // Incrementing counter for session IDs
      };
      
      function init() {
        // Find all editable fields
        const selectors = 'input:not([type="hidden"]), textarea, select';
        state.fields = Array.from(container.querySelectorAll(selectors))
          .filter(f => !f.id.match(/^(recordId|submitBtn|cancelBtn)$/));
        
        state.fields.forEach((field, index) => {
          field.dataset.inlineIndex = index;
          
          // For select elements, use a wrapper click since disabled selects don't fire clicks
          if (field.tagName === 'SELECT') {
            const parent = field.parentElement;
            parent.style.cursor = 'pointer';
            parent.addEventListener('click', function(e) {
              // Only handle inline edit clicks when not in bulk mode
              if (!state.allowEditing && field.classList.contains('locked') && canEdit()) {
                e.preventDefault();
                e.stopPropagation();
                enableField(field);
              }
            });
            
            // Handle change event for selects - only for inline edit mode
            field.addEventListener('change', function() {
              if (state.allowEditing) return; // Skip in bulk edit mode
              if (state.currentField === this) {
                this.dataset.selectSaved = 'true';
                saveField(this, false);
              }
            });
          } else {
            // Regular click for inputs/textareas
            field.addEventListener('click', function(e) {
              // Only handle inline edit clicks when not in bulk mode
              if (!state.allowEditing && this.classList.contains('locked') && canEdit()) {
                enableField(this);
                e.stopPropagation();
              }
            });
          }
          
          // Blur to save (only when not in bulk edit mode)
          field.addEventListener('blur', function(e) {
            if (state.allowEditing) return; // Skip in bulk edit mode
            if (state.currentField === this) {
              // Skip if select already saved via change event
              if (this.dataset.selectSaved === 'true') {
                delete this.dataset.selectSaved;
                return;
              }
              // Small delay to allow Tab to be processed first
              setTimeout(() => {
                if (state.currentField === this) {
                  saveField(this, false);
                }
              }, 10);
            }
          });
          
          // Keyboard handling (only active when not in bulk edit mode)
          field.addEventListener('keydown', function(e) {
            if (state.allowEditing) return; // Skip in bulk edit mode - let form handle Tab
            if (!state.currentField) return;
            
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
        // Allow editing if we have a record ID (existing record) or bulk edit mode (new record)
        return !!state.getRecordId() || state.allowEditing;
      }
      
      function enableField(field) {
        // Save any currently editing field first
        if (state.currentField && state.currentField !== field) {
          saveFieldSync(state.currentField);
        }
        
        state.currentField = field;
        // Assign a new session ID for this edit
        state.sessionCounter++;
        field.dataset.editSession = state.sessionCounter;
        // Store original value per-field for async recovery
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
        // In bulk edit mode (new contact), don't lock fields
        if (state.allowEditing) {
          field.classList.remove('inline-editing', 'saving');
          delete field.dataset.selectSaved;
          // Keep originalValue as the current field value for next edit
          field.dataset.originalValue = field.value;
          if (state.currentField === field) {
            state.currentField = null;
          }
          return;
        }
        
        field.classList.add('locked');
        field.classList.remove('inline-editing', 'saving');
        delete field.dataset.selectSaved;
        // Update originalValue to the saved value (or current value) for next edit
        field.dataset.originalValue = savedValue !== undefined ? savedValue : field.value;
        
        if (field.tagName === 'SELECT') {
          field.disabled = true;
        } else {
          field.readOnly = true;
        }
        
        if (state.currentField === field) {
          state.currentField = null;
        }
      }
      
      function saveFieldSync(field) {
        // Called when switching between fields - save previous field
        // Simply delegate to saveField - it handles all the async save logic
        // The "sync" name is historical - it returns immediately but save is async
        saveField(field, false);
      }
      
      function saveField(field, moveToNext, direction) {
        const fieldName = field.id || field.name;
        let newValue = normalizeValue(field.value, fieldName);
        // Get per-field original value (stored when field was enabled)
        const originalVal = field.dataset.originalValue || '';
        
        // No change - just close
        if (newValue === originalVal) {
          disableField(field);
          if (moveToNext) focusNextField(field, direction);
          return;
        }
        
        field.classList.add('saving');
        
        // Track this save with its own session ID (captured at save time, not from current field state)
        const saveSessionId = parseInt(field.dataset.editSession) || 0;
        // Use composite key: field + sessionId to prevent overwrites
        const saveKey = `${field.id || field.name}_${saveSessionId}`;
        state.pendingSaves.set(saveKey, { 
          field: field,
          originalValue: originalVal, 
          fieldName: fieldName, 
          sessionId: saveSessionId 
        });
        
        performSave(field, fieldName, newValue, function(success) {
          // Check if this callback's session is still current
          const currentSessionId = parseInt(field.dataset.editSession) || 0;
          const isStale = saveSessionId !== currentSessionId;
          
          // Get our specific save context using the composite key
          const saveContext = state.pendingSaves.get(saveKey);
          state.pendingSaves.delete(saveKey);  // Clean up our entry
          
          if (isStale) {
            // Stale callback - user has started a new edit session
            // Don't touch field state - the new session owns it now
            if (!success) {
              console.warn('Save failed for previous edit session');
            }
            return;
          }
          
          const revertVal = saveContext ? saveContext.originalValue : originalVal;
          
          if (success) {
            // Success: lock the field with the NEW saved value as the baseline
            disableField(field, newValue);  // Pass the saved value
            field.classList.add('save-success');
            setTimeout(() => field.classList.remove('save-success'), 500);
            if (moveToNext) focusNextField(field, direction);
          } else {
            // Failure: check if user has already moved to another field
            const userMovedOn = state.currentField !== null && state.currentField !== field;
            
            // Revert the field value using per-field context
            field.value = revertVal;
            field.classList.remove('saving');
            delete field.dataset.selectSaved;
            
            if (userMovedOn) {
              // User is editing a different field - just lock this one
              field.classList.add('locked');
              field.classList.remove('inline-editing');
              field.dataset.originalValue = revertVal;  // Keep for next edit
              if (field.tagName === 'SELECT') {
                field.disabled = true;
              } else {
                field.readOnly = true;
              }
              // Error was already shown by performSave, don't steal focus
            } else {
              // User hasn't moved on - keep field editable for retry
              field.classList.add('inline-editing');
              field.classList.remove('locked');
              // Keep originalValue for retry
              field.dataset.originalValue = revertVal;
              if (field.tagName === 'SELECT') {
                field.disabled = false;
              } else {
                field.readOnly = false;
              }
              state.currentField = field;
              field.focus();
            }
            // Don't move to next field on failure
          }
        }, originalVal);
      }
      
      function performSave(field, fieldName, newValue, callback, originalVal) {
        const airtableField = state.fieldMap[fieldName];
        if (!airtableField) {
          console.warn('No Airtable mapping for field:', fieldName);
          if (callback) callback(true);
          return;
        }
        
        const recordId = state.getRecordId();
        
        // In bulk edit mode (new contact), skip server save - form submit handles it
        if (!recordId || state.allowEditing) {
          if (callback) callback(true);
          return;
        }
        
        // Use passed originalVal or fall back to per-field dataset
        const revertValue = originalVal !== undefined ? originalVal : (field.dataset.originalValue || '');
        
        if (state.saveCallback) {
          state.saveCallback(recordId, airtableField, newValue, field, fieldName, function(success) {
            if (success && state.onFieldSave) {
              state.onFieldSave(fieldName, newValue);
            }
            if (callback) callback(success);
          });
        } else {
          // Default save via google.script.run
          google.script.run
            .withSuccessHandler(function() {
              if (state.onFieldSave) {
                state.onFieldSave(fieldName, newValue);
              }
              if (callback) callback(true);
            })
            .withFailureHandler(function(err) {
              console.error('Failed to save:', err);
              if (revertValue !== undefined) {
                field.value = revertValue;  // Revert to original
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
        
        // Wrap around
        if (nextIndex >= state.fields.length) nextIndex = 0;
        if (nextIndex < 0) nextIndex = state.fields.length - 1;
        
        // Skip hidden fields, but limit attempts to prevent infinite loop
        let attempts = 0;
        while (attempts < state.fields.length) {
          const nextField = state.fields[nextIndex];
          // Check if field is visible
          if (nextField && nextField.offsetParent !== null) {
            enableField(nextField);
            return;
          }
          nextIndex += direction;
          if (nextIndex >= state.fields.length) nextIndex = 0;
          if (nextIndex < 0) nextIndex = state.fields.length - 1;
          attempts++;
        }
        
        // No visible field found - just blur and exit
        currentField.blur();
      }
      
      function cancelEdit(field) {
        if (state.currentField !== field) return;
        // Use per-field original value
        const originalVal = field.dataset.originalValue || '';
        if (field.value !== originalVal) {
          field.value = originalVal;
        }
        disableField(field);
      }
      
      function lockAllFields() {
        state.allowEditing = false;
        state.fields.forEach(field => {
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
        state.currentField = null;
        state.pendingSaves.clear();
      }
      
      function unlockAllFields() {
        state.allowEditing = true;
        state.fields.forEach(field => {
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
        getFields: () => state.fields
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

  // Field name to Airtable field mapping for Contacts
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

  // Initialize inline editing for contact form
  let contactInlineEditor = null;
  
  function initInlineEditing() {
    contactInlineEditor = InlineEditingManager.init('#profileColumns', {
      fieldMap: CONTACT_FIELD_MAP,
      getRecordId: () => document.getElementById('recordId').value,
      onFieldSave: function(fieldName, newValue) {
        // Update local record cache
        if (currentContactRecord) {
          const airtableField = CONTACT_FIELD_MAP[fieldName];
          if (airtableField) {
            currentContactRecord.fields[airtableField] = newValue;
          }
        }
        
        // Update header title if name changed
        if (['firstName', 'middleName', 'lastName'].includes(fieldName)) {
          updateHeaderTitle(false);
        }
        
        // Handle gender change
        if (fieldName === 'gender') {
          handleGenderChange();
        }
      }
    });
  }

  // Legacy function - now only used for showing hint on hover
  function handleFormClick(event) {
    // For existing contacts, inline editing handles clicks
    // For new contacts, clicking anywhere enables all fields
    const recordId = document.getElementById('recordId').value;
    if (!recordId) {
      enableNewContactMode();
    }
  }

  function toggleContactStatus() {
    const recordId = document.getElementById('recordId').value;
    if (!recordId || !currentContactRecord) return;
    
    const currentStatus = currentContactRecord.fields.Status || "Active";
    const newStatus = currentStatus === "Active" ? "Inactive" : "Active";
    const contactName = currentContactRecord.fields.PreferredName || currentContactRecord.fields.FirstName || "this contact";
    
    const actionWord = newStatus === "Active" ? "activate" : "deactivate";
    showConfirmModal(`Are you sure you want to ${actionWord} ${contactName}?`, function() {
      google.script.run
        .withSuccessHandler(function() {
          currentContactRecord.fields.Status = newStatus;
          renderContactMetaBar(currentContactRecord.fields);
        })
        .withFailureHandler(function(err) {
          console.error("Failed to update status:", err);
        })
        .updateRecord("Contacts", recordId, "Status", newStatus);
    });
  }

  function showConfirmModal(message, onConfirm) {
    const modal = document.getElementById('confirmModal');
    const msgEl = document.getElementById('confirmModalMessage');
    const okBtn = document.getElementById('confirmModalOk');
    const cancelBtn = document.getElementById('confirmModalCancel');
    
    msgEl.textContent = message;
    modal.style.display = 'flex';
    
    const cleanup = () => {
      modal.style.display = 'none';
      okBtn.onclick = null;
      cancelBtn.onclick = null;
    };
    
    okBtn.onclick = () => { cleanup(); if (onConfirm) onConfirm(); };
    cancelBtn.onclick = cleanup;
  }

  // Enable new contact mode - all fields editable with submit button
  function enableNewContactMode() {
    // Use InlineEditingManager to unlock all fields
    if (contactInlineEditor) {
      contactInlineEditor.unlockAll();
    } else {
      // Fallback if manager not initialized yet
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

  // Disable all field editing (for new contact cancel or after save)
  function disableAllFieldEditing() {
    // Use InlineEditingManager to lock all fields
    if (contactInlineEditor) {
      contactInlineEditor.lockAll();
    } else {
      // Fallback if manager not initialized yet
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

  // Legacy aliases for compatibility
  function enableEditMode() { enableNewContactMode(); }
  function disableEditMode() { disableAllFieldEditing(); }
  
  function cancelNewContact() {
    document.getElementById('contactForm').reset();
    toggleProfileView(false);
    disableAllFieldEditing();
  }
  
  function cancelEditMode() { cancelNewContact(); }

  function selectContact(record) {
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
  }
  
  // Collapsible section pattern - reusable for any collapsible field groups
  window.toggleCollapsible = function(sectionId) {
    const section = document.getElementById(sectionId);
    if (section) {
      section.classList.toggle('expanded');
    }
  };

  // Search dropdown behavior
  window.showSearchDropdown = function() {
    const dropdown = document.getElementById('searchDropdown');
    if (dropdown) dropdown.classList.add('open');
  };
  
  window.hideSearchDropdown = function() {
    const dropdown = document.getElementById('searchDropdown');
    if (dropdown) dropdown.classList.remove('open');
  };
  
  // Close dropdown when clicking outside
  document.addEventListener('click', function(e) {
    const wrapper = document.querySelector('.header-search-wrapper');
    if (wrapper && !wrapper.contains(e.target)) {
      hideSearchDropdown();
    }
  });

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
  
  let currentOppToDelete = null;
  
  function confirmDeleteOpportunity(oppId, oppName) {
    currentOppToDelete = { id: oppId, name: oppName };
    document.getElementById('deleteOppConfirmMessage').innerText = `Are you sure you want to delete "${oppName}"? This action cannot be undone.`;
    openModal('deleteOppConfirmModal');
  }
  
  function closeDeleteOppConfirmModal() {
    closeModal('deleteOppConfirmModal');
    currentOppToDelete = null;
  }
  
  function executeDeleteOpportunity() {
    if (!currentOppToDelete) return;
    const { id, name } = currentOppToDelete;
    
    closeModal('deleteOppConfirmModal', function() {
      google.script.run
        .withSuccessHandler(function(result) {
          if (result.success) {
            showAlert('Success', `"${name}" has been deleted.`, 'success');
            closeOppPanel();
            if (currentContactRecord) {
              google.script.run.withSuccessHandler(function(updatedContact) {
                if (updatedContact) {
                  currentContactRecord = updatedContact;
                  loadOpportunities(updatedContact.fields);
                }
              }).getContactById(currentContactRecord.id);
            }
          } else {
            showAlert('Cannot Delete', result.error || 'Failed to delete opportunity.', 'error');
          }
        })
        .withFailureHandler(function(err) {
          showAlert('Error', err.message, 'error');
        })
        .deleteOpportunity(id);
    });
    currentOppToDelete = null;
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

  // --- EMAIL COMPOSER ---
  let currentEmailContext = null;
  
  // Email template links (editable, saved to Airtable Settings table for team-wide sync)
  const DEFAULT_EMAIL_LINKS = {
    officeMap: 'https://maps.app.goo.gl/qm2ohJP2j1t6GqCt9',
    ourTeam: 'https://stellaris.loans/our-team',
    factFind: 'https://drive.google.com/file/d/1_U6kKck5IA3TBtFdJEzyxs_XpvcKvg9s/view?usp=sharing',
    myGov: 'https://my.gov.au/',
    myGovVideo: 'https://www.youtube.com/watch?v=bSMs2XO1V7Y',
    incomeStatementInstructions: 'https://drive.google.com/file/d/1Y8B4zPLb_DTkV2GZnlGztm-HMfA3OWYP/view?usp=sharing'
  };
  
  // Settings key mapping
  const SETTINGS_KEYS = {
    officeMap: 'email_link_office_map',
    ourTeam: 'email_link_our_team',
    factFind: 'email_link_fact_find',
    myGov: 'email_link_mygov',
    myGovVideo: 'email_link_mygov_video',
    incomeStatementInstructions: 'email_link_income_instructions'
  };
  
  let EMAIL_LINKS = { ...DEFAULT_EMAIL_LINKS };
  let userSignature = '';
  let emailSettingsLoaded = false;
  let emailQuill = null;
  
  function loadEmailLinksFromSettings() {
    google.script.run.withSuccessHandler(function(settings) {
      if (settings) {
        Object.keys(SETTINGS_KEYS).forEach(key => {
          const settingKey = SETTINGS_KEYS[key];
          if (settings[settingKey]) {
            EMAIL_LINKS[key] = settings[settingKey];
          }
        });
        emailSettingsLoaded = true;
      }
    }).getAllSettings();
  }
  
  loadEmailLinksFromSettings();
  
  function openEmailSettings() {
    const officeMap = document.getElementById('settingOfficeMap');
    const ourTeam = document.getElementById('settingOurTeam');
    const factFind = document.getElementById('settingFactFind');
    const myGov = document.getElementById('settingMyGov');
    const myGovVideo = document.getElementById('settingMyGovVideo');
    const incomeInstructions = document.getElementById('settingIncomeInstructions');
    const signature = document.getElementById('settingSignature');
    const previewContainer = document.getElementById('signaturePreviewContainer');
    
    // Populate fields (null-safe)
    if (officeMap) officeMap.value = EMAIL_LINKS.officeMap || '';
    if (ourTeam) ourTeam.value = EMAIL_LINKS.ourTeam || '';
    if (factFind) factFind.value = EMAIL_LINKS.factFind || '';
    if (myGov) myGov.value = EMAIL_LINKS.myGov || '';
    if (myGovVideo) myGovVideo.value = EMAIL_LINKS.myGovVideo || '';
    if (incomeInstructions) incomeInstructions.value = EMAIL_LINKS.incomeStatementInstructions || '';
    if (signature) signature.value = userSignature || '';
    
    // Update signature preview and set generatedSignatureHtml for copy functions
    if (previewContainer) {
      if (userSignature) {
        previewContainer.innerHTML = userSignature;
        generatedSignatureHtml = userSignature; // For copy functions
      } else {
        previewContainer.innerHTML = '<span style="color:#999; font-style:italic;">No signature set. Click "Regenerate" to create one.</span>';
      }
    }
    
    // Load user profile info for signature display
    if (currentUserProfile) {
      updateSignatureUserInfo();
    } else {
      google.script.run.withSuccessHandler(function(result) {
        if (result) {
          currentUserProfile = result;
          updateSignatureUserInfo();
        }
      }).getUserSignature();
    }
    
    openModal('emailSettingsModal');
  }
  
  // Global settings accessible from header cog
  function openGlobalSettings() {
    openEmailSettings();
  }
  
  function closeEmailSettings() {
    closeModal('emailSettingsModal');
  }
  
  function saveEmailSettings() {
    const newLinks = {
      officeMap: document.getElementById('settingOfficeMap').value,
      ourTeam: document.getElementById('settingOurTeam').value,
      factFind: document.getElementById('settingFactFind').value,
      myGov: document.getElementById('settingMyGov').value,
      myGovVideo: document.getElementById('settingMyGovVideo').value,
      incomeStatementInstructions: document.getElementById('settingIncomeInstructions').value
    };
    
    // Update local copy immediately
    Object.assign(EMAIL_LINKS, newLinks);
    
    // Save each changed setting to Airtable
    let saveCount = 0;
    let savedCount = 0;
    Object.keys(SETTINGS_KEYS).forEach(key => {
      const settingKey = SETTINGS_KEYS[key];
      saveCount++;
      google.script.run.withSuccessHandler(function() {
        savedCount++;
        if (savedCount === saveCount) {
          checkSignatureAndClose();
        }
      }).withFailureHandler(function(err) {
        savedCount++;
        console.error('Failed to save setting:', settingKey, err);
        if (savedCount === saveCount) {
          checkSignatureAndClose();
        }
      }).updateSetting(settingKey, newLinks[key]);
    });
    
    function checkSignatureAndClose() {
      const newSignature = document.getElementById('settingSignature').value;
      if (newSignature !== userSignature) {
        google.script.run.withSuccessHandler(function() {
          userSignature = newSignature;
          updateEmailPreview();
          showAlert('Saved', 'Settings and signature updated for the whole team', 'success');
        }).withFailureHandler(function(err) {
          showAlert('Error', 'Settings saved, but signature failed: ' + err, 'error');
        }).updateUserSignature(newSignature);
      } else {
        showAlert('Saved', 'Email template links updated for the whole team', 'success');
      }
      closeEmailSettings();
      updateEmailPreview();
    }
  }
  
  let currentUserProfile = null;
  
  function loadUserSignature() {
    google.script.run.withSuccessHandler(function(result) {
      if (result) {
        currentUserProfile = result;
        if (result.signature) {
          userSignature = result.signature;
        }
      }
    }).getUserSignature();
  }
  
  loadUserSignature();
  
  // --- SIGNATURE GENERATOR ---
  let generatedSignatureHtml = '';
  
  function updateSignatureUserInfo() {
    const nameEl = document.getElementById('sigGenName');
    const titleEl = document.getElementById('sigGenTitle');
    if (nameEl && currentUserProfile) {
      nameEl.innerText = currentUserProfile.name || 'Unknown';
    }
    if (titleEl && currentUserProfile) {
      titleEl.innerText = currentUserProfile.title || '';
    }
  }
  
  function regenerateSignature() {
    // Regenerate signature from current user profile
    if (currentUserProfile) {
      generateSignaturePreview();
      // Update the hidden textarea with the new signature
      const textarea = document.getElementById('settingSignature');
      if (textarea) {
        textarea.value = generatedSignatureHtml;
      }
      showAlert('Regenerated', 'Signature updated. Click Save to store it.', 'success');
    } else {
      google.script.run.withSuccessHandler(function(result) {
        if (result) {
          currentUserProfile = result;
          updateSignatureUserInfo();
          generateSignaturePreview();
          // Update the hidden textarea with the new signature
          const textarea = document.getElementById('settingSignature');
          if (textarea) {
            textarea.value = generatedSignatureHtml;
          }
          showAlert('Regenerated', 'Signature updated. Click Save to store it.', 'success');
        }
      }).getUserSignature();
    }
  }
  
  function generateSignaturePreview() {
    if (!currentUserProfile) return;
    
    const name = currentUserProfile.name || 'Your Name';
    const title = currentUserProfile.title || 'Your Title';
    
    // Exact team-wide signature template - Mercury/Gmail compatible
    const signatureHtml = `<table cellpadding="0" cellspacing="0" border="0" style="font-family: Arial, sans-serif; font-size: 10pt; color: #333333;">
    <tbody>
        <tr>
            <td style="padding-bottom: 10px; line-height: 1.5;">
                Best wishes,<br><br>
                <strong style="font-size: 11pt;">${name}</strong><br>
                <strong>${title}</strong><br>
                <strong>Stellaris Finance Broking</strong>
            </td>
        </tr>
        <tr>
            <td style="line-height: 1.5;">
                Phone: 0488 839 212<br>
                Office: Suite 18 / 56 Creaney Drive, Kingsley WA 6026<br>
                Website: <a href="https://www.stellaris.loans" target="_blank" style="color: #1155cc; text-decoration: none;">www.stellaris.loans</a><br>
                Book an Appointment with Tim: <a href="https://calendly.com/tim-kerin" target="_blank" style="color: #1155cc; text-decoration: none;">calendly.com/tim-kerin</a>
            </td>
        </tr>
        <tr>
            <td style="padding-top: 15px;">
                <img 
                    src="https://img1.wsimg.com/isteam/ip/2c5f94ee-4964-4e9b-9b9c-a55121f8611b/WEB_Stellaris_Email%20Signature_Midnight.png" 
                    alt="Stellaris Finance Broking Email Signature Graphic" 
                    width="320" 
                    height="104" 
                    style="display: block; border: 0; width: 320px; height: 104px;">
            </td>
        </tr>
        <tr>
            <td style="padding-top: 20px; font-size: 9pt; font-style: italic; color: #888888; line-height: 1.4;">
                Credit Representative 379175 is authorised under Australian Credit Licence 389328
                <br><br>
                <strong>Confidentiality Notice</strong><br>
                This email and its attachments are confidential and intended for the recipient only. If you are not the intended recipient, please notify us and delete this message. Unauthorized use, dissemination, or copying is prohibited. The views expressed are those of the sender unless stated otherwise. We do not guarantee that attachments are free from viruses; the user assumes all responsibility for any resulting damage. We value your privacy. Your information may be used to provide financial services and may be shared with third parties as required by law.
                <br><br>
                <strong>Important Notice</strong><br>
                We will never ask you to transfer money or make payments via email. If you receive any such requests, please do not respond and contact us directly before taking any action. Your security is our priority. More information on identifying and protecting yourself from phishing attacks generally can be found on the government's ScamWatch website, www.scamwatch.gov.au.
            </td>
        </tr>
    </tbody>
</table>`;
    
    generatedSignatureHtml = signatureHtml;
    // Update the preview container in Settings modal
    const previewContainer = document.getElementById('signaturePreviewContainer');
    if (previewContainer) {
      previewContainer.innerHTML = signatureHtml;
    }
  }
  
  async function copySignatureForGmail() {
    const previewEl = document.getElementById('signaturePreviewContainer');
    if (!previewEl || !generatedSignatureHtml) {
      showAlert('No Signature', 'Generate a signature first before copying.', 'error');
      return;
    }
    
    const gmailInstructions = `Signature has been copied to your clipboard. To update in Gmail:

1. Settings cog (top right of Gmail) then "See all settings"
2. Scroll down to signature, select "Stellaris Signature"
3. Click in the field to the right, remove everything from it and then Ctrl-V the new signature.
4. Scroll to the bottom and click Save Changes and you're done.`;
    
    try {
      // Copy rich text (HTML blob) for Gmail paste
      if (navigator.clipboard && navigator.clipboard.write) {
        const htmlBlob = new Blob([generatedSignatureHtml], { type: 'text/html' });
        const textBlob = new Blob([previewEl.innerText], { type: 'text/plain' });
        const clipboardItem = new ClipboardItem({
          'text/html': htmlBlob,
          'text/plain': textBlob
        });
        await navigator.clipboard.write([clipboardItem]);
        showAlert('Copied for Gmail', gmailInstructions, 'success');
      } else {
        // Fallback: select and copy the preview element directly
        const range = document.createRange();
        range.selectNodeContents(previewEl);
        const selection = window.getSelection();
        selection.removeAllRanges();
        selection.addRange(range);
        document.execCommand('copy');
        selection.removeAllRanges();
        showAlert('Copied for Gmail', gmailInstructions, 'success');
      }
    } catch (err) {
      // Last resort fallback
      const range = document.createRange();
      range.selectNodeContents(previewEl);
      const selection = window.getSelection();
      selection.removeAllRanges();
      selection.addRange(range);
      document.execCommand('copy');
      selection.removeAllRanges();
      showAlert('Copied for Gmail', gmailInstructions, 'success');
    }
  }
  
  async function copySignatureForMercury() {
    const previewEl = document.getElementById('signaturePreviewContainer');
    if (!previewEl || !generatedSignatureHtml) {
      showAlert('No Signature', 'Generate a signature first before copying.', 'error');
      return;
    }
    
    const mercuryInstructions = `Signature has been copied to your clipboard. To paste in Mercury, log into Mercury then:

1. Admin tile
2. Email Profiles tab
3. On the left, select the Profile you want to update the signature for
4. On the right, Edit Details tab
5. In the Email Signature pane, click the three dots in the top right (More misc) then < > (Code view)
6. Delete all that code
7. Ctrl-V the code you just copied here in Integrity
8. Click < > (Code View) again to turn it off and make sure the rendered sig looks right, then click the Preview tab to be sure. If it's good, you're done.`;
    
    try {
      // Copy HTML code for Mercury (they need to paste into code view)
      if (navigator.clipboard && navigator.clipboard.writeText) {
        await navigator.clipboard.writeText(generatedSignatureHtml);
        showAlert('Copied for Mercury', mercuryInstructions, 'success');
      } else {
        throw new Error('Clipboard API not available');
      }
    } catch (err) {
      // Fallback: create temp textarea with HTML
      const textarea = document.createElement('textarea');
      textarea.value = generatedSignatureHtml;
      textarea.style.position = 'fixed';
      textarea.style.left = '-9999px';
      document.body.appendChild(textarea);
      textarea.select();
      document.execCommand('copy');
      document.body.removeChild(textarea);
      showAlert('Copied for Mercury', mercuryInstructions, 'success');
    }
  }
  
  // Legacy function for compatibility - now inline in Settings modal
  function useGeneratedSignature() {
    const textarea = document.getElementById('settingSignature');
    if (textarea) {
      textarea.value = generatedSignatureHtml;
    }
    
    const previewContainer = document.getElementById('signaturePreviewContainer');
    if (previewContainer) {
      previewContainer.innerHTML = generatedSignatureHtml;
    }
    
    showAlert('Applied!', 'Signature added. Click Save to store it.', 'success');
  }
  
  const EMAIL_TEMPLATE = {
    subject: {
      Office: 'Confirmation of Appointment - {{appointmentTime}}',
      Phone: 'Confirmation of Phone Appointment - {{appointmentTime}}',
      Video: 'Confirmation of Google Meet Appointment - {{appointmentTime}}'
    },
    
    opening: {
      Office: "I'm writing to confirm your appointment with {{brokerIntro}} on {{appointmentTime}} (Perth time) ({{daysUntil}} days from today) at our office - Kingsley Professional Centre, 18 / 56 Creaney Drive, Kingsley ({{officeMapLink}}).",
      Phone: "I'm writing to confirm your phone appointment with {{brokerIntro}} on {{appointmentTime}} (Perth time) ({{daysUntil}} days from today). {{brokerFirst}} will call you on {{phoneNumber}}.",
      Video: "I'm writing to confirm your video call appointment with {{brokerIntro}} on {{appointmentTime}} (Perth time) ({{daysUntil}} days from today) using Google Meet URL: {{meetUrl}}. If you have any trouble logging in, please call or text our team on 0488 839 212."
    },
    
    meetLine: {
      New: "Click {{ourTeamLink}} to meet {{brokerFirst}} and the rest of the Team at Stellaris Finance Broking. We will be supporting you each step of the way!",
      Repeat: "Click {{ourTeamLink}} to get reacquainted with the Team at Stellaris Finance Broking. We will be supporting you each step of the way!"
    },
    
    preparation: {
      Shae: `In preparation for your appointment, please email me the following information:

{{factFindLink}} (please note you cannot access this file directly, you will need to download it to your device and fill it in)
- Complete with as much detail as possible
- Include any Buy Now Pay Later services (like Humm, Zip or Afterpay) that you have accounts with under the Personal Loans section at the bottom of Page 3

Income
a) PAYG Income – Your latest two consecutive payslips and your 2024-25 Income Statement (which can be downloaded from {{myGovLink}}. If you need help creating a myGov account, watch {{myGovVideoLink}}. For instructions on how to download your Income Statement, {{incomeInstructionsLink}})
b) Self Employed Income - From each of the last two financial years, your Tax Return, Financial Statements and Notice of Assessment

I work part time – please try to ensure you email the above evidence well ahead of your appointment to allow ample time to process your information.`,
      Team: `In preparation for your appointment, please email Shae (shae@stellaris.loans) the following information:

{{factFindLink}} (please note you cannot access this file directly, you will need to download it to your device and fill it in)
- Complete with as much detail as possible
- Include any Buy Now Pay Later services (like Humm, Zip or Afterpay) that you have accounts with under the Personal Loans section at the bottom of Page 3

Income
a) PAYG Income – Your latest two consecutive payslips and 2024-25 Income Statement (which can be downloaded from {{myGovLink}}. If you need help creating a myGov account, watch {{myGovVideoLink}}. For instructions on how to download your Income Statement, {{incomeInstructionsLink}})
b) Self Employed Income - From each of the last two financial years, your Tax Return, Financial Statements and Notice of Assessment

Please try to ensure you email the above evidence well ahead of your appointment to allow ample time to process your information.`,
      OpenBanking: `You will soon receive invitations to share your information with us via Frollo's Open Banking and Connective's Client Centre. These two systems streamline the collection of your key financial and personal data, including your contact details, employment history, savings and liabilities, to give us a complete picture of your situation.{{prefillNote}}`
    },
    
    prefillNote: {
      New: '',
      Repeat: ' I have prefilled as much as I can using the information we previously received from you.'
    },
    
    closing: {
      New: `Do not hesitate to contact our team on 0488 839 212 if you have any questions.

We look forward to working with you!

Best wishes,
{{sender}}`,
      Repeat: `Do not hesitate to contact our team on 0488 839 212 if you have any questions.

We look forward to working with you again!

Best wishes,
{{sender}}`
    }
  };
  
  function openEmailComposer(opportunityData, contactData) {
    // Collect emails from Primary Applicant and Applicants
    const emails = [];
    if (contactData.EmailAddress1) emails.push(contactData.EmailAddress1);
    
    // Add emails from Applicants (linked records have email in their data)
    if (opportunityData._applicantEmails && Array.isArray(opportunityData._applicantEmails)) {
      opportunityData._applicantEmails.forEach(email => {
        if (email && !emails.includes(email)) emails.push(email);
      });
    }
    
    // Use appointment data from Appointments table if available, fall back to Taco fields
    const apptData = opportunityData._appointmentData || {};
    
    currentEmailContext = {
      opportunity: opportunityData,
      contact: contactData,
      greeting: contactData.PreferredName || contactData.FirstName || 'there',
      broker: opportunityData['Taco: Broker'] || 'our Mortgage Broker',
      brokerFirst: (opportunityData['Taco: Broker'] || '').split(' ')[0] || 'the broker',
      appointmentTime: formatAppointmentTime(apptData.appointmentTime || opportunityData['Taco: Appointment Time']),
      phoneNumber: apptData.phoneNumber || opportunityData['Taco: Appt Phone Number'] || '[phone number]',
      meetUrl: apptData.meetUrl || opportunityData['Taco: Appt Meet URL'] || '[Google Meet URL]',
      emails: emails,
      sender: 'Shae'
    };
    
    const apptType = apptData.typeOfAppointment || opportunityData['Taco: Type of Appointment'] || 'Phone';
    document.getElementById('emailApptType').value = apptType;
    
    const isNew = opportunityData['Taco: New or Existing Client'] === 'New Client';
    document.getElementById('emailClientType').value = isNew ? 'New' : 'Repeat';
    
    document.getElementById('emailPrepHandler').value = 'Shae';
    
    document.getElementById('emailTo').value = currentEmailContext.emails.join(', ');
    
    const modal = document.getElementById('emailComposer');
    modal.classList.add('visible');
    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        modal.classList.add('showing');
      });
    });
    
    // Initialize Quill editor if not already done
    if (!emailQuill && typeof Quill !== 'undefined') {
      emailQuill = new Quill('#emailPreviewBody', {
        modules: {
          toolbar: '#emailQuillToolbar'
        },
        theme: 'snow'
      });
    }
    
    updateEmailPreview();
  }
  
  function closeEmailComposer() {
    const modal = document.getElementById('emailComposer');
    modal.classList.remove('showing');
    setTimeout(() => {
      modal.classList.remove('visible');
      currentEmailContext = null;
    }, 250);
  }
  
  function sendEmail() {
    if (!currentEmailContext) return;
    
    const to = document.getElementById('emailTo').value;
    const subject = document.getElementById('emailSubject').value;
    // Get HTML content from Quill editor
    const body = emailQuill ? emailQuill.root.innerHTML : document.getElementById('emailPreviewBody').innerHTML;
    
    const sendBtn = document.getElementById('emailSendBtn');
    sendBtn.innerText = 'Sending...';
    sendBtn.disabled = true;
    
    google.script.run.withSuccessHandler(function(result) {
      if (result && result.success) {
        showAlert('Success', 'Email sent successfully!', 'success');
        closeEmailComposer();
      } else {
        showAlert('Error', result?.error || 'Failed to send email. Gmail API integration required.', 'error');
      }
      sendBtn.innerText = 'Send';
      sendBtn.disabled = false;
    }).withFailureHandler(function(err) {
      showAlert('Error', 'Failed to send email: ' + (err.message || 'Gmail API integration required.'), 'error');
      sendBtn.innerText = 'Send';
      sendBtn.disabled = false;
    }).sendEmail(to, subject, body);
  }
  
  function updateEmailPreview() {
    if (!currentEmailContext) return;
    
    const apptType = document.getElementById('emailApptType').value;
    const clientType = document.getElementById('emailClientType').value;
    const prepHandler = document.getElementById('emailPrepHandler').value;
    
    const brokerIntro = clientType === 'New' 
      ? `our Mortgage Broker ${currentEmailContext.broker}`
      : currentEmailContext.brokerFirst;
    
    const variables = {
      greeting: currentEmailContext.greeting,
      broker: currentEmailContext.broker,
      brokerFirst: currentEmailContext.brokerFirst,
      brokerIntro: brokerIntro,
      appointmentTime: currentEmailContext.appointmentTime,
      appointmentType: apptType,
      clientType: clientType,
      prepHandler: prepHandler,
      daysUntil: calculateDaysUntil(currentEmailContext.appointmentTime),
      phoneNumber: currentEmailContext.phoneNumber,
      meetUrl: currentEmailContext.meetUrl,
      sender: currentEmailContext.sender,
      prefillNote: EMAIL_TEMPLATE.prefillNote[clientType],
      officeMapLink: `<a href="${EMAIL_LINKS.officeMap}" target="_blank" style="color:#0066CC;">Office</a>`,
      ourTeamLink: `<a href="${EMAIL_LINKS.ourTeam}" target="_blank" style="color:#0066CC;">here</a>`,
      factFindLink: `<a href="${EMAIL_LINKS.factFind}" target="_blank" style="color:#0066CC;">Fact Find</a>`,
      myGovLink: `<a href="${EMAIL_LINKS.myGov}" target="_blank" style="color:#0066CC;">myGov</a>`,
      myGovVideoLink: `<a href="${EMAIL_LINKS.myGovVideo}" target="_blank" style="color:#0066CC;">this video</a>`,
      incomeInstructionsLink: `<a href="${EMAIL_LINKS.incomeStatementInstructions}" target="_blank" style="color:#0066CC;">click here</a>`
    };
    
    // Try to use Airtable template first
    const confirmationTemplate = getTemplateByType('Confirmation') || getTemplateByName('Appointment Confirmation');
    
    if (confirmationTemplate && confirmationTemplate.body) {
      // Use Airtable template with conditional parser
      const subject = parseConditionalTemplate(confirmationTemplate.subject || EMAIL_TEMPLATE.subject[apptType], variables);
      document.getElementById('emailSubject').value = subject;
      
      let body = parseConditionalTemplate(confirmationTemplate.body, variables);
      
      if (userSignature) {
        body += '<br><br>' + userSignature.replace(/\n/g, '<br>');
      }
      
      if (emailQuill) {
        emailQuill.clipboard.dangerouslyPasteHTML(body);
      } else {
        document.getElementById('emailPreviewBody').innerHTML = body;
      }
    } else {
      // Fallback to hardcoded template
      const subject = replaceVariables(EMAIL_TEMPLATE.subject[apptType], variables);
      document.getElementById('emailSubject').value = subject;
      
      let body = `Hi ${variables.greeting},<br><br>`;
      body += replaceVariables(EMAIL_TEMPLATE.opening[apptType], variables) + '<br><br>';
      body += replaceVariables(EMAIL_TEMPLATE.meetLine[clientType], variables) + '<br><br>';
      body += replaceVariables(EMAIL_TEMPLATE.preparation[prepHandler], variables) + '<br><br>';
      body += replaceVariables(EMAIL_TEMPLATE.closing[clientType], variables);
      
      if (userSignature) {
        body += '<br><br>' + userSignature.replace(/\n/g, '<br>');
      }
      
      if (emailQuill) {
        emailQuill.clipboard.dangerouslyPasteHTML(body);
      } else {
        document.getElementById('emailPreviewBody').innerHTML = body;
      }
    }
  }
  
  function replaceVariables(template, variables) {
    return template.replace(/\{\{(\w+)\}\}/g, (match, key) => variables[key] || match);
  }
  
  // --- AIRTABLE EMAIL TEMPLATES ---
  let airtableTemplates = [];
  let templatesLoaded = false;
  let currentEditingTemplate = null;
  
  // Available template variables with descriptions (for Variable Picker)
  const TEMPLATE_VARIABLES = {
    'Client Info': [
      { name: 'greeting', description: 'Client first name or preferred greeting' },
      { name: 'clientType', description: '"New" or "Repeat"' }
    ],
    'Broker Info': [
      { name: 'broker', description: 'Full broker name' },
      { name: 'brokerFirst', description: 'Broker first name' },
      { name: 'brokerIntro', description: '"our Mortgage Broker [Name]" for new clients, just first name for repeat' },
      { name: 'sender', description: 'Current user (email sender) name' }
    ],
    'Appointment Details': [
      { name: 'appointmentType', description: '"Office", "Phone", or "Video"' },
      { name: 'appointmentTime', description: 'Formatted appointment date/time' },
      { name: 'daysUntil', description: 'Number of days until appointment' },
      { name: 'phoneNumber', description: 'Client phone number' },
      { name: 'meetUrl', description: 'Google Meet URL for video calls' }
    ],
    'Preparation': [
      { name: 'prepHandler', description: '"Shae", "Team", or "OpenBanking"' },
      { name: 'prefillNote', description: 'Auto-fill note based on client type' }
    ],
    'Links': [
      { name: 'officeMapLink', description: 'Office map link (clickable)' },
      { name: 'ourTeamLink', description: 'Team page link (clickable)' },
      { name: 'factFindLink', description: 'Fact Find document link' },
      { name: 'myGovLink', description: 'myGov link' },
      { name: 'myGovVideoLink', description: 'myGov help video link' },
      { name: 'incomeInstructionsLink', description: 'Income statement instructions link' }
    ]
  };
  
  // Condition variables available for {{if}} blocks
  const CONDITION_VARIABLES = [
    { name: 'appointmentType', options: ['Office', 'Phone', 'Video'] },
    { name: 'clientType', options: ['New', 'Repeat'] },
    { name: 'prepHandler', options: ['Shae', 'Team', 'OpenBanking'] }
  ];
  
  function loadEmailTemplates() {
    google.script.run.withSuccessHandler(function(templates) {
      airtableTemplates = templates || [];
      templatesLoaded = true;
      console.log('Email templates loaded:', airtableTemplates.length);
    }).withFailureHandler(function(err) {
      console.error('Failed to load email templates:', err);
      airtableTemplates = [];
      templatesLoaded = true;
    }).getEmailTemplates();
  }
  
  // Load templates on startup
  loadEmailTemplates();
  
  // Parse conditional template syntax and render with context
  function parseConditionalTemplate(template, context) {
    if (!template) return '';
    
    // First, process conditional blocks {{if}}...{{endif}}
    let result = processConditionals(template, context);
    
    // Then replace simple variables {{varName}}
    result = replaceVariables(result, context);
    
    return result;
  }
  
  // Process {{if var=value}}...{{elseif var=value}}...{{else}}...{{endif}} blocks
  function processConditionals(template, context) {
    // Regex to match {{if ...}} blocks including nested content
    const conditionalPattern = /\{\{if\s+(\w+)\s*=\s*(\w+)\}\}([\s\S]*?)\{\{endif\}\}/gi;
    
    return template.replace(conditionalPattern, function(match, varName, varValue, innerContent) {
      // Parse the inner content for elseif/else branches
      const branches = parseConditionalBranches(innerContent);
      const contextValue = context[varName];
      
      // Find matching branch
      for (const branch of branches) {
        if (branch.type === 'if' || branch.type === 'elseif') {
          if (branch.value === contextValue) {
            // Recursively process nested conditionals
            return processConditionals(branch.content, context);
          }
        } else if (branch.type === 'else') {
          // Else is the fallback
          return processConditionals(branch.content, context);
        }
      }
      
      // Check if the initial if condition matches
      if (varValue === contextValue) {
        // Get content before first elseif/else
        const firstBranchContent = innerContent.split(/\{\{(?:elseif|else)/i)[0];
        return processConditionals(firstBranchContent, context);
      }
      
      // No match found, return empty
      return '';
    });
  }
  
  // Parse branches within a conditional block
  function parseConditionalBranches(content) {
    const branches = [];
    
    // Split by elseif and else, keeping the delimiters
    const parts = content.split(/(\{\{elseif\s+\w+\s*=\s*\w+\}\}|\{\{else\}\})/gi);
    
    let currentBranch = { type: 'if', content: parts[0] || '', value: null };
    
    for (let i = 1; i < parts.length; i++) {
      const part = parts[i];
      
      if (/\{\{elseif\s+(\w+)\s*=\s*(\w+)\}\}/i.test(part)) {
        branches.push(currentBranch);
        const match = part.match(/\{\{elseif\s+(\w+)\s*=\s*(\w+)\}\}/i);
        currentBranch = { type: 'elseif', varName: match[1], value: match[2], content: '' };
      } else if (/\{\{else\}\}/i.test(part)) {
        branches.push(currentBranch);
        currentBranch = { type: 'else', content: '' };
      } else {
        currentBranch.content += part;
      }
    }
    
    branches.push(currentBranch);
    return branches;
  }
  
  // Get template by name
  function getTemplateByName(name) {
    return airtableTemplates.find(t => t.name.toLowerCase() === name.toLowerCase());
  }
  
  // Get template by type
  function getTemplateByType(type) {
    return airtableTemplates.find(t => t.type === type);
  }
  
  // --- TEMPLATE EDITOR FUNCTIONS ---
  let templateEditorQuill = null;
  let templateSubjectQuill = null;
  let activeTemplateEditor = 'body'; // 'subject' or 'body'
  let templatePreviewContext = null; // Stores context for live preview
  
  // Sample data for preview when no opportunity is selected
  const SAMPLE_PREVIEW_DATA = {
    greeting: 'Sarah',
    clientType: 'New',
    broker: 'Michael Thompson',
    brokerFirst: 'Michael',
    brokerIntro: 'our Mortgage Broker Michael Thompson',
    sender: 'Shae',
    appointmentType: 'Office',
    appointmentTime: 'Tuesday 15th January at 10:00am',
    daysUntil: '3',
    phoneNumber: '0412 345 678',
    meetUrl: 'https://meet.google.com/abc-defg-hij',
    prepHandler: 'Team',
    prefillNote: '',
    officeMapLink: '<a href="#" style="color:#0066CC;">Office</a>',
    ourTeamLink: '<a href="#" style="color:#0066CC;">here</a>',
    factFindLink: '<a href="#" style="color:#0066CC;">Fact Find</a>',
    myGovLink: '<a href="#" style="color:#0066CC;">myGov</a>',
    myGovVideoLink: '<a href="#" style="color:#0066CC;">myGov Video</a>',
    incomeInstructionsLink: '<a href="#" style="color:#0066CC;">Income Instructions</a>'
  };
  
  // Render template preview with context (shows missing vars with indicator)
  function renderTemplatePreview() {
    if (!templateSubjectQuill || !templateEditorQuill) return;
    
    const subjectText = templateSubjectQuill.getText().trim();
    const bodyHtml = templateEditorQuill.root.innerHTML;
    
    // Build context with control overrides
    const context = getPreviewContextWithOverrides();
    
    // Render subject (plain text)
    const renderedSubject = renderWithMissingIndicators(subjectText, context);
    document.getElementById('templatePreviewSubject').innerHTML = renderedSubject;
    
    // Render body (HTML with conditionals)
    const renderedBody = renderBodyWithMissingIndicators(bodyHtml, context);
    document.getElementById('templatePreviewBody').innerHTML = renderedBody;
  }
  
  // Get preview context with overrides from control dropdowns
  function getPreviewContextWithOverrides() {
    const baseContext = templatePreviewContext || SAMPLE_PREVIEW_DATA;
    
    // Read override values from controls
    const apptTypeEl = document.getElementById('previewApptType');
    const clientTypeEl = document.getElementById('previewClientType');
    const prepHandlerEl = document.getElementById('previewPrepHandler');
    
    const appointmentType = apptTypeEl ? apptTypeEl.value : baseContext.appointmentType;
    const clientType = clientTypeEl ? clientTypeEl.value : baseContext.clientType;
    const prepHandler = prepHandlerEl ? prepHandlerEl.value : baseContext.prepHandler;
    
    // Compute derived values based on overrides
    const brokerIntro = clientType === 'New' 
      ? `our Mortgage Broker ${baseContext.broker}`
      : baseContext.brokerFirst;
    
    const prefillNote = clientType === 'Repeat' 
      ? ' I have prefilled as much as I can using the information we previously received from you.'
      : '';
    
    return {
      ...baseContext,
      appointmentType,
      clientType,
      prepHandler,
      brokerIntro,
      prefillNote
    };
  }
  
  // Called when preview control dropdowns change
  function updateTemplatePreviewFromControls() {
    renderTemplatePreview();
  }
  
  // Initialize preview controls with current context values
  function initializePreviewControls() {
    const context = templatePreviewContext || SAMPLE_PREVIEW_DATA;
    
    const apptTypeEl = document.getElementById('previewApptType');
    const clientTypeEl = document.getElementById('previewClientType');
    const prepHandlerEl = document.getElementById('previewPrepHandler');
    
    if (apptTypeEl) apptTypeEl.value = context.appointmentType || 'Office';
    if (clientTypeEl) clientTypeEl.value = context.clientType || 'New';
    if (prepHandlerEl) prepHandlerEl.value = context.prepHandler || 'Team';
  }
  
  // Render template text with missing variable indicators
  function renderWithMissingIndicators(template, context) {
    if (!template) return '';
    
    // First process conditionals
    let result = processConditionalsForPreview(template, context);
    
    // Then replace variables, showing missing indicator for undefined ones
    result = result.replace(/\{\{(\w+)\}\}/g, (match, varName) => {
      const value = context[varName];
      if (value !== undefined && value !== null && value !== '') {
        return value;
      }
      return `<span class="preview-missing-var">[${varName} not set]</span>`;
    });
    
    return result;
  }
  
  // Render body HTML with missing indicators and condition visualization
  function renderBodyWithMissingIndicators(html, context) {
    if (!html) return '';
    
    // Process conditionals first
    let result = processConditionalsForPreview(html, context);
    
    // Replace variables
    result = result.replace(/\{\{(\w+)\}\}/g, (match, varName) => {
      const value = context[varName];
      if (value !== undefined && value !== null && value !== '') {
        return value;
      }
      return `<span class="preview-missing-var">[${varName} not set]</span>`;
    });
    
    return result;
  }
  
  // Process conditionals for preview with visualization
  function processConditionalsForPreview(template, context) {
    if (!template) return '';
    
    // Match {{if var=value}}...{{endif}} blocks
    const conditionalPattern = /\{\{if\s+(\w+)\s*=\s*(\w+)\}\}([\s\S]*?)\{\{endif\}\}/gi;
    
    return template.replace(conditionalPattern, function(match, varName, varValue, innerContent) {
      try {
        const contextValue = context[varName];
        
        // Parse branches
        const branches = parseConditionalBranches(innerContent);
        
        // Safety check for empty branches
        if (!branches || branches.length === 0) {
          return '';
        }
        
        // Add the initial if condition value
        branches[0].value = varValue;
        
        // Find matching branch
        let matchedContent = '';
        let matchedCondition = '';
        
        for (const branch of branches) {
          if (branch.type === 'if' || branch.type === 'elseif') {
            if (branch.value === contextValue) {
              matchedContent = branch.content;
              matchedCondition = `${varName} = ${branch.value}`;
              break;
            }
          } else if (branch.type === 'else') {
            matchedContent = branch.content;
            matchedCondition = 'else (fallback)';
            break;
          }
        }
        
        // If no branch matched, return nothing
        if (!matchedContent && !matchedCondition) {
          return '';
        }
        
        // Recursively process nested conditionals
        if (matchedContent) {
          matchedContent = processConditionalsForPreview(matchedContent, context);
        }
        
        // Return with visual condition indicator
        if (matchedCondition) {
          return `<div class="preview-condition-block"><div class="preview-condition-label">IF: ${matchedCondition}</div>${matchedContent}</div>`;
        }
        
        return matchedContent || '';
      } catch (err) {
        console.error('Error processing conditional:', err);
        return match; // Return original on error
      }
    });
  }
  
  // Debounce function for preview updates
  let previewUpdateTimeout = null;
  function schedulePreviewUpdate() {
    if (previewUpdateTimeout) clearTimeout(previewUpdateTimeout);
    previewUpdateTimeout = setTimeout(renderTemplatePreview, 150);
  }
  
  function openTemplateEditor(templateId) {
    console.log('openTemplateEditor called with:', templateId);
    const modal = document.getElementById('templateEditorModal');
    if (!modal) {
      console.error('Template editor modal not found');
      return;
    }
    
    try {
      // Initialize Quill for subject line (no toolbar)
      if (!templateSubjectQuill && typeof Quill !== 'undefined') {
        templateSubjectQuill = new Quill('#templateEditorSubject', {
          modules: { toolbar: false },
          theme: 'snow',
          placeholder: 'Email subject with {{variables}}'
        });
        templateSubjectQuill.on('text-change', () => {
          highlightVariables(templateSubjectQuill);
          schedulePreviewUpdate();
        });
        templateSubjectQuill.root.addEventListener('focus', () => { activeTemplateEditor = 'subject'; });
      }
      
      // Initialize Quill for body
      if (!templateEditorQuill && typeof Quill !== 'undefined') {
        templateEditorQuill = new Quill('#templateEditorBody', {
          modules: { toolbar: '#templateEditorToolbar' },
          theme: 'snow'
        });
        templateEditorQuill.on('text-change', () => {
          highlightVariables(templateEditorQuill);
          schedulePreviewUpdate();
        });
        templateEditorQuill.root.addEventListener('focus', () => { activeTemplateEditor = 'body'; });
      }
    } catch (err) {
      console.error('Error initializing Quill editors:', err);
    }
    
    // Populate variable picker
    populateVariablePicker();
    
    // Setup condition variable dropdown handler
    const conditionVarSelect = document.getElementById('conditionVariable');
    conditionVarSelect.onchange = function() {
      updateConditionOptions(this.value);
    };
    
    if (templateId) {
      // Editing existing template
      const template = airtableTemplates.find(t => t.id === templateId);
      if (template) {
        currentEditingTemplate = template;
        document.getElementById('templateEditorTitle').innerText = 'Edit Template';
        document.getElementById('templateEditorName').value = template.name;
        if (templateSubjectQuill) {
          templateSubjectQuill.setText(template.subject || '');
          setTimeout(() => highlightVariables(templateSubjectQuill), 50);
        }
        if (templateEditorQuill) {
          templateEditorQuill.clipboard.dangerouslyPasteHTML(template.body);
          setTimeout(() => highlightVariables(templateEditorQuill), 50);
        }
      }
    } else {
      // New template
      currentEditingTemplate = null;
      document.getElementById('templateEditorTitle').innerText = 'New Template';
      document.getElementById('templateEditorName').value = '';
      if (templateSubjectQuill) templateSubjectQuill.setText('');
      if (templateEditorQuill) templateEditorQuill.setText('');
    }
    
    // Update preview data badge
    const dataBadge = document.getElementById('previewDataSource');
    if (templatePreviewContext) {
      dataBadge.textContent = 'Live Data';
      dataBadge.classList.add('live-data');
    } else {
      dataBadge.textContent = 'Sample Data';
      dataBadge.classList.remove('live-data');
    }
    
    modal.style.display = 'flex';
    setTimeout(() => {
      modal.classList.add('showing');
      // Initialize controls and render preview
      initializePreviewControls();
      setTimeout(renderTemplatePreview, 100);
    }, 10);
  }
  
  let isHighlighting = false;
  function highlightVariables(quillInstance) {
    if (!quillInstance || isHighlighting) return;
    isHighlighting = true;
    
    try {
      const text = quillInstance.getText();
      const regex = /\{\{[^}]+\}\}/g;
      let match;
      
      // Remove existing highlights first
      quillInstance.formatText(0, text.length, 'background', false);
      
      // Apply sky background to all {{variable}} patterns
      while ((match = regex.exec(text)) !== null) {
        quillInstance.formatText(match.index, match[0].length, 'background', '#D0DFE6');
      }
    } finally {
      isHighlighting = false;
    }
  }
  
  function closeTemplateEditor() {
    const modal = document.getElementById('templateEditorModal');
    if (modal) {
      modal.classList.remove('showing');
      setTimeout(() => { modal.style.display = 'none'; }, 250);
    }
    currentEditingTemplate = null;
    templatePreviewContext = null; // Clear preview context
  }
  
  function populateVariablePicker() {
    const container = document.getElementById('variablePickerList');
    if (!container) return;
    
    let html = '';
    for (const [groupName, variables] of Object.entries(TEMPLATE_VARIABLES)) {
      html += `<div class="variable-group">`;
      html += `<div class="variable-group-title">${groupName}</div>`;
      for (const v of variables) {
        html += `<div class="variable-item" onclick="insertVariable('${v.name}')" title="${v.description}">`;
        html += `<span class="variable-item-name">{{${v.name}}}</span>`;
        html += `<span class="variable-item-desc">${v.description}</span>`;
        html += `</div>`;
      }
      html += `</div>`;
    }
    container.innerHTML = html;
  }
  
  function insertVariable(varName) {
    const quill = activeTemplateEditor === 'subject' ? templateSubjectQuill : templateEditorQuill;
    if (!quill) return;
    
    const range = quill.getSelection();
    const insertPos = range ? range.index : quill.getLength() - 1;
    quill.insertText(insertPos, `{{${varName}}}`);
    quill.setSelection(insertPos + varName.length + 4);
  }
  
  function updateConditionOptions(varName) {
    const container = document.getElementById('conditionOptionsContainer');
    const optionsDiv = document.getElementById('conditionOptions');
    
    if (!varName) {
      container.style.display = 'none';
      return;
    }
    
    const condVar = CONDITION_VARIABLES.find(v => v.name === varName);
    if (!condVar) {
      container.style.display = 'none';
      return;
    }
    
    let html = `<div class="condition-option-label">Select options to include:</div>`;
    condVar.options.forEach(opt => {
      html += `<div class="condition-option-row">`;
      html += `<input type="checkbox" id="condOpt_${opt}" value="${opt}" checked>`;
      html += `<span>${opt}</span>`;
      html += `</div>`;
    });
    
    optionsDiv.innerHTML = html;
    container.style.display = 'block';
  }
  
  function insertConditionBlock() {
    if (!templateEditorQuill) return;
    
    const varName = document.getElementById('conditionVariable').value;
    if (!varName) {
      showAlert('Error', 'Please select a variable first', 'error');
      return;
    }
    
    const condVar = CONDITION_VARIABLES.find(v => v.name === varName);
    if (!condVar) return;
    
    // Get selected options
    const selectedOptions = [];
    condVar.options.forEach(opt => {
      const checkbox = document.getElementById(`condOpt_${opt}`);
      if (checkbox && checkbox.checked) {
        selectedOptions.push(opt);
      }
    });
    
    if (selectedOptions.length === 0) {
      showAlert('Error', 'Please select at least one option', 'error');
      return;
    }
    
    // Build conditional block
    let block = '';
    selectedOptions.forEach((opt, idx) => {
      if (idx === 0) {
        block += `{{if ${varName}=${opt}}}\n[Content for ${opt}]\n`;
      } else {
        block += `{{elseif ${varName}=${opt}}}\n[Content for ${opt}]\n`;
      }
    });
    block += `{{endif}}`;
    
    // Insert at cursor
    const range = templateEditorQuill.getSelection();
    const insertPos = range ? range.index : templateEditorQuill.getLength();
    templateEditorQuill.insertText(insertPos, block);
    
    // Reset dropdown
    document.getElementById('conditionVariable').value = '';
    document.getElementById('conditionOptionsContainer').style.display = 'none';
  }
  
  function saveTemplate() {
    const name = document.getElementById('templateEditorName').value.trim();
    const subject = templateSubjectQuill ? templateSubjectQuill.getText().trim() : '';
    const body = templateEditorQuill ? templateEditorQuill.root.innerHTML : '';
    
    if (!name) {
      showAlert('Error', 'Please enter a template name', 'error');
      return;
    }
    
    const btn = document.getElementById('templateEditorSaveBtn');
    btn.innerText = 'Saving...';
    btn.disabled = true;
    
    const fields = { name, subject, body };
    
    if (currentEditingTemplate) {
      // Update existing
      google.script.run.withSuccessHandler(function(result) {
        btn.innerText = 'Save Template';
        btn.disabled = false;
        if (result) {
          showAlert('Success', 'Template saved successfully', 'success');
          loadEmailTemplates();
          closeTemplateEditor();
        } else {
          showAlert('Error', 'Failed to save template', 'error');
        }
      }).withFailureHandler(function(err) {
        btn.innerText = 'Save Template';
        btn.disabled = false;
        showAlert('Error', 'Failed to save template: ' + (err.message || 'Unknown error'), 'error');
      }).updateEmailTemplate(currentEditingTemplate.id, fields);
    } else {
      // Create new
      google.script.run.withSuccessHandler(function(result) {
        btn.innerText = 'Save Template';
        btn.disabled = false;
        if (result) {
          showAlert('Success', 'Template created successfully', 'success');
          loadEmailTemplates();
          closeTemplateEditor();
        } else {
          showAlert('Error', 'Failed to create template', 'error');
        }
      }).withFailureHandler(function(err) {
        btn.innerText = 'Save Template';
        btn.disabled = false;
        showAlert('Error', 'Failed to create template: ' + (err.message || 'Unknown error'), 'error');
      }).createEmailTemplate(fields);
    }
  }
  
  // --- TEMPLATE LIST (Central Hub) ---
  function openTemplateList() {
    const modal = document.getElementById('templateListModal');
    if (!modal) return;
    
    refreshTemplateList();
    
    modal.style.display = 'flex';
    setTimeout(() => modal.classList.add('showing'), 10);
  }
  
  function closeTemplateList() {
    const modal = document.getElementById('templateListModal');
    if (modal) {
      modal.classList.remove('showing');
      setTimeout(() => { modal.style.display = 'none'; }, 250);
    }
  }
  
  function refreshTemplateList() {
    const container = document.getElementById('templateListContainer');
    if (!container) return;
    
    if (airtableTemplates.length === 0) {
      container.innerHTML = `<div style="text-align:center; padding:30px;">
        <div style="color:#888; font-style:italic; margin-bottom:15px;">No templates yet.</div>
        <button type="button" onclick="seedDefaultTemplate()" style="padding:10px 20px; background:var(--color-star); color:white; border:none; border-radius:6px; cursor:pointer; font-size:13px; font-weight:500;">Load Default Confirmation Template</button>
        <div style="font-size:11px; color:#999; margin-top:10px;">This will create the standard Appointment Confirmation template with all variations.</div>
      </div>`;
      return;
    }
    
    let html = '';
    airtableTemplates.forEach(t => {
      html += `<div class="template-list-item">`;
      html += `<div><span class="template-list-name">${t.name}</span></div>`;
      html += `<div class="template-list-actions">`;
      html += `<span class="template-list-type">${t.type || 'General'}</span>`;
      html += `<button class="template-list-edit" onclick="closeTemplateList(); openTemplateEditor('${t.id}')">Edit</button>`;
      html += `</div>`;
      html += `</div>`;
    });
    container.innerHTML = html;
  }
  
  function createNewTemplate() {
    closeTemplateList();
    openTemplateEditor(null);
  }
  
  // Seed the default Appointment Confirmation template
  function seedDefaultTemplate() {
    const defaultSubject = `{{if appointmentType=Office}}Confirmation of Appointment - {{appointmentTime}}{{elseif appointmentType=Phone}}Confirmation of Phone Appointment - {{appointmentTime}}{{elseif appointmentType=Video}}Confirmation of Google Meet Appointment - {{appointmentTime}}{{endif}}`;
    
    const defaultBody = `Hi {{greeting}},<br><br>{{if appointmentType=Office}}I'm writing to confirm your appointment with {{brokerIntro}} on {{appointmentTime}} (Perth time) ({{daysUntil}} days from today) at our office - Kingsley Professional Centre, 18 / 56 Creaney Drive, Kingsley ({{officeMapLink}}).{{elseif appointmentType=Phone}}I'm writing to confirm your phone appointment with {{brokerIntro}} on {{appointmentTime}} (Perth time) ({{daysUntil}} days from today). {{brokerFirst}} will call you on {{phoneNumber}}.{{elseif appointmentType=Video}}I'm writing to confirm your video call appointment with {{brokerIntro}} on {{appointmentTime}} (Perth time) ({{daysUntil}} days from today) using Google Meet URL: {{meetUrl}}. If you have any trouble logging in, please call or text our team on 0488 839 212.{{endif}}<br><br>{{if clientType=New}}Click {{ourTeamLink}} to meet {{brokerFirst}} and the rest of the Team at Stellaris Finance Broking. We will be supporting you each step of the way!{{elseif clientType=Repeat}}Click {{ourTeamLink}} to get reacquainted with the Team at Stellaris Finance Broking. We will be supporting you each step of the way!{{endif}}<br><br>{{if prepHandler=Shae}}In preparation for your appointment, please email me the following information:<br><br>{{factFindLink}} (please note you cannot access this file directly, you will need to download it to your device and fill it in)<br>- Complete with as much detail as possible<br>- Include any Buy Now Pay Later services (like Humm, Zip or Afterpay) that you have accounts with under the Personal Loans section at the bottom of Page 3<br><br>Income<br>a) PAYG Income – Your latest two consecutive payslips and your 2024-25 Income Statement (which can be downloaded from {{myGovLink}}. If you need help creating a myGov account, watch {{myGovVideoLink}}. For instructions on how to download your Income Statement, {{incomeInstructionsLink}})<br>b) Self Employed Income - From each of the last two financial years, your Tax Return, Financial Statements and Notice of Assessment<br><br>I work part time – please try to ensure you email the above evidence well ahead of your appointment to allow ample time to process your information.{{elseif prepHandler=Team}}In preparation for your appointment, please email Shae (shae@stellaris.loans) the following information:<br><br>{{factFindLink}} (please note you cannot access this file directly, you will need to download it to your device and fill it in)<br>- Complete with as much detail as possible<br>- Include any Buy Now Pay Later services (like Humm, Zip or Afterpay) that you have accounts with under the Personal Loans section at the bottom of Page 3<br><br>Income<br>a) PAYG Income – Your latest two consecutive payslips and 2024-25 Income Statement (which can be downloaded from {{myGovLink}}. If you need help creating a myGov account, watch {{myGovVideoLink}}. For instructions on how to download your Income Statement, {{incomeInstructionsLink}})<br>b) Self Employed Income - From each of the last two financial years, your Tax Return, Financial Statements and Notice of Assessment<br><br>Please try to ensure you email the above evidence well ahead of your appointment to allow ample time to process your information.{{elseif prepHandler=OpenBanking}}You will soon receive invitations to share your information with us via Frollo's Open Banking and Connective's Client Centre. These two systems streamline the collection of your key financial and personal data, including your contact details, employment history, savings and liabilities, to give us a complete picture of your situation.{{prefillNote}}{{endif}}<br><br>{{if clientType=New}}Do not hesitate to contact our team on 0488 839 212 if you have any questions.<br><br>We look forward to working with you!<br><br>Best wishes,<br>{{sender}}{{elseif clientType=Repeat}}Do not hesitate to contact our team on 0488 839 212 if you have any questions.<br><br>We look forward to working with you again!<br><br>Best wishes,<br>{{sender}}{{endif}}`;
    
    showAlert('Creating...', 'Creating default template...', 'success');
    
    google.script.run.withSuccessHandler(function(result) {
      if (result) {
        showAlert('Success', 'Default template created! Refresh to see it.', 'success');
        loadEmailTemplates();
        refreshTemplateList();
      } else {
        showAlert('Error', 'Failed to create template', 'error');
      }
    }).withFailureHandler(function(err) {
      showAlert('Error', 'Failed to create template: ' + (err.message || 'Unknown error'), 'error');
    }).createEmailTemplate({
      name: 'Appointment Confirmation',
      type: 'Confirmation',
      subject: defaultSubject,
      body: defaultBody,
      description: 'Standard appointment confirmation email with variations for appointment type, client type, and preparation method.',
      active: true
    });
  }
  
  // Open template editor from email composer context
  function openCurrentTemplateEditor() {
    console.log('openCurrentTemplateEditor called, templates:', airtableTemplates.length);
    
    // If we have email context, use it for live preview
    if (currentEmailContext) {
      const apptType = document.getElementById('emailApptType').value;
      const clientType = document.getElementById('emailClientType').value;
      const prepHandler = document.getElementById('emailPrepHandler').value;
      
      const brokerIntro = clientType === 'New' 
        ? `our Mortgage Broker ${currentEmailContext.broker}`
        : currentEmailContext.brokerFirst;
      
      templatePreviewContext = {
        greeting: currentEmailContext.greeting,
        broker: currentEmailContext.broker,
        brokerFirst: currentEmailContext.brokerFirst,
        brokerIntro: brokerIntro,
        appointmentTime: formatAppointmentTime(currentEmailContext.appointmentTime),
        appointmentType: apptType,
        clientType: clientType,
        prepHandler: prepHandler,
        daysUntil: calculateDaysUntil(currentEmailContext.appointmentTime),
        phoneNumber: currentEmailContext.phoneNumber,
        meetUrl: currentEmailContext.meetUrl,
        sender: currentEmailContext.sender,
        prefillNote: EMAIL_TEMPLATE.prefillNote[clientType] || '',
        officeMapLink: `<a href="${EMAIL_LINKS.officeMap}" target="_blank" style="color:#0066CC;">Office</a>`,
        ourTeamLink: `<a href="${EMAIL_LINKS.ourTeam}" target="_blank" style="color:#0066CC;">here</a>`,
        factFindLink: `<a href="${EMAIL_LINKS.factFind}" target="_blank" style="color:#0066CC;">Fact Find</a>`,
        myGovLink: `<a href="${EMAIL_LINKS.myGov}" target="_blank" style="color:#0066CC;">myGov</a>`,
        myGovVideoLink: `<a href="${EMAIL_LINKS.myGovVideo}" target="_blank" style="color:#0066CC;">myGov Video</a>`,
        incomeInstructionsLink: `<a href="${EMAIL_LINKS.incomeInstructions}" target="_blank" style="color:#0066CC;">click here</a>`
      };
    } else {
      // No email context - use sample data
      templatePreviewContext = null;
    }
    
    // Find the Confirmation template (or first available)
    const confirmationTemplate = airtableTemplates.find(t => 
      t.type === 'Confirmation' || t.name.toLowerCase().includes('confirmation')
    );
    
    if (confirmationTemplate) {
      console.log('Opening confirmation template:', confirmationTemplate.id);
      openTemplateEditor(confirmationTemplate.id);
    } else if (airtableTemplates.length > 0) {
      console.log('Opening first template:', airtableTemplates[0].id);
      openTemplateEditor(airtableTemplates[0].id);
    } else {
      console.log('No templates, opening new');
      // No templates exist, open new template editor
      openTemplateEditor(null);
    }
  }
  
  // Format ISO date string to human-readable Perth time format
  function formatAppointmentTime(dateStr) {
    if (!dateStr) return '[appointment time]';
    
    try {
      // Handle ISO date strings
      if (dateStr.includes('T') || dateStr.includes('Z')) {
        const date = new Date(dateStr);
        if (isNaN(date.getTime())) return dateStr; // Return original if invalid
        
        // Format in Perth timezone (GMT+8)
        const options = {
          weekday: 'long',
          day: 'numeric',
          month: 'long',
          hour: 'numeric',
          minute: '2-digit',
          hour12: true,
          timeZone: 'Australia/Perth'
        };
        
        const formatted = date.toLocaleString('en-AU', options);
        
        // Add ordinal suffix to day
        const dayMatch = formatted.match(/(\d+)/);
        if (dayMatch) {
          const day = parseInt(dayMatch[1]);
          const suffix = getOrdinalSuffix(day);
          return formatted.replace(/(\d+)/, `${day}${suffix}`);
        }
        return formatted;
      }
      
      // Already formatted string, return as-is
      return dateStr;
    } catch (e) {
      console.error('Error formatting date:', e);
      return dateStr;
    }
  }
  
  function getOrdinalSuffix(day) {
    if (day > 3 && day < 21) return 'th';
    switch (day % 10) {
      case 1: return 'st';
      case 2: return 'nd';
      case 3: return 'rd';
      default: return 'th';
    }
  }
  
  function calculateDaysUntil(appointmentTimeStr) {
    if (!appointmentTimeStr) return '?';
    
    try {
      let apptDate;
      
      // Handle ISO date strings
      if (appointmentTimeStr.includes('T') || appointmentTimeStr.includes('Z')) {
        apptDate = new Date(appointmentTimeStr);
        if (isNaN(apptDate.getTime())) return '?';
      } else {
        // Parse human-readable format
        const months = {
          'january': 0, 'jan': 0,
          'february': 1, 'feb': 1,
          'march': 2, 'mar': 2,
          'april': 3, 'apr': 3,
          'may': 4,
          'june': 5, 'jun': 5,
          'july': 6, 'jul': 6,
          'august': 7, 'aug': 7,
          'september': 8, 'sep': 8, 'sept': 8,
          'october': 9, 'oct': 9,
          'november': 10, 'nov': 10,
          'december': 11, 'dec': 11
        };
        
        const cleanStr = appointmentTimeStr.toLowerCase().replace(/,/g, '');
        
        const dayMatch = cleanStr.match(/(\d{1,2})(st|nd|rd|th)?/);
        if (!dayMatch) return '?';
        const day = parseInt(dayMatch[1]);
        
        let monthIdx = -1;
        for (const [name, idx] of Object.entries(months)) {
          if (cleanStr.includes(name)) {
            monthIdx = idx;
            break;
          }
        }
        if (monthIdx === -1) return '?';
        
        const now = new Date();
        let year = now.getFullYear();
        apptDate = new Date(year, monthIdx, day);
        
        const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
        if (apptDate < today) {
          apptDate = new Date(year + 1, monthIdx, day);
        }
      }
      
      // Calculate days difference
      const now = new Date();
      const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
      const apptDateOnly = new Date(apptDate.getFullYear(), apptDate.getMonth(), apptDate.getDate());
      
      const diffTime = apptDateOnly - today;
      const diffDays = Math.round(diffTime / (1000 * 60 * 60 * 24));
      return diffDays >= 0 ? String(diffDays) : '?';
    } catch (e) {
      console.error('Error calculating days until:', e);
      return '?';
    }
  }
  
  function openInGmail() {
    if (!currentEmailContext) return;
    
    const to = document.getElementById('emailTo').value;
    const subject = document.getElementById('emailSubject').value;
    const previewEl = document.getElementById('emailPreviewBody');
    
    // Convert HTML to plain text with URLs preserved
    const body = convertHtmlToPlainTextWithUrls(previewEl.innerHTML);
    
    const gmailUrl = `https://mail.google.com/mail/?view=cm&to=${encodeURIComponent(to)}&su=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
    
    window.open(gmailUrl, '_blank');
    closeEmailComposer();
  }
  
  function convertHtmlToPlainTextWithUrls(html) {
    // Create a temporary element
    const temp = document.createElement('div');
    temp.innerHTML = html;
    
    // Replace <a> tags with "text (url)" format
    const links = temp.querySelectorAll('a');
    links.forEach(link => {
      const text = link.textContent;
      const href = link.getAttribute('href');
      const replacement = document.createTextNode(`${text} (${href})`);
      link.parentNode.replaceChild(replacement, link);
    });
    
    // Replace <br> with newlines
    let result = temp.innerHTML;
    result = result.replace(/<br\s*\/?>/gi, '\n');
    
    // Remove remaining HTML tags
    const div = document.createElement('div');
    div.innerHTML = result;
    return div.textContent || div.innerText || '';
  }
  
  function openEmailComposerFromPanel(opportunityId) {
    google.script.run.withSuccessHandler(function(oppData) {
      if (!oppData) {
        showAlert('Error', 'Could not load opportunity data', 'error');
        return;
      }
      const fields = oppData.fields || {};
      const contactFields = currentContactRecord ? currentContactRecord.fields : {};
      
      // Also fetch appointments for this opportunity
      google.script.run.withSuccessHandler(function(appointments) {
        // Find the first scheduled appointment (future or most recent)
        if (appointments && appointments.length > 0) {
          // Sort by appointment time, prefer scheduled/future appointments
          const scheduled = appointments.find(a => a.status === 'Scheduled' && a.appointmentTime);
          const appt = scheduled || appointments[0];
          if (appt) {
            fields._appointmentData = {
              appointmentTime: appt.appointmentTime,
              typeOfAppointment: appt.typeOfAppointment,
              phoneNumber: appt.phoneNumber,
              meetUrl: appt.meetUrl
            };
          }
        }
        
        // Fetch emails from Primary Applicant and Applicants
        const applicantIds = [];
        if (fields['Primary Applicant'] && fields['Primary Applicant'].length > 0) {
          applicantIds.push(...fields['Primary Applicant']);
        }
        if (fields['Applicants'] && fields['Applicants'].length > 0) {
          applicantIds.push(...fields['Applicants']);
        }
        
        if (applicantIds.length > 0) {
          let fetchedCount = 0;
          const emails = [];
          applicantIds.forEach(id => {
            google.script.run.withSuccessHandler(function(contact) {
              if (contact && contact.fields && contact.fields.EmailAddress1) {
                emails.push(contact.fields.EmailAddress1);
              }
              fetchedCount++;
              if (fetchedCount === applicantIds.length) {
                fields._applicantEmails = emails;
                openEmailComposer(fields, contactFields);
              }
            }).withFailureHandler(function() {
              fetchedCount++;
              if (fetchedCount === applicantIds.length) {
                fields._applicantEmails = emails;
                openEmailComposer(fields, contactFields);
              }
            }).getRecordById('Contacts', id);
          });
        } else {
          openEmailComposer(fields, contactFields);
        }
      }).withFailureHandler(function() {
        // Continue without appointment data
        openEmailComposer(fields, contactFields);
      }).getAppointmentsForOpportunity(opportunityId);
    }).getRecordById('Opportunities', opportunityId);
  }

  // --- SPOUSE LOGIC ---
  function renderSpouseSection(f) {
     const badgeEl = document.getElementById('spouseBadge');
     const statusEl = document.getElementById('spouseStatusText');
     const dateEl = document.getElementById('spouseHistoryDate');
     const accordionEl = document.getElementById('spouseHistoryAccordion');
     const historyList = document.getElementById('spouseHistoryList');
     const arrowEl = document.getElementById('spouseHistoryArrow');

     const spouseName = (f['Spouse Name'] && f['Spouse Name'].length > 0) ? f['Spouse Name'][0] : null;
     const spouseId = (f['Spouse'] && f['Spouse'].length > 0) ? f['Spouse'][0] : null;

     // Reset accordion
     if (accordionEl) accordionEl.style.display = 'none';
     if (historyList) { historyList.innerHTML = ''; historyList.style.display = 'none'; }
     if (arrowEl) arrowEl.classList.remove('expanded');

     // Get connection date from history
     let connectionDate = '';
     const rawLogs = f['Spouse History Text']; 
     let parsedLogs = [];
     if (rawLogs && Array.isArray(rawLogs) && rawLogs.length > 0) {
        parsedLogs = rawLogs.map(parseSpouseHistoryEntry).filter(Boolean);
        parsedLogs.sort((a, b) => b.timestamp - a.timestamp);
        const connLog = parsedLogs.find(e => e.displayText.toLowerCase().includes('connected to'));
        if (connLog) connectionDate = connLog.displayDate;
     }

     if (spouseName && spouseId) {
        if (badgeEl) badgeEl.style.display = 'inline-block';
        statusEl.innerHTML = spouseName;
        statusEl.className = 'connection-name';
        statusEl.setAttribute('data-contact-id', spouseId);
        // Attach quick-view to spouse name
        attachQuickViewToElement(statusEl, spouseId);
        
        // Show history accordion if more than one entry, otherwise show date in left column
        if (parsedLogs.length > 1 && accordionEl && historyList) {
           if (dateEl) dateEl.textContent = '';
           accordionEl.style.display = 'inline-flex';
           parsedLogs.forEach(entry => { renderHistoryItem(entry, historyList); });
        } else {
           if (dateEl) dateEl.textContent = connectionDate;
        }
     } else {
        if (badgeEl) badgeEl.style.display = 'none';
        statusEl.innerHTML = "Single";
        statusEl.className = 'connection-name single';
        statusEl.removeAttribute('data-contact-id');
        if (dateEl) dateEl.textContent = '';
     }
  }
  
  function toggleSpouseHistory() {
     const historyList = document.getElementById('spouseHistoryList');
     const arrowEl = document.getElementById('spouseHistoryArrow');
     if (historyList && arrowEl) {
        const isExpanded = arrowEl.classList.contains('expanded');
        historyList.style.display = isExpanded ? 'none' : 'block';
        arrowEl.classList.toggle('expanded');
     }
  }
  
  function toggleConnectionsAccordion() {
     const content = document.getElementById('connectionsContent');
     const arrowEl = document.getElementById('connectionsAccordionArrow');
     if (content && arrowEl) {
        const isExpanded = arrowEl.classList.contains('expanded');
        content.style.display = isExpanded ? 'none' : 'flex';
        arrowEl.classList.toggle('expanded');
     }
  }
  
  function parseSpouseHistoryEntry(logString) {
     const match = logString.match(/^(\d{4})-(\d{2})-(\d{2})\s+(\d{2}):(\d{2}):(\d{2}):\s*(connected as spouse to|disconnected as spouse from)\s+(.+)$/);
     if (!match) return null;
     const [, year, month, day, hours, mins, secs, action, spouseName] = match;
     const timestamp = new Date(parseInt(year), parseInt(month) - 1, parseInt(day), parseInt(hours), parseInt(mins), parseInt(secs));
     const displayDate = `${day}/${month}/${year}`;
     const shortAction = action.replace(' as spouse', '');
     const displayText = `${shortAction} ${spouseName}`;
     return { timestamp, displayDate, displayText };
  }

  function renderHistoryItem(entry, container) {
     const li = document.createElement('li');
     li.className = 'spouse-history-item';
     li.innerText = `${entry.displayDate}: ${entry.displayText}`;
     const expandLink = container.querySelector('.expand-link');
     if(expandLink) { container.insertBefore(li, expandLink); } else { container.appendChild(li); }
  }

  // --- CONNECTIONS LOGIC ---
  let connectionRoleTypes = [];
  
  let allConnectionsData = [];
  let connectionsExpanded = false;
  const CONNECTIONS_COLLAPSED_LIMIT = 8;
  
  function loadConnections(contactId) {
    const leftList = document.getElementById('connectionsListLeft');
    const rightList = document.getElementById('connectionsListRight');
    if (!leftList || !rightList) return;
    leftList.innerHTML = '<li class="connections-empty">Loading...</li>';
    rightList.innerHTML = '';
    
    google.script.run.withSuccessHandler(function(connections) {
      allConnectionsData = connections || [];
      connectionsExpanded = false;
      renderConnectionsList(allConnectionsData);
    }).withFailureHandler(function(err) {
      leftList.innerHTML = '<li class="connections-empty">Error loading connections</li>';
    }).getConnectionsForContact(contactId);
  }
  
  function toggleConnectionsExpand() {
    connectionsExpanded = !connectionsExpanded;
    renderConnectionsList(allConnectionsData);
  }
  
  function renderConnectionsPills(connections) {
    const container = document.getElementById('connectionsPillContainer');
    if (!container) return;
    container.innerHTML = '';
    
    if (!connections || connections.length === 0) {
      container.innerHTML = '<span class="connections-empty">No connections yet</span>';
      return;
    }
    
    // Short role labels for pills
    const pillRoleLabels = {
      'parent': 'Parent',
      'child': 'Child',
      'sibling': 'Sibling',
      'friend': 'Friend',
      'employer of': 'Employer',
      'employee of': 'Employee',
      'referred by': 'Referrer',
      'has referred': 'Referred',
      'family': 'Family'
    };
    
    const getPillRoleLabel = (role) => {
      const r = (role || '').toLowerCase().trim();
      return pillRoleLabels[r] || role;
    };
    
    const getPillRoleClass = (role) => {
      const r = (role || '').toLowerCase();
      if (r.includes('parent') || r.includes('child')) return 'parent';
      if (r.includes('sibling')) return 'sibling';
      if (r.includes('friend')) return 'friend';
      if (r.includes('employer') || r.includes('employee')) return 'employer';
      if (r.includes('referred') || r.includes('referral')) return 'referred';
      if (r.includes('household')) return 'household';
      if (r.includes('family')) return 'family';
      return '';
    };
    
    // Sort connections: Referrer first, then family roles, then friends/referred
    const roleOrder = ['referred by', 'parent', 'child', 'sibling', 'employer of', 'employee of', 'family', 'friend', 'has referred'];
    connections.sort((a, b) => {
      const aRole = (a.myRole || '').toLowerCase();
      const bRole = (b.myRole || '').toLowerCase();
      const aIdx = roleOrder.findIndex(r => aRole.includes(r));
      const bIdx = roleOrder.findIndex(r => bRole.includes(r));
      if (aIdx !== bIdx) return (aIdx === -1 ? 999 : aIdx) - (bIdx === -1 ? 999 : bIdx);
      return (a.otherContactName || '').localeCompare(b.otherContactName || '');
    });
    
    connections.forEach(conn => {
      const pill = document.createElement('div');
      pill.className = 'connection-pill' + (conn.note && conn.note.trim() ? ' has-note' : '');
      pill.setAttribute('data-conn-id', conn.id);
      pill.setAttribute('data-contact-id', conn.otherContactId || '');
      
      const roleClass = getPillRoleClass(conn.myRole);
      const roleLabel = getPillRoleLabel(conn.myRole);
      
      const deceasedSuffix = conn.otherContactDeceased ? ' (DECEASED)' : '';
      pill.innerHTML = `
        <span class="pill-role ${roleClass}">${roleLabel}</span>
        <span class="pill-name">${conn.otherContactName || 'Unknown'}${deceasedSuffix}</span>
      `;
      if (conn.otherContactDeceased) pill.style.opacity = '0.6';
      
      // Click to open connection details modal
      pill.addEventListener('click', function(e) {
        openConnectionDetailsModal(conn);
      });
      
      // Attach quick-view on hover for the name
      if (conn.otherContactId) {
        attachQuickViewToElement(pill, conn.otherContactId);
      }
      
      container.appendChild(pill);
    });
  }
  
  function renderConnectionsList(connections) {
    const leftList = document.getElementById('connectionsListLeft');
    const rightList = document.getElementById('connectionsListRight');
    const accordionWrapper = document.getElementById('connectionsAccordionWrapper');
    const noAccordionAdd = document.getElementById('connectionsNoAccordionAdd');
    const connectionsContent = document.getElementById('connectionsContent');
    const accordionArrow = document.getElementById('connectionsAccordionArrow');
    if (!leftList || !rightList) return;
    leftList.innerHTML = '';
    rightList.innerHTML = '';
    
    // Count friends and refers to determine if accordion needed
    let friendCount = 0;
    let refersCount = 0;
    if (connections && connections.length > 0) {
      connections.forEach(conn => {
        const role = (conn.myRole || '').toLowerCase().trim();
        if (role === 'friend') friendCount++;
        if (role === 'has referred') refersCount++;
      });
    }
    const needsAccordion = friendCount >= 6 || refersCount >= 6;
    
    // Toggle accordion vs simple add row
    if (accordionWrapper) accordionWrapper.style.display = needsAccordion ? 'block' : 'none';
    if (noAccordionAdd) noAccordionAdd.style.display = needsAccordion ? 'none' : 'flex';
    
    // Ensure accordion is expanded by default and content visible
    if (needsAccordion && connectionsContent && accordionArrow) {
      connectionsAccordionExpanded = true;
      connectionsContent.style.display = 'flex';
      accordionArrow.classList.add('expanded');
    }
    
    if (!connections || connections.length === 0) {
      leftList.innerHTML = '<li class="connections-empty">No connections yet</li>';
      return;
    }
    
    // Role display names with "of" suffix
    const roleDisplayMap = {
      'parent': 'Parent of',
      'child': 'Child of',
      'sibling': 'Sibling of',
      'friend': 'Friend of',
      'employer of': 'Employer of',
      'employee of': 'Employee of',
      'referred by': 'Referred by',
      'has referred': 'Has Referred'
    };
    
    const getDisplayRole = (role) => {
      const r = (role || '').toLowerCase().trim();
      return roleDisplayMap[r] || role;
    };
    
    // Format date as DD/MM/YYYY
    const formatConnectionDate = (dateStr) => {
      if (!dateStr) return '';
      try {
        const d = new Date(dateStr);
        if (isNaN(d.getTime())) return '';
        const day = String(d.getDate()).padStart(2, '0');
        const month = String(d.getMonth() + 1).padStart(2, '0');
        const year = d.getFullYear();
        return `${day}/${month}/${year}`;
      } catch (e) {
        return '';
      }
    };
    
    // Roles that show date
    const rolesWithDate = ['referred by', 'has referred'];
    
    // Define role ordering for left and right columns
    const leftRoles = ['Referred by', 'Parent', 'Child', 'Sibling', 'Employer of', 'Employee of'];
    const rightRoles = ['Friend', 'Has Referred'];
    
    // Roles to group when count > 2
    const groupableRoles = ['friend', 'has referred'];
    
    // Sort connections into left and right buckets
    const leftConns = [];
    const rightConns = [];
    
    connections.forEach(conn => {
      const role = (conn.myRole || '').trim();
      const isLeft = leftRoles.some(r => role.toLowerCase().includes(r.toLowerCase()));
      const isRight = rightRoles.some(r => role.toLowerCase().includes(r.toLowerCase()));
      
      if (isLeft) {
        leftConns.push(conn);
      } else if (isRight) {
        rightConns.push(conn);
      } else {
        leftConns.push(conn);
      }
    });
    
    // Sort each column by role priority then name (or date for Has Referred)
    const sortByRolePriority = (conns, roles, sortHasReferredByDate = false) => {
      return conns.sort((a, b) => {
        const aRole = (a.myRole || '').toLowerCase();
        const bRole = (b.myRole || '').toLowerCase();
        const aIdx = roles.findIndex(r => aRole.includes(r.toLowerCase()));
        const bIdx = roles.findIndex(r => bRole.includes(r.toLowerCase()));
        if (aIdx !== bIdx) return (aIdx === -1 ? 999 : aIdx) - (bIdx === -1 ? 999 : bIdx);
        
        // For Has Referred, sort by date (newest first)
        if (sortHasReferredByDate && aRole === 'has referred' && bRole === 'has referred') {
          const dateA = a.createdOn ? new Date(a.createdOn) : new Date(0);
          const dateB = b.createdOn ? new Date(b.createdOn) : new Date(0);
          return dateB - dateA;
        }
        
        return (a.otherContactName || '').localeCompare(b.otherContactName || '');
      });
    };
    
    sortByRolePriority(leftConns, leftRoles);
    sortByRolePriority(rightConns, rightRoles, true);
    
    // Group connections by role for groupable roles
    const groupByRole = (conns) => {
      const groups = {};
      const ungrouped = [];
      
      conns.forEach(conn => {
        const role = (conn.myRole || '').toLowerCase().trim();
        const isGroupable = groupableRoles.some(g => role === g);
        
        if (isGroupable) {
          if (!groups[role]) groups[role] = [];
          groups[role].push(conn);
        } else {
          ungrouped.push(conn);
        }
      });
      
      return { groups, ungrouped };
    };
    
    // Build connection data attribute for modal
    const buildConnDataAttr = (conn) => {
      return `data-conn-id="${conn.id}" data-conn-name="${escapeHtmlForAttr(conn.otherContactName)}" data-conn-created="${conn.createdOn || ''}" data-conn-modified="${conn.modifiedOn || ''}"`;
    };
    
    // Render individual connection
    const renderSingleConnection = (list, conn) => {
      const li = document.createElement('li');
      li.className = 'connection-item connection-clickable';
      li.setAttribute('data-conn-id', conn.id);
      li.setAttribute('data-conn-name', conn.otherContactName || '');
      li.setAttribute('data-conn-created', conn.createdOn || '');
      li.setAttribute('data-conn-modified', conn.modifiedOn || '');
      li.setAttribute('data-conn-note', conn.note || '');
      
      const badgeClass = getRoleBadgeClass(conn.myRole);
      const displayRole = getDisplayRole(conn.myRole);
      const role = (conn.myRole || '').toLowerCase().trim();
      const showDate = rolesWithDate.includes(role);
      const dateDisplay = showDate ? formatConnectionDate(conn.createdOn) : '';
      const hasNote = conn.note && conn.note.trim();
      const deceasedSuffix = conn.otherContactDeceased ? ' (DECEASED)' : '';
      
      li.innerHTML = `
        <div class="connection-info">
          <span class="connection-role-badge ${badgeClass}">${displayRole}</span>
          <span class="connection-name" data-contact-id="${conn.otherContactId || ''}">${conn.otherContactName}${deceasedSuffix}</span>
          ${dateDisplay ? `<span class="connection-date">${dateDisplay}</span>` : ''}
        </div>
        <button type="button" class="conn-note-icon ${hasNote ? 'has-note' : ''}" data-conn-id="${conn.id}" title="Add/view note"></button>
      `;
      if (conn.otherContactDeceased) li.style.opacity = '0.6';
      
      // Note icon click handler - read from data attribute for updated values
      const noteIcon = li.querySelector('.conn-note-icon');
      if (noteIcon) {
        noteIcon.addEventListener('click', function(e) {
          e.stopPropagation();
          const currentNote = li.getAttribute('data-conn-note') || '';
          openConnectionNotePopover(this, conn.id, currentNote);
        });
      }
      
      // Click handler for the whole item (except the name and note icon)
      li.addEventListener('click', function(e) {
        if (e.target.classList.contains('connection-name') || e.target.classList.contains('conn-note-icon')) {
          e.stopPropagation();
          // Let quick-view or note handle it
        } else {
          openConnectionDetailsModal(conn);
        }
      });
      
      // Attach quick-view to the connection name
      const nameEl = li.querySelector('.connection-name');
      if (nameEl && conn.otherContactId) {
        attachQuickViewToElement(nameEl, conn.otherContactId);
      }
      
      list.appendChild(li);
    };
    
    // Render grouped connections
    const renderGroupedConnections = (list, role, conns) => {
      if (conns.length === 0) return;
      
      // Sort Has Referred by date (newest first) within the group
      if (role === 'has referred') {
        conns.sort((a, b) => {
          const dateA = a.createdOn ? new Date(a.createdOn) : new Date(0);
          const dateB = b.createdOn ? new Date(b.createdOn) : new Date(0);
          return dateB - dateA;
        });
      }
      
      const badgeClass = getRoleBadgeClass(role);
      const groupId = `conn-group-${role.replace(/\s+/g, '-')}`;
      const showDate = rolesWithDate.includes(role);
      
      // Header text
      const headerText = role === 'friend' 
        ? `Friend of ${conns.length} Contacts`
        : `Has Referred ${conns.length} Contacts`;
      
      const li = document.createElement('li');
      li.className = 'connection-group';
      li.innerHTML = `
        <div class="connection-group-header" onclick="toggleConnectionGroup('${groupId}')">
          <span class="connection-role-badge ${badgeClass}">${headerText}</span>
          <span class="connection-group-toggle" id="${groupId}-toggle">▼</span>
        </div>
        <ul class="connection-group-list" id="${groupId}" style="display: none;"></ul>
      `;
      list.appendChild(li);
      
      const subList = li.querySelector(`#${groupId}`);
      conns.forEach(conn => {
        const subLi = document.createElement('li');
        subLi.className = 'connection-subitem connection-clickable';
        subLi.setAttribute('data-conn-id', conn.id);
        subLi.setAttribute('data-conn-name', conn.otherContactName || '');
        subLi.setAttribute('data-conn-created', conn.createdOn || '');
        subLi.setAttribute('data-conn-modified', conn.modifiedOn || '');
        subLi.setAttribute('data-conn-note', conn.note || '');
        
        const dateDisplay = showDate ? formatConnectionDate(conn.createdOn) : '';
        const hasNote = conn.note && conn.note.trim();
        const deceasedSuffix = conn.otherContactDeceased ? ' (DECEASED)' : '';
        
        subLi.innerHTML = `
          <div class="connection-subitem-content">
            <span class="connection-name" data-contact-id="${conn.otherContactId || ''}">${conn.otherContactName}${deceasedSuffix}</span>
            ${dateDisplay ? `<span class="connection-date">${dateDisplay}</span>` : ''}
          </div>
          <button type="button" class="conn-note-icon ${hasNote ? 'has-note' : ''}" data-conn-id="${conn.id}" title="Add/view note"></button>
        `;
        if (conn.otherContactDeceased) subLi.style.opacity = '0.6';
        
        // Note icon click handler - read from data attribute for updated values
        const noteIcon = subLi.querySelector('.conn-note-icon');
        if (noteIcon) {
          noteIcon.addEventListener('click', function(e) {
            e.stopPropagation();
            const currentNote = subLi.getAttribute('data-conn-note') || '';
            openConnectionNotePopover(this, conn.id, currentNote);
          });
        }
        
        subLi.addEventListener('click', function(e) {
          if (e.target.classList.contains('connection-name') || e.target.classList.contains('conn-note-icon')) {
            e.stopPropagation();
            // Let quick-view or note handle it
          } else {
            openConnectionDetailsModal(conn);
          }
        });
        
        // Attach quick-view to the connection name
        const nameEl = subLi.querySelector('.connection-name');
        if (nameEl && conn.otherContactId) {
          attachQuickViewToElement(nameEl, conn.otherContactId);
        }
        
        subList.appendChild(subLi);
      });
    };
    
    // Render function with grouping logic
    const renderToList = (list, conns) => {
      if (conns.length === 0) return;
      
      const { groups, ungrouped } = groupByRole(conns);
      
      // Render ungrouped first
      ungrouped.forEach(conn => renderSingleConnection(list, conn));
      
      // Render grouped (only if more than 2)
      Object.keys(groups).forEach(role => {
        const roleConns = groups[role];
        if (roleConns.length > 2) {
          renderGroupedConnections(list, role, roleConns);
        } else {
          roleConns.forEach(conn => renderSingleConnection(list, conn));
        }
      });
    };
    
    // Render all connections (scrollable container handles overflow)
    // Render each connection individually (no grouping)
    leftConns.forEach(conn => renderSingleConnection(leftList, conn));
    rightConns.forEach(conn => renderSingleConnection(rightList, conn));
  }
  
  function openConnectionDetailsModal(conn) {
    const modal = document.getElementById('deactivateConnectionModal');
    const title = modal.querySelector('.modal-title');
    const body = modal.querySelector('.modal-body-content');
    
    // Get the current contact's full name from fields (use Calculated Name if available)
    let currentContactName = 'Contact';
    if (currentContactRecord && currentContactRecord.fields) {
      const f = currentContactRecord.fields;
      currentContactName = f['Calculated Name'] || 
        `${f.FirstName || ''} ${f.MiddleName || ''} ${f.LastName || ''}`.replace(/\s+/g, ' ').trim();
    }
    
    // Get role display with "of" suffix
    const roleDisplayMap = {
      'parent': 'Parent of',
      'child': 'Child of',
      'sibling': 'Sibling of',
      'friend': 'Friend of',
      'employer of': 'Employer of',
      'employee of': 'Employee of',
      'referred by': 'Referred by',
      'has referred': 'Has Referred'
    };
    const myRoleLower = (conn.myRole || '').toLowerCase().trim();
    const displayRole = roleDisplayMap[myRoleLower] || conn.myRole;
    
    // Format dates in Perth timezone (GMT+8)
    const formatAuditDateTime = (dateStr) => {
      if (!dateStr) return '';
      try {
        const d = new Date(dateStr);
        if (isNaN(d.getTime())) return '';
        // Convert to Perth time (GMT+8)
        const perthOffset = 8 * 60;
        const localOffset = d.getTimezoneOffset();
        const perthTime = new Date(d.getTime() + (perthOffset + localOffset) * 60000);
        const hours = String(perthTime.getHours()).padStart(2, '0');
        const mins = String(perthTime.getMinutes()).padStart(2, '0');
        const day = String(perthTime.getDate()).padStart(2, '0');
        const month = String(perthTime.getMonth() + 1).padStart(2, '0');
        const year = perthTime.getFullYear();
        return `${hours}:${mins} ${day}/${month}/${year}`;
      } catch (e) {
        return '';
      }
    };
    
    const createdDateTime = formatAuditDateTime(conn.createdOn);
    const modifiedDateTime = formatAuditDateTime(conn.modifiedOn);
    const createdBy = conn.createdByName || '';
    const modifiedBy = conn.modifiedByName || '';
    
    const createdText = createdDateTime ? `${createdDateTime}${createdBy ? ' by ' + createdBy : ''}` : 'Unknown';
    const modifiedText = modifiedDateTime ? `${modifiedDateTime}${modifiedBy ? ' by ' + modifiedBy : ''}` : 'Unknown';
    
    // Title: "Current Contact: Role of Other Contact"
    title.textContent = `${currentContactName}: ${displayRole} ${conn.otherContactName}`;
    
    body.innerHTML = `
      <div class="panel-audit-section" style="margin-bottom: 15px; text-align: left;">
        <div><span class="audit-label">Created</span> <span class="audit-value">${createdText}</span></div>
        <div><span class="audit-label">Modified</span> <span class="audit-value">${modifiedText}</span></div>
      </div>
      <div class="connection-modal-remove">
        <span class="remove-label">Remove this connection?</span>
        <button type="button" class="btn-danger conn-modal-btn" onclick="executeDeactivateConnection()">Remove</button>
      </div>
    `;
    
    // Add close button to modal-body (parent), not modal-body-content
    const modalBody = modal.querySelector('.modal-body');
    let closeDiv = modalBody.querySelector('.connection-modal-close');
    if (!closeDiv) {
      closeDiv = document.createElement('div');
      closeDiv.className = 'connection-modal-close';
      closeDiv.innerHTML = '<button type="button" class="btn-secondary conn-modal-btn" onclick="closeDeactivateConnectionModal()">Close</button>';
      modalBody.appendChild(closeDiv);
    }
    
    // Store connection info for the confirm action
    modal.setAttribute('data-conn-id', conn.id);
    modal.setAttribute('data-conn-name', conn.otherContactName || '');
    
    modal.classList.add('visible');
    setTimeout(() => modal.classList.add('showing'), 10);
  }
  
  window.toggleConnectionGroup = function(groupId) {
    const list = document.getElementById(groupId);
    const toggle = document.getElementById(groupId + '-toggle');
    if (list && toggle) {
      const isHidden = list.style.display === 'none';
      list.style.display = isHidden ? 'block' : 'none';
      toggle.textContent = isHidden ? '▲' : '▼';
    }
  };
  
  function getRoleBadgeClass(role) {
    if (!role) return '';
    const r = role.toLowerCase();
    if (r === 'parent' || r === 'child') return r;
    if (r === 'sibling') return 'sibling';
    if (r === 'friend') return 'friend';
    if (r.includes('employer')) return 'employer';
    if (r.includes('employee')) return 'employee';
    if (r.includes('referred') || r.includes('referral')) return 'referred';
    if (r.includes('household')) return 'household';
    return '';
  }
  
  function escapeHtmlForAttr(str) {
    if (!str) return '';
    return str.replace(/'/g, "&#39;").replace(/"/g, '&quot;');
  }
  
  function escapeHtml(str) {
    if (!str) return '';
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
  }
  
  function unescapeHtml(str) {
    if (!str) return '';
    return str.replace(/&#39;/g, "'").replace(/&quot;/g, '"');
  }
  
  function openAddConnectionModal() {
    if (!currentContactRecord) return;
    
    const modal = document.getElementById('addConnectionModal');
    document.getElementById('connectionStep1').style.display = 'flex';
    document.getElementById('connectionStep2').style.display = 'none';
    document.getElementById('connectionSearchInput').value = '';
    document.getElementById('connectionSearchResults').innerHTML = '';
    document.getElementById('connectionSearchResults').style.display = 'none';
    
    // Load role types if not already loaded
    if (connectionRoleTypes.length === 0) {
      google.script.run.withSuccessHandler(function(types) {
        connectionRoleTypes = types || [];
        populateConnectionRoleSelect();
      }).getConnectionRoleTypes();
    }
    
    modal.classList.add('visible');
    setTimeout(() => modal.classList.add('showing'), 10);
    
    // Pre-load recent contacts
    loadRecentContactsForConnectionModal();
  }
  
  function closeAddConnectionModal() {
    const modal = document.getElementById('addConnectionModal');
    modal.classList.remove('showing');
    setTimeout(() => modal.classList.remove('visible'), 250);
  }
  
  function loadRecentContactsForConnectionModal() {
    const results = document.getElementById('connectionSearchResults');
    results.innerHTML = '<div class="search-option" style="color:#999; font-style:italic;">Loading recent contacts...</div>';
    results.style.display = 'block';
    
    google.script.run.withSuccessHandler(function(contacts) {
      renderConnectionSearchResults(contacts);
    }).getRecentContacts();
  }
  
  function handleConnectionSearch(event) {
    const query = event.target.value.trim();
    if (query.length < 2) {
      loadRecentContactsForConnectionModal();
      return;
    }
    
    google.script.run.withSuccessHandler(function(contacts) {
      renderConnectionSearchResults(contacts);
    }).searchContacts(query);
  }
  
  function renderConnectionSearchResults(contacts) {
    const results = document.getElementById('connectionSearchResults');
    results.innerHTML = '';
    results.style.display = 'block';
    
    if (!contacts || contacts.length === 0) {
      results.innerHTML = '<div class="search-option" style="color:#999;">No contacts found</div>';
      return;
    }
    
    // Filter out current contact
    const currentId = currentContactRecord?.id;
    const filtered = contacts.filter(c => c.id !== currentId);
    
    if (filtered.length === 0) {
      results.innerHTML = '<div class="search-option" style="color:#999;">No other contacts found</div>';
      return;
    }
    
    filtered.slice(0, 15).forEach(contact => {
      const f = contact.fields;
      const name = `${f.FirstName || ''} ${f.MiddleName || ''} ${f.LastName || ''}`.replace(/\s+/g, ' ').trim();
      const div = document.createElement('div');
      div.className = 'search-option';
      div.innerHTML = `<strong>${name}</strong>${f.EmailAddress1 ? `<br><span style="font-size:11px; color:#666;">${f.EmailAddress1}</span>` : ''}`;
      div.onclick = function() { selectConnectionTarget(contact.id, name); };
      results.appendChild(div);
    });
  }
  
  function selectConnectionTarget(contactId, contactName) {
    document.getElementById('targetConnectionContactId').value = contactId;
    document.getElementById('targetContactNameConn').innerText = contactName;
    
    // Set current contact name
    const f = currentContactRecord.fields;
    const currentName = `${f.FirstName || ''} ${f.LastName || ''}`.trim();
    document.getElementById('currentContactNameConn').innerText = currentName;
    
    // Populate role select and show step 2
    populateConnectionRoleSelect();
    
    document.getElementById('connectionStep1').style.display = 'none';
    document.getElementById('connectionStep2').style.display = 'flex';
  }
  
  function populateConnectionRoleSelect() {
    const select = document.getElementById('connectionRoleSelect');
    select.innerHTML = '<option value="">-- Select relationship --</option>';
    
    connectionRoleTypes.forEach((pair, index) => {
      const option = document.createElement('option');
      option.value = index;
      option.textContent = `${pair.role1} of this contact`;
      select.appendChild(option);
    });
  }
  
  function backToConnectionStep1() {
    document.getElementById('connectionStep1').style.display = 'flex';
    document.getElementById('connectionStep2').style.display = 'none';
  }
  
  function executeCreateConnection() {
    const selectEl = document.getElementById('connectionRoleSelect');
    const roleIndex = selectEl.value;
    if (roleIndex === '') {
      alert('Please select a relationship type');
      return;
    }
    
    const pair = connectionRoleTypes[parseInt(roleIndex)];
    const contact1Id = currentContactRecord.id;
    const contact2Id = document.getElementById('targetConnectionContactId').value;
    
    const btn = document.getElementById('confirmConnectionBtn');
    btn.disabled = true;
    btn.innerText = 'Creating...';
    
    google.script.run.withSuccessHandler(function(result) {
      btn.disabled = false;
      btn.innerText = 'Create Connection';
      
      if (result.success) {
        closeAddConnectionModal();
        loadConnections(currentContactRecord.id);
      } else {
        alert('Error: ' + (result.error || 'Failed to create connection'));
      }
    }).withFailureHandler(function(err) {
      btn.disabled = false;
      btn.innerText = 'Create Connection';
      alert('Error: ' + err.message);
    }).createConnection(contact1Id, contact2Id, pair.role1, pair.role2);
  }
  
  function closeDeactivateConnectionModal() {
    const modal = document.getElementById('deactivateConnectionModal');
    modal.classList.remove('showing');
    setTimeout(() => modal.classList.remove('visible'), 250);
  }
  
  function executeDeactivateConnection() {
    const modal = document.getElementById('deactivateConnectionModal');
    const connectionId = modal.getAttribute('data-conn-id');
    
    if (!connectionId) {
      closeDeactivateConnectionModal();
      return;
    }
    
    google.script.run.withSuccessHandler(function(result) {
      closeDeactivateConnectionModal();
      if (result.success) {
        loadConnections(currentContactRecord.id);
      } else {
        alert('Error: ' + (result.error || 'Failed to remove connection'));
      }
    }).withFailureHandler(function(err) {
      closeDeactivateConnectionModal();
      alert('Error: ' + err.message);
    }).deactivateConnection(connectionId);
  }

  // --- INLINE EDIT LOGIC ---
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
           if(currentContactRecord) { loadOpportunities(currentContactRecord.fields); }
        }
        if(fieldKey === 'Status' && val === 'Won') {
           triggerWonCelebration();
        }
        if(fieldKey === 'Status' && currentContactRecord) {
           loadOpportunities(currentContactRecord.fields);
        }
        // Toggle conditional Taco fields
        if(fieldKey === 'Taco: Type of Appointment') {
           const phoneWrap = document.getElementById('field_wrap_Taco: Appt Phone Number');
           const videoWrap = document.getElementById('field_wrap_Taco: Appt Meet URL');
           if (phoneWrap) phoneWrap.style.display = val === 'Phone' ? '' : 'none';
           if (videoWrap) videoWrap.style.display = val === 'Video' ? '' : 'none';
        }
        if(fieldKey === 'Taco: How appt booked') {
           const otherWrap = document.getElementById('field_wrap_Taco: How Appt Booked Other');
           if (otherWrap) otherWrap.style.display = val === 'Other' ? '' : 'none';
           // Update Appt Reminder label reactively based on Calendly selection
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
  function resetForm() {
    toggleProfileView(true); document.getElementById('contactForm').reset();
    document.getElementById('recordId').value = "";
    // Explicitly enable new contact mode
    enableNewContactMode();
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
  function handleSearch(event) {
    // Ignore navigation keys - don't reset highlight or search on arrow/enter
    if (['ArrowDown', 'ArrowUp', 'Enter', 'Escape'].includes(event.key)) return;
    
    const query = event.target.value; const statusEl = document.getElementById('searchStatus');
    clearTimeout(loadingTimer);
    searchHighlightIndex = -1; // Reset keyboard navigation when typing
    showSearchDropdown(); // Reopen dropdown when typing
    if(query.length === 0) { statusEl.innerText = ""; loadContacts(); return; }
    clearTimeout(searchTimeout); statusEl.innerText = "Typing...";
    searchTimeout = setTimeout(() => {
      statusEl.innerText = "Searching...";
      const statusFilterToSend = contactStatusFilter === 'All' ? null : contactStatusFilter;
      google.script.run.withSuccessHandler(function(records) {
         statusEl.innerText = records.length > 0 ? `Found ${records.length} matches` : "No matches found";
         renderList(records);
      }).searchContacts(query, statusFilterToSend);
    }, 500);
  }
  function loadContacts() {
    const loadingDiv = document.getElementById('loading'); const list = document.getElementById('contactList');
    list.innerHTML = ''; loadingDiv.style.display = 'block'; loadingDiv.innerHTML = 'Loading directory...';
    clearTimeout(loadingTimer);

    // Show a simpler retry message if loading takes too long
    loadingTimer = setTimeout(() => { 
       loadingDiv.innerHTML = `
         <div style="margin-top:15px; text-align:center;">
           <p style="color:#666; font-size:13px;">Taking a while to connect...</p>
           <button onclick="loadContacts()" style="padding:8px 16px; background:var(--color-cedar); color:white; border:none; border-radius:4px; cursor:pointer; font-size:12px; margin-top:8px;">Try Again</button>
         </div>
       `; 
    }, 4000);

    const statusFilterToSend = contactStatusFilter === 'All' ? null : contactStatusFilter;
    google.script.run.withSuccessHandler(function(records) {
         clearTimeout(loadingTimer); document.getElementById('loading').style.display = 'none';
         if (!records || records.length === 0) { 
           list.innerHTML = '<li style="padding:20px; color:#999; text-align:center; font-size:13px;">No contacts found</li>'; 
           return; 
         }
         renderList(records);
      }).getRecentContacts(statusFilterToSend);
  }
  function renderList(records) {
    const list = document.getElementById('contactList'); 
    document.getElementById('loading').style.display = 'none'; 
    list.innerHTML = '';
    currentSearchRecords = records;
    // Preserve highlight if valid, otherwise reset
    if (searchHighlightIndex >= records.length) {
      searchHighlightIndex = records.length > 0 ? records.length - 1 : -1;
    }
    records.forEach((record, index) => {
      const f = record.fields; const item = document.createElement('li'); item.className = 'contact-item';
      item.dataset.index = index;
      const fullName = formatName(f);
      const initials = getInitials(f.FirstName, f.LastName);
      const avatarColor = getAvatarColor(fullName);
      const modifiedTooltip = formatModifiedTooltip(f);
      const modifiedShort = formatModifiedShort(f);
      const isDeceased = f.Deceased === true;
      const deceasedBadge = isDeceased ? '<span class="deceased-badge-small">DECEASED</span>' : '';
      item.innerHTML = `<div class="contact-avatar" style="background-color:${avatarColor}">${initials}</div><div class="contact-info"><span class="contact-name">${fullName}${deceasedBadge}</span><div class="contact-details-row">${formatDetailsRow(f)}</div></div>${modifiedShort ? `<span class="contact-modified" title="${modifiedTooltip || ''}">${modifiedShort}</span>` : ''}`;
      if (isDeceased) item.style.opacity = '0.6';
      item.onclick = function() { selectContact(record); }; list.appendChild(item);
    });
    // Re-apply highlight if one was preserved
    if (searchHighlightIndex >= 0) {
      updateSearchHighlight();
    }
  }
  
  function updateSearchHighlight() {
    const items = document.querySelectorAll('#contactList .contact-item');
    items.forEach((item, i) => {
      item.classList.toggle('keyboard-highlight', i === searchHighlightIndex);
    });
    if (searchHighlightIndex >= 0 && items[searchHighlightIndex]) {
      items[searchHighlightIndex].scrollIntoView({ block: 'nearest', behavior: 'smooth' });
    }
  }
  
  function handleSearchKeydown(e) {
    const dropdown = document.getElementById('searchDropdown');
    if (!dropdown || !dropdown.classList.contains('open')) return;
    
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      if (currentSearchRecords.length > 0) {
        searchHighlightIndex = Math.min(searchHighlightIndex + 1, currentSearchRecords.length - 1);
        updateSearchHighlight();
      }
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      if (currentSearchRecords.length > 0) {
        searchHighlightIndex = Math.max(searchHighlightIndex - 1, 0);
        updateSearchHighlight();
      }
    } else if (e.key === 'Enter') {
      e.preventDefault();
      if (searchHighlightIndex >= 0 && currentSearchRecords[searchHighlightIndex]) {
        selectContact(currentSearchRecords[searchHighlightIndex]);
      } else {
        // No item highlighted - trigger search immediately
        const searchInput = document.getElementById('searchInput');
        if (searchInput && searchInput.value.trim()) {
          clearTimeout(searchTimeout);
          const statusEl = document.getElementById('searchStatus');
          statusEl.innerText = "Searching...";
          const statusFilterToSend = contactStatusFilter === 'All' ? null : contactStatusFilter;
          google.script.run.withSuccessHandler(function(records) {
            statusEl.innerText = records.length > 0 ? `Found ${records.length} matches` : "No matches found";
            renderList(records);
          }).searchContacts(searchInput.value.trim(), statusFilterToSend);
        }
      }
    }
  }
  function formatName(f) {
    return `${f.FirstName || ''} ${f.MiddleName || ''} ${f.LastName || ''}`.replace(/\s+/g, ' ').trim();
  }
  function formatDetailsRow(f) {
    const parts = [];
    if (f.EmailAddress1) parts.push(`<span>${f.EmailAddress1}</span>`);
    if (f.Mobile) parts.push(`<span>${f.Mobile}</span>`);
    return parts.join('');
  }
  function parseModifiedFormula(modifiedStr) {
    // Parse "Modified: HH:MM DD/MM/YYYY by Name" or "HH:MM DD/MM/YYYY by Name"
    if (!modifiedStr) return null;
    const match = modifiedStr.match(/(\d{2}):(\d{2})\s+(\d{2})\/(\d{2})\/(\d{4})/);
    if (!match) return null;
    const [, hours, mins, day, month, year] = match;
    // Treat as Perth time (GMT+8)
    const utcMs = Date.UTC(parseInt(year), parseInt(month) - 1, parseInt(day), parseInt(hours), parseInt(mins));
    const perthOffsetMs = 8 * 60 * 60 * 1000;
    return new Date(utcMs - perthOffsetMs);
  }
  function formatModifiedTooltip(f) {
    const modified = f.Modified;
    if (!modified) return null;
    
    const modDate = parseModifiedFormula(modified);
    if (!modDate) return null;
    
    // Extract "by Name" part
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
  }
  function formatModifiedShort(f) {
    const modified = f.Modified;
    if (!modified) return null;
    
    const modDate = parseModifiedFormula(modified);
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

  function loadOpportunities(f) {
    const oppList = document.getElementById('oppList'); const loader = document.getElementById('oppLoading');
    document.getElementById('oppSortBtn').style.display = 'none'; oppList.innerHTML = ''; loader.style.display = 'block';
    let oppsToFetch = []; let roleMap = {};
    const addIds = (ids, roleName) => { if(!ids) return; (Array.isArray(ids) ? ids : [ids]).forEach(id => { oppsToFetch.push(id); roleMap[id] = roleName; }); };
    addIds(f['Opportunities - Primary Applicant'], 'Primary Applicant');
    addIds(f['Opportunities - Applicant'], 'Applicant');
    addIds(f['Opportunities - Guarantor'], 'Guarantor');
    if(oppsToFetch.length === 0) { loader.style.display = 'none'; oppList.innerHTML = '<li style="color:#CCC; font-size:12px; font-style:italic;">No opportunities linked.</li>'; return; }
    google.script.run.withSuccessHandler(function(oppRecords) {
       loader.style.display = 'none';
       oppRecords.forEach(r => r._role = roleMap[r.id] || "Linked");
       if(oppRecords.length > 1) { document.getElementById('oppSortBtn').style.display = 'inline'; }
       currentOppRecords = oppRecords; renderOppList();
    }).getLinkedOpportunities(oppsToFetch);
  }
  function toggleOppSort() {
     if(currentOppSortDirection === 'asc') currentOppSortDirection = 'desc'; else currentOppSortDirection = 'asc';
     renderOppList();
  }
  function renderOppList() {
     const oppList = document.getElementById('oppList'); oppList.innerHTML = '';
     const sorted = [...currentOppRecords].sort((a, b) => {
         const nameA = (a.fields['Opportunity Name'] || "").toLowerCase();
         const nameB = (b.fields['Opportunity Name'] || "").toLowerCase();
         if(currentOppSortDirection === 'asc') return nameA.localeCompare(nameB); return nameB.localeCompare(nameA);
     });
     sorted.forEach(opp => {
         const fields = opp.fields; const name = fields['Opportunity Name'] || "Unnamed Opportunity"; const role = opp._role;
         const status = fields['Status'] || '';
         const oppType = fields['Opportunity Type'] || '';
         const statusClass = status === 'Won' ? 'status-won' : status === 'Lost' ? 'status-lost' : '';
         const li = document.createElement('li'); li.className = `opp-item ${statusClass}`;
         const statusBadge = status ? `<span class="opp-status-badge ${statusClass}">${status}</span>` : '';
         const typeLabel = oppType ? `<span class="opp-type">${oppType}</span>` : '';
         const roleLabel = role ? `<span class="opp-role-badge">${role}</span>` : '';
         li.innerHTML = `
           <span class="opp-title">${name}</span>
           <div class="opp-meta-row">${statusBadge}${typeLabel}${roleLabel}</div>
         `;
         li.onclick = function() { panelHistory = []; loadPanelRecord('Opportunities', opp.id); }; oppList.appendChild(li);
     });
  }
  function handleFormSubmit(formObject) {
    event.preventDefault();
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
  function loadPanelRecord(table, id) {
    const panel = document.getElementById('oppDetailPanel'); const content = document.getElementById('panelContent');
    const titleEl = document.getElementById('panelTitle'); const backBtn = document.getElementById('panelBackBtn');
    panel.classList.add('open'); content.innerHTML = `<div style="text-align:center; color:#999; margin-top:50px;">Loading...</div>`;
    google.script.run.withSuccessHandler(function(response) {
      if (!response || !response.data) { content.innerHTML = "Error loading."; return; }

      currentPanelData = {};
      response.data.forEach(item => { 
        if(item.type === 'link') {
          currentPanelData[item.key] = item.value;
        } else {
          // Store raw values for all fields (for appointment creation, etc.)
          currentPanelData[item.key] = item.value;
        }
      });

      panelHistory.push({ table: table, id: id, title: response.title });
      updateBackButton(); titleEl.innerText = response.title;
      
      // Helper to render a single field
      function renderField(item, tbl, recId) {
        const tacoClass = item.tacoField ? ' taco-field' : '';
        if (item.key === 'Opportunity Name') {
          const safeValue = (item.value || "").toString().replace(/"/g, "&quot;");
          return `<div class="detail-group${tacoClass}"><div class="detail-label">${item.label}</div><div id="view_${item.key}" onclick="toggleFieldEdit('${item.key}')" class="editable-field"><div class="detail-value" style="display:flex; justify-content:space-between; align-items:center;"><span id="display_${item.key}">${item.value || ''}</span><span class="edit-field-icon">✎</span></div></div><div id="edit_${item.key}" style="display:none;"><div class="edit-wrapper"><input type="text" id="input_${item.key}" value="${safeValue}" class="edit-input"><div class="edit-btn-row"><button onclick="cancelFieldEdit('${item.key}')" class="btn-cancel-field">Cancel</button><button id="btn_save_${item.key}" onclick="saveFieldEdit('${tbl}', '${recId}', '${item.key}')" class="btn-save-field">Save</button></div></div></div></div>`;
        }
        if (item.type === 'select') {
          const currentVal = item.value || '';
          const options = item.options || [];
          let optionsHtml = options.map(opt => `<option value="${opt}" ${opt === currentVal ? 'selected' : ''}>${opt}</option>`).join('');
          return `<div class="detail-group${tacoClass}"><div class="detail-label">${item.label}</div><div id="view_${item.key}" onclick="toggleFieldEdit('${item.key}')" class="editable-field"><div class="detail-value" style="display:flex; justify-content:space-between; align-items:center;"><span id="display_${item.key}">${currentVal || '<span style="color:#CCC; font-style:italic;">Not set</span>'}</span><span class="edit-field-icon">✎</span></div></div><div id="edit_${item.key}" style="display:none;"><div class="edit-wrapper"><select id="input_${item.key}" class="edit-input">${optionsHtml}</select><div class="edit-btn-row"><button onclick="cancelFieldEdit('${item.key}')" class="btn-cancel-field">Cancel</button><button id="btn_save_${item.key}" onclick="saveFieldEdit('${tbl}', '${recId}', '${item.key}')" class="btn-save-field">Save</button></div></div></div></div>`;
        }
        if (item.type === 'readonly') {
          const displayVal = item.value || '';
          if (!displayVal) return '';
          return `<div class="detail-group${tacoClass}"><div class="detail-label">${item.label}</div><div class="detail-value readonly-field">${displayVal}</div></div>`;
        }
        if (item.type === 'long-text') {
          const safeValue = (item.value || "").toString().replace(/"/g, "&quot;");
          const displayVal = item.value || '<span style="color:#CCC; font-style:italic;">Not set</span>';
          return `<div class="detail-group${tacoClass}"><div class="detail-label">${item.label}</div><div id="view_${item.key}" onclick="toggleFieldEdit('${item.key}')" class="editable-field"><div class="detail-value" style="display:flex; justify-content:space-between; align-items:flex-start;"><span id="display_${item.key}" style="white-space:pre-wrap; flex:1;">${displayVal}</span><span class="edit-field-icon" style="margin-left:8px;">✎</span></div></div><div id="edit_${item.key}" style="display:none;"><div class="edit-wrapper"><textarea id="input_${item.key}" class="edit-input" rows="3" style="resize:vertical;">${safeValue}</textarea><div class="edit-btn-row"><button onclick="cancelFieldEdit('${item.key}')" class="btn-cancel-field">Cancel</button><button id="btn_save_${item.key}" onclick="saveFieldEdit('${tbl}', '${recId}', '${item.key}')" class="btn-save-field">Save</button></div></div></div></div>`;
        }
        if (item.type === 'date') {
          const rawVal = item.value || '';
          let displayVal = '<span style="color:#CCC; font-style:italic;">Not set</span>';
          let inputVal = '';
          if (rawVal) {
            const parts = rawVal.split('/');
            if (parts.length === 3) { inputVal = `${parts[2].length === 2 ? '20' + parts[2] : parts[2]}-${parts[1]}-${parts[0]}`; displayVal = `${parts[0]}/${parts[1]}/${parts[2].slice(-2)}`; }
            else if (rawVal.includes('-')) { const isoParts = rawVal.split('-'); if (isoParts.length === 3) { inputVal = rawVal; displayVal = `${isoParts[2]}/${isoParts[1]}/${isoParts[0].slice(-2)}`; } }
            else { displayVal = rawVal; inputVal = rawVal; }
          }
          return `<div class="detail-group${tacoClass}"><div class="detail-label">${item.label}</div><div id="view_${item.key}" onclick="toggleFieldEdit('${item.key}')" class="editable-field"><div class="detail-value" style="display:flex; justify-content:space-between; align-items:center;"><span id="display_${item.key}">${displayVal}</span><span class="edit-field-icon">✎</span></div></div><div id="edit_${item.key}" style="display:none;"><div class="edit-wrapper"><input type="date" id="input_${item.key}" value="${inputVal}" class="edit-input"><div class="edit-btn-row"><button onclick="cancelFieldEdit('${item.key}')" class="btn-cancel-field">Cancel</button><button id="btn_save_${item.key}" onclick="saveDateField('${tbl}', '${recId}', '${item.key}')" class="btn-save-field">Save</button></div></div></div></div>`;
        }
        if (item.type === 'checkbox') {
          const isChecked = item.value === true || item.value === 'true' || item.value === 'Yes';
          const checkedAttr = isChecked ? 'checked' : '';
          return `<div class="detail-group${tacoClass}"><div class="checkbox-field"><input type="checkbox" id="input_${item.key}" ${checkedAttr} onchange="saveCheckboxField('${tbl}', '${recId}', '${item.key}', this.checked)"><label for="input_${item.key}">${item.label}</label></div></div>`;
        }
        if (item.type === 'url') {
          const safeValue = (item.value || "").toString().replace(/"/g, "&quot;");
          const displayVal = item.value ? `<a href="${item.value}" target="_blank" style="color:var(--color-sky);">${item.value}</a>` : '<span style="color:#CCC; font-style:italic;">Not set</span>';
          return `<div class="detail-group${tacoClass}"><div class="detail-label">${item.label}</div><div id="view_${item.key}" onclick="toggleFieldEdit('${item.key}')" class="editable-field"><div class="detail-value" style="display:flex; justify-content:space-between; align-items:center;"><span id="display_${item.key}">${displayVal}</span><span class="edit-field-icon">✎</span></div></div><div id="edit_${item.key}" style="display:none;"><div class="edit-wrapper"><input type="url" id="input_${item.key}" value="${safeValue}" class="edit-input" placeholder="https://..."><div class="edit-btn-row"><button onclick="cancelFieldEdit('${item.key}')" class="btn-cancel-field">Cancel</button><button id="btn_save_${item.key}" onclick="saveFieldEdit('${tbl}', '${recId}', '${item.key}')" class="btn-save-field">Save</button></div></div></div></div>`;
        }
        if (['Primary Applicant', 'Applicants', 'Guarantors'].includes(item.key)) {
          let linkHtml = '';
          if (item.value.length === 0) linkHtml = '<span style="color:#CCC; font-style:italic;">None</span>';
          else item.value.forEach(link => { linkHtml += `<span class="data-link panel-contact-link" data-contact-id="${link.id}" data-contact-table="${link.table}">${link.name}</span>`; });
          return `<div class="detail-group${tacoClass}"><div class="detail-label">${item.label}</div><div id="view_${item.key}" onclick="toggleLinkedEdit('${item.key}')" class="editable-field"><div class="detail-value" style="display:flex; justify-content:space-between; align-items:center;"><span>${linkHtml}</span><span class="edit-field-icon">✎</span></div></div><div id="edit_${item.key}" style="display:none;"><div class="edit-wrapper"><div id="chip_container_${item.key}" class="link-chip-container"></div><input type="text" placeholder="Add contact..." class="link-search-input" onkeyup="handleLinkedSearch(event, '${item.key}')"><div id="error_${item.key}" class="input-error"></div><div id="results_${item.key}" class="link-results"></div><div class="edit-btn-row" style="margin-top:10px;"><button onclick="cancelLinkedEdit('${item.key}')" class="btn-cancel-field">Cancel</button><button id="btn_save_${item.key}" onclick="saveLinkedEdit('${tbl}', '${recId}', '${item.key}')" class="btn-save-field">Save</button></div></div></div></div>`;
        }
        if (item.type === 'link') {
          const links = item.value; let linkHtml = '';
          if (links.length === 0) linkHtml = '<span style="color:#CCC; font-style:italic;">None</span>';
          else { links.forEach(link => { linkHtml += `<a class="data-link" onclick="loadPanelRecord('${link.table}', '${link.id}')">${link.name}</a>`; }); }
          return `<div class="detail-group${tacoClass}"><div class="detail-label">${item.label}</div><div class="detail-value" style="border:none;">${linkHtml}</div></div>`;
        }
        if (item.tacoField) {
          const safeValue = (item.value || "").toString().replace(/"/g, "&quot;");
          const displayVal = item.value || '<span style="color:#CCC; font-style:italic;">Not set</span>';
          return `<div class="detail-group${tacoClass}"><div class="detail-label">${item.label}</div><div id="view_${item.key}" onclick="toggleFieldEdit('${item.key}')" class="editable-field"><div class="detail-value" style="display:flex; justify-content:space-between; align-items:center;"><span id="display_${item.key}">${displayVal}</span><span class="edit-field-icon">✎</span></div></div><div id="edit_${item.key}" style="display:none;"><div class="edit-wrapper"><input type="text" id="input_${item.key}" value="${safeValue}" class="edit-input"><div class="edit-btn-row"><button onclick="cancelFieldEdit('${item.key}')" class="btn-cancel-field">Cancel</button><button id="btn_save_${item.key}" onclick="saveFieldEdit('${tbl}', '${recId}', '${item.key}')" class="btn-save-field">Save</button></div></div></div></div>`;
        }
        if (item.value === undefined || item.value === null || item.value === "") return '';
        return `<div class="detail-group${tacoClass}"><div class="detail-label">${item.label}</div><div class="detail-value">${item.value}</div></div>`;
      }
      
      let html = '';
      
      // For Opportunities, use smart layout
      if (table === 'Opportunities') {
        const dataMap = {};
        response.data.forEach(item => { dataMap[item.key] = item; });
        
        // Audit section with Evidence button on right
        if (response.audit && (response.audit.Created || response.audit.Modified)) {
          const oppName = (dataMap['Opportunity Name']?.value || response.title || '').replace(/'/g, "\\'");
          const oppType = (dataMap['Opportunity Type']?.value || '').replace(/'/g, "\\'");
          const lender = (dataMap['Lender']?.value || '').replace(/'/g, "\\'");
          html += `<div class="panel-audit-header">`;
          html += `<div class="panel-audit-section">`;
          if (response.audit.Created) html += `<div>${response.audit.Created}</div>`;
          if (response.audit.Modified) html += `<div>${response.audit.Modified}</div>`;
          html += `</div>`;
          html += `<button type="button" class="btn-evidence-top" onclick="openEvidenceModal('${id}', '${oppName}', '${oppType}', '${lender}')">📋 EVIDENCE & DATA COLLECTION</button>`;
          html += `</div>`;
        }
      } else {
        // Audit section for non-Opportunities
        if (response.audit && (response.audit.Created || response.audit.Modified)) {
          let auditHtml = '<div class="panel-audit-section">';
          if (response.audit.Created) auditHtml += `<div>${response.audit.Created}</div>`;
          if (response.audit.Modified) auditHtml += `<div>${response.audit.Modified}</div>`;
          auditHtml += '</div>';
          html += auditHtml;
        }
      }
      
      // Continue Opportunities layout
      if (table === 'Opportunities') {
        const dataMap = {};
        response.data.forEach(item => { dataMap[item.key] = item; });
        
        // Row 1: Opportunity Name, Status, Opportunity Type
        html += '<div class="panel-row panel-row-3">';
        ['Opportunity Name', 'Status', 'Opportunity Type'].forEach(key => {
          if (dataMap[key]) html += renderField(dataMap[key], table, id);
        });
        html += '</div>';
        
        // Taco fields section with custom layout
        const tacoFields = response.data.filter(item => item.tacoField);
        if (tacoFields.length > 0) {
          html += '<div class="taco-section-box">';
          html += `<div class="taco-section-header"><img src="https://taco.insightprocessing.com.au/static/images/taco.jpg" alt="Taco"><span>Fields from Taco Enquiry tab</span></div>`;
          html += '<div id="tacoFieldsContainer">';
          
          // Row 1: New or Existing Client, Lead Source (3rd empty)
          html += '<div class="taco-row">';
          if (dataMap['Taco: New or Existing Client']) html += renderField(dataMap['Taco: New or Existing Client'], table, id);
          if (dataMap['Taco: Lead Source']) html += renderField(dataMap['Taco: Lead Source'], table, id);
          html += '<div class="detail-group"></div>'; // empty 3rd column
          html += '</div>';
          
          // Row 2: Last thing we did, How can we help, CM notes
          html += '<div class="taco-row">';
          if (dataMap['Taco: Last thing we did']) html += renderField(dataMap['Taco: Last thing we did'], table, id);
          if (dataMap['Taco: How can we help']) html += renderField(dataMap['Taco: How can we help'], table, id);
          if (dataMap['Taco: CM notes']) html += renderField(dataMap['Taco: CM notes'], table, id);
          html += '</div>';
          
          // Row 3: Broker, Broker Assistant, Client Manager
          html += '<div class="taco-row">';
          if (dataMap['Taco: Broker']) html += renderField(dataMap['Taco: Broker'], table, id);
          if (dataMap['Taco: Broker Assistant']) html += renderField(dataMap['Taco: Broker Assistant'], table, id);
          if (dataMap['Taco: Client Manager']) html += renderField(dataMap['Taco: Client Manager'], table, id);
          html += '</div>';
          
          // Row 4: Converted to Appt (alone on left)
          html += '<div class="taco-row">';
          if (dataMap['Taco: Converted to Appt']) html += renderField(dataMap['Taco: Converted to Appt'], table, id);
          html += '</div>';
          
          html += '</div></div>'; // close tacoFieldsContainer and taco-section-box
        }
        
        // Appointments section - linked from Appointments table
        html += `<div class="appointments-section" style="margin-top:15px;">`;
        html += `<div id="appointmentsContainer" data-opportunity-id="${id}"><div style="color:#888; padding:10px;">Loading appointments...</div></div>`;
        html += `<div class="opp-action-buttons">`;
        html += `<button type="button" onclick="openAppointmentForm('${id}')">+ ADD APPOINTMENT</button>`;
        html += `</div></div>`;
        
        // Load appointments asynchronously
        setTimeout(() => loadAppointmentsForOpportunity(id), 100);
        
        // Row: Primary Applicant, Applicants, Guarantors, Loan Applications
        const applicantKeys = ['Primary Applicant', 'Applicants', 'Guarantors', 'Loan Applications'];
        html += '<div class="panel-row panel-row-4" style="margin-top:20px;">';
        applicantKeys.forEach(key => {
          if (dataMap[key]) html += renderField(dataMap[key], table, id);
        });
        html += '</div>';
        
        // Lead Source row
        if (dataMap['Lead Source Major'] || dataMap['Lead Source Minor']) {
          html += '<div class="panel-row panel-row-2">';
          if (dataMap['Lead Source Major']) html += renderField(dataMap['Lead Source Major'], table, id);
          if (dataMap['Lead Source Minor']) html += renderField(dataMap['Lead Source Minor'], table, id);
          html += '</div>';
        }
        
        // Remaining fields
        const usedKeys = new Set(['Opportunity Name', 'Status', 'Opportunity Type', 'Lead Source Major', 'Lead Source Minor', ...applicantKeys]);
        const remaining = response.data.filter(item => !item.tacoField && !usedKeys.has(item.key));
        if (remaining.length > 0) {
          html += '<div style="margin-top:15px; display:grid; grid-template-columns:repeat(3, 1fr); gap:12px 15px;">';
          remaining.forEach(item => { html += renderField(item, table, id); });
          html += '</div>';
        }
        
        // Delete button only (Send Confirmation moved to Taco section)
        const safeName = (response.title || '').replace(/'/g, "\\'");
        html += `<div style="margin-top:30px; padding-top:20px; border-top:1px solid #EEE;">`;
        html += `<button type="button" class="btn-delete btn-inline" onclick="confirmDeleteOpportunity('${id}', '${safeName}')">Delete Opportunity</button>`;
        html += `</div>`;
      } else {
        // Non-Opportunity tables: render sequentially
        response.data.forEach(item => { html += renderField(item, table, id); });
      }
      
      content.innerHTML = html;
    }).getRecordDetail(table, id);
  }
  function popHistory() { if (panelHistory.length <= 1) return; panelHistory.pop(); const prev = panelHistory[panelHistory.length - 1]; panelHistory.pop(); loadPanelRecord(prev.table, prev.id); }
  function updateBackButton() { const btn = document.getElementById('panelBackBtn'); if (panelHistory.length > 1) { btn.style.display = 'block'; } else { btn.style.display = 'none'; } }
  function closeOppPanel() { document.getElementById('oppDetailPanel').classList.remove('open'); panelHistory = []; }
  
  // --- APPOINTMENTS MANAGEMENT ---
  let currentAppointmentOpportunityId = null;
  let editingAppointmentId = null;
  
  // Helper function to render editable appointment fields (label above value like Taco fields)
  function renderApptField(apptId, label, fieldKey, value, type, options = []) {
    const displayValue = value || '-';
    let valueHtml = '';
    
    if (type === 'datetime') {
      const formatted = formatDatetimeForDisplay(value);
      valueHtml = `<div class="detail-value appt-editable" onclick="editApptField('${apptId}', '${fieldKey}', '${type}')" data-appt-id="${apptId}" data-field="${fieldKey}" data-value="${value || ''}" style="display:flex; justify-content:space-between; align-items:center;"><span>${formatted}</span><span class="edit-field-icon">✎</span></div>`;
    } else if (type === 'select') {
      valueHtml = `<div class="detail-value appt-editable" onclick="editApptField('${apptId}', '${fieldKey}', '${type}', ${JSON.stringify(options).replace(/"/g, '&quot;')})" data-appt-id="${apptId}" data-field="${fieldKey}" style="display:flex; justify-content:space-between; align-items:center;"><span>${displayValue}</span><span class="edit-field-icon">✎</span></div>`;
    } else if (type === 'textarea') {
      const escaped = (value || '').replace(/"/g, '&quot;');
      valueHtml = `<div class="detail-value appt-editable" onclick="editApptField('${apptId}', '${fieldKey}', '${type}')" data-appt-id="${apptId}" data-field="${fieldKey}" data-value="${escaped}" style="display:flex; justify-content:space-between; align-items:flex-start;"><span style="white-space:pre-wrap; flex:1;">${displayValue}</span><span class="edit-field-icon" style="margin-left:8px;">✎</span></div>`;
    } else {
      valueHtml = `<div class="detail-value appt-editable" onclick="editApptField('${apptId}', '${fieldKey}', '${type}')" data-appt-id="${apptId}" data-field="${fieldKey}" style="display:flex; justify-content:space-between; align-items:center;"><span>${displayValue}</span><span class="edit-field-icon">✎</span></div>`;
    }
    
    return `<div class="detail-group"><div class="detail-label">${label}</div>${valueHtml}</div>`;
  }
  
  // Helper function to render editable appointment fields without edit icon (for Notes/Video URL)
  function renderApptFieldNoIcon(apptId, label, fieldKey, value, type) {
    const displayValue = value || '';
    const escaped = (value || '').replace(/"/g, '&quot;');
    // For textarea fields like Notes, auto-size based on content (min 1 line)
    const isTextarea = type === 'textarea';
    const lineHeight = 20;
    const lines = displayValue ? displayValue.split('\n').length : 1;
    const minHeight = isTextarea ? `${Math.max(lineHeight, lines * lineHeight)}px` : 'auto';
    const style = isTextarea 
      ? `white-space:pre-wrap; min-height:${minHeight}; padding:8px; border:1px solid #ddd; border-radius:4px; cursor:text;`
      : `padding:8px; border:1px solid #ddd; border-radius:4px; cursor:text;`;
    const valueHtml = `<div class="detail-value appt-editable appt-notes-field" onclick="editApptField('${apptId}', '${fieldKey}', '${type}')" data-appt-id="${apptId}" data-field="${fieldKey}" data-value="${escaped}" style="${style}">${displayValue || '-'}</div>`;
    return `<div class="detail-group"><div class="detail-label">${label}</div>${valueHtml}</div>`;
  }
  
  // Helper function to render appointment checkboxes
  function renderApptCheckbox(apptId, label, fieldKey, checked) {
    return `<div class="detail-group"><div class="checkbox-field"><input type="checkbox" ${checked ? 'checked' : ''} onchange="updateApptCheckbox('${apptId}', '${fieldKey}', this.checked)"><label>${label}</label></div></div>`;
  }
  
  // Edit appointment field inline
  function editApptField(apptId, fieldKey, type, options) {
    const valueSpan = document.querySelector(`[data-appt-id="${apptId}"][data-field="${fieldKey}"]`);
    if (!valueSpan) return;
    
    const currentValue = valueSpan.dataset.value || valueSpan.textContent;
    const parent = valueSpan.parentElement;
    
    let inputHtml = '';
    if (type === 'datetime') {
      const dtValue = formatDatetimeForInput(currentValue);
      inputHtml = `<input type="datetime-local" class="inline-edit-input" value="${dtValue}" onblur="saveApptField('${apptId}', '${fieldKey}', this.value, '${type}')" onkeydown="if(event.key==='Enter'){this.blur();}if(event.key==='Escape'){cancelApptEdit('${apptId}', '${fieldKey}');}">`;
    } else if (type === 'select') {
      let optHtml = options.map(o => `<option value="${o}" ${o === currentValue ? 'selected' : ''}>${o}</option>`).join('');
      inputHtml = `<select class="inline-edit-input" onchange="saveApptField('${apptId}', '${fieldKey}', this.value, '${type}')" onblur="saveApptField('${apptId}', '${fieldKey}', this.value, '${type}')">${optHtml}</select>`;
    } else if (type === 'textarea') {
      inputHtml = `<textarea class="inline-edit-input auto-resize-textarea" rows="1" onblur="saveApptField('${apptId}', '${fieldKey}', this.value, '${type}')" oninput="autoResizeTextarea(this)" onfocus="autoResizeTextarea(this)" onkeydown="if(event.key==='Escape'){cancelApptEdit('${apptId}', '${fieldKey}');}">${currentValue === '-' ? '' : currentValue}</textarea>`;
    } else {
      inputHtml = `<input type="text" class="inline-edit-input" value="${currentValue === '-' ? '' : currentValue}" onblur="saveApptField('${apptId}', '${fieldKey}', this.value, '${type}')" onkeydown="if(event.key==='Enter'){this.blur();}if(event.key==='Escape'){cancelApptEdit('${apptId}', '${fieldKey}');}">`;
    }
    
    valueSpan.outerHTML = inputHtml;
    const input = parent.querySelector('.inline-edit-input');
    if (input) {
      input.focus();
      if (input.classList.contains('auto-resize-textarea')) {
        autoResizeTextarea(input);
      }
    }
  }
  
  // Save appointment field
  function saveApptField(apptId, fieldKey, value, type) {
    const opportunityId = document.getElementById('appointmentsContainer')?.dataset.opportunityId;
    
    // If setting appointment time to a future date and status is currently blank, auto-set to Scheduled
    if (fieldKey === 'appointmentTime' && value) {
      // Parse as local time (datetime-local gives YYYY-MM-DDTHH:MM format without timezone)
      const apptTime = new Date(value);
      const now = new Date();
      if (apptTime > now) {
        const statusEl = document.querySelector(`[data-appt-id="${apptId}"][data-field="appointmentStatus"]`);
        // Get current status from the displayed text
        const statusSpan = statusEl?.querySelector('span');
        const currentStatus = statusSpan?.textContent?.trim() || '';
        // Only auto-set if status is blank/not set (not if already has a real status)
        if (!currentStatus || currentStatus === 'Not Set' || currentStatus === '-' || currentStatus === '') {
          google.script.run.updateAppointment(apptId, 'appointmentStatus', 'Scheduled');
        }
      }
    }
    
    google.script.run
      .withSuccessHandler(function() {
        // Reload appointments to reflect changes
        if (opportunityId) loadAppointmentsForOpportunity(opportunityId);
      })
      .withFailureHandler(function(err) {
        console.error('Error updating appointment field:', err);
        alert('Error updating field: ' + (err.message || err));
        if (opportunityId) loadAppointmentsForOpportunity(opportunityId);
      })
      .updateAppointment(apptId, fieldKey, value);
  }
  
  // Update appointment checkbox
  function updateApptCheckbox(apptId, fieldKey, checked) {
    google.script.run
      .withSuccessHandler(function() {
        console.log('Appointment checkbox updated');
      })
      .withFailureHandler(function(err) {
        console.error('Error updating appointment checkbox:', err);
        alert('Error updating: ' + (err.message || err));
      })
      .updateAppointment(apptId, fieldKey, checked);
  }
  
  // Auto-resize textarea based on content
  function autoResizeTextarea(textarea) {
    textarea.style.height = 'auto';
    textarea.style.height = textarea.scrollHeight + 'px';
  }
  
  // ==================== NOTE POPOVER SYSTEM ====================
  
  let activeNotePopover = null;
  let noteSaveTimeout = null;
  
  /**
   * NOTE_FIELDS Configuration
   * Add new note fields here - all logic (popover, save, icons) is handled automatically.
   * 
   * Each entry requires:
   *   - fieldId: The HTML hidden field ID (used for form submission)
   *   - airtableField: The corresponding Airtable field name
   *   - inputId: The visible input field this note is attached to
   */
  const NOTE_FIELDS = [
    { fieldId: 'email1Comment', airtableField: 'EmailAddress1Comment', inputId: 'email1' },
    { fieldId: 'email2Comment', airtableField: 'EmailAddress2Comment', inputId: 'email2' },
    { fieldId: 'email3Comment', airtableField: 'EmailAddress3Comment', inputId: 'email3' },
    { fieldId: 'genderOther', airtableField: 'Gender - Other', inputId: 'gender' }
  ];

  // Build lookup map from config (for backward compatibility)
  const NOTE_FIELD_MAP = NOTE_FIELDS.reduce((map, f) => {
    map[f.fieldId] = f.airtableField;
    return map;
  }, {});

  /**
   * Initialize note icon on a field wrapper
   * Call this to attach note functionality to any .input-with-note wrapper
   */
  function initNoteIcon(wrapper, fieldId) {
    // Check if already initialized
    if (wrapper.querySelector('.note-icon')) return;
    
    // Create the icon button
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
    
    // Set initial state based on hidden field value
    const hiddenField = document.getElementById(fieldId);
    if (hiddenField) {
      updateNoteIconState(iconBtn, hiddenField.value);
    }
  }

  /**
   * Initialize all configured note fields
   * Called on page load to set up note icons on all configured fields
   */
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

  /**
   * Populate note field values from contact data
   * @param {Object} contact - The contact record from Airtable
   */
  function populateNoteFields(contact) {
    NOTE_FIELDS.forEach(config => {
      const field = document.getElementById(config.fieldId);
      if (field) {
        field.value = contact[config.airtableField] || '';
      }
    });
    updateAllNoteIcons();
  }

  /**
   * Get note field values for form submission
   * @returns {Object} Object with fieldId: value pairs
   */
  function getNoteFieldValues() {
    const values = {};
    NOTE_FIELDS.forEach(config => {
      const field = document.getElementById(config.fieldId);
      values[config.fieldId] = field ? field.value : '';
    });
    return values;
  }
  
  window.openNotePopover = function(iconBtn, fieldId) {
    // Close any existing popover
    closeNotePopover();
    
    const hiddenField = document.getElementById(fieldId);
    if (!hiddenField) return;
    
    const currentValue = hiddenField.value || '';
    const rect = iconBtn.getBoundingClientRect();
    const containerRect = iconBtn.closest('.field-with-note').getBoundingClientRect();
    
    // Create popover
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
    
    // Position the popover
    document.body.appendChild(popover);
    
    // Calculate position - below and to the left of the icon
    const popoverRect = popover.getBoundingClientRect();
    let top = rect.bottom + 5;
    let left = rect.right - popoverRect.width;
    
    // Keep within viewport
    if (left < 10) left = 10;
    if (top + popoverRect.height > window.innerHeight - 10) {
      top = rect.top - popoverRect.height - 5;
    }
    
    popover.style.position = 'fixed';
    popover.style.top = top + 'px';
    popover.style.left = left + 'px';
    
    // Store reference
    activeNotePopover = {
      element: popover,
      fieldId: fieldId,
      iconBtn: iconBtn,
      originalValue: currentValue
    };
    
    // Focus textarea
    const textarea = popover.querySelector('textarea');
    textarea.focus();
    textarea.setSelectionRange(textarea.value.length, textarea.value.length);
    
    // Auto-save on input (debounced)
    textarea.addEventListener('input', function() {
      if (noteSaveTimeout) clearTimeout(noteSaveTimeout);
      const status = document.getElementById('notePopoverStatus');
      status.textContent = '';
      status.className = 'note-popover-status';
      
      noteSaveTimeout = setTimeout(() => {
        saveNoteFromPopover();
      }, 800);
    });
    
    // Save on blur (if clicking outside)
    textarea.addEventListener('blur', function(e) {
      // Small delay to check if clicking on close button
      setTimeout(() => {
        if (activeNotePopover && !activeNotePopover.element.contains(document.activeElement)) {
          saveNoteFromPopover(true);
        }
      }, 100);
    });
    
    // Keyboard handling for textarea
    textarea.addEventListener('keydown', function(e) {
      if (e.key === 'Escape') {
        saveNoteFromPopover(true);
      }
    });
    
    // Done button handling
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
  };
  
  function saveNoteFromPopover(andClose = false) {
    if (!activeNotePopover) return;
    
    const textarea = document.getElementById('notePopoverTextarea');
    const status = document.getElementById('notePopoverStatus');
    const hiddenField = document.getElementById(activeNotePopover.fieldId);
    
    if (!textarea || !hiddenField) return;
    
    const newValue = textarea.value;
    const recordId = currentContactRecord?.id;
    const originalValue = activeNotePopover.originalValue;
    const airtableField = NOTE_FIELD_MAP[activeNotePopover.fieldId];
    const iconBtn = activeNotePopover.iconBtn;
    
    // Update hidden field immediately
    hiddenField.value = newValue;
    
    // Update icon state
    updateNoteIconState(iconBtn, newValue);
    
    // Close immediately if requested (user trusts it will save)
    if (andClose) {
      // Clear any pending debounce
      if (noteSaveTimeout) {
        clearTimeout(noteSaveTimeout);
        noteSaveTimeout = null;
      }
      closeNotePopover();
    }
    
    // Skip save if no record (new contact) or value unchanged
    if (!recordId || newValue === originalValue || !airtableField) {
      return;
    }
    
    // Show saving status (only if popover still open)
    if (status && !andClose) {
      status.textContent = 'Saving...';
      status.className = 'note-popover-status saving';
    }
    
    // Save to Airtable in background
    google.script.run
      .withSuccessHandler(function() {
        console.log('Note saved successfully');
      })
      .withFailureHandler(function(err) {
        console.error('Error saving note:', err);
      })
      .updateRecord('Contacts', recordId, airtableField, newValue);
  }
  
  window.closeNotePopover = function() {
    if (noteSaveTimeout) {
      clearTimeout(noteSaveTimeout);
      noteSaveTimeout = null;
    }
    
    if (activeNotePopover) {
      activeNotePopover.element.remove();
      activeNotePopover = null;
    }
  };
  
  // ==================== CONNECTION NOTE POPOVER ====================
  
  let activeConnNotePopover = null;
  let connNoteSaveTimeout = null;
  
  window.openConnectionNotePopover = function(iconBtn, connectionId, currentNote) {
    // Close any existing popover
    closeConnectionNotePopover();
    closeNotePopover();
    
    const rect = iconBtn.getBoundingClientRect();
    
    // Create popover
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
    
    // Position the popover
    document.body.appendChild(popover);
    
    // Calculate position - below and to the left of the icon
    const popoverRect = popover.getBoundingClientRect();
    let top = rect.bottom + 5;
    let left = rect.right - popoverRect.width;
    
    // Keep within viewport
    if (left < 10) left = 10;
    if (top + popoverRect.height > window.innerHeight - 10) {
      top = rect.top - popoverRect.height - 5;
    }
    
    popover.style.position = 'fixed';
    popover.style.top = top + 'px';
    popover.style.left = left + 'px';
    popover.style.zIndex = '10000';
    
    // Store state
    activeConnNotePopover = {
      element: popover,
      connectionId: connectionId,
      originalValue: currentNote || '',
      iconBtn: iconBtn
    };
    
    // Focus textarea
    const textarea = document.getElementById('connNotePopoverTextarea');
    textarea.focus();
    
    // Auto-save on input with debounce
    textarea.addEventListener('input', function() {
      if (connNoteSaveTimeout) clearTimeout(connNoteSaveTimeout);
      connNoteSaveTimeout = setTimeout(() => saveConnNoteFromPopover(false), 800);
    });
    
    // Done button handler
    const doneBtn = document.getElementById('connNotePopoverDone');
    doneBtn.addEventListener('click', function() {
      saveConnNoteFromPopover(true);
    });
    
    // Keyboard handling
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
  };
  
  function saveConnNoteFromPopover(andClose = false) {
    if (!activeConnNotePopover) return;
    
    const textarea = document.getElementById('connNotePopoverTextarea');
    const status = document.getElementById('connNotePopoverStatus');
    
    if (!textarea) return;
    
    const newValue = textarea.value;
    const connectionId = activeConnNotePopover.connectionId;
    const originalValue = activeConnNotePopover.originalValue;
    const iconBtn = activeConnNotePopover.iconBtn;
    
    // Update icon state
    if (iconBtn) {
      if (newValue && newValue.trim()) {
        iconBtn.classList.add('has-note');
      } else {
        iconBtn.classList.remove('has-note');
      }
    }
    
    // Update data attribute on parent element
    const parentEl = iconBtn?.closest('[data-conn-note]');
    if (parentEl) {
      parentEl.setAttribute('data-conn-note', newValue);
    }
    
    // Close immediately if requested
    if (andClose) {
      if (connNoteSaveTimeout) {
        clearTimeout(connNoteSaveTimeout);
        connNoteSaveTimeout = null;
      }
      closeConnectionNotePopover();
    }
    
    // Skip save if value unchanged
    if (newValue === originalValue) {
      return;
    }
    
    // Update original value to prevent duplicate saves
    if (activeConnNotePopover) {
      activeConnNotePopover.originalValue = newValue;
    }
    
    // Show saving status (only if popover still open)
    if (status && !andClose) {
      status.textContent = 'Saving...';
      status.className = 'note-popover-status saving';
    }
    
    // Save to Airtable
    google.script.run
      .withSuccessHandler(function() {
        console.log('Connection note saved successfully');
      })
      .withFailureHandler(function(err) {
        console.error('Error saving connection note:', err);
      })
      .updateConnectionNote(connectionId, newValue);
  }
  
  window.closeConnectionNotePopover = function() {
    if (connNoteSaveTimeout) {
      clearTimeout(connNoteSaveTimeout);
      connNoteSaveTimeout = null;
    }
    
    if (activeConnNotePopover) {
      activeConnNotePopover.element.remove();
      activeConnNotePopover = null;
    }
  };
  
  // Close connection note popover when clicking outside
  document.addEventListener('click', function(e) {
    if (activeConnNotePopover && !activeConnNotePopover.element.contains(e.target) && !e.target.classList.contains('conn-note-icon')) {
      // Don't close if user was selecting text
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
  
  function updateNoteIconState(iconBtn, value) {
    if (!iconBtn) return;
    if (value && value.trim()) {
      iconBtn.classList.add('has-note');
    } else {
      iconBtn.classList.remove('has-note');
    }
  }
  
  // Update all note icons when contact is loaded
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
  
  // Close popover when clicking outside (but not during text selection)
  document.addEventListener('click', function(e) {
    if (activeNotePopover && !activeNotePopover.element.contains(e.target) && !e.target.classList.contains('note-icon')) {
      // Don't close if user was selecting text (selection extends outside popover)
      const selection = window.getSelection();
      if (selection && selection.toString().length > 0) {
        // Check if selection started inside the popover textarea
        const textarea = document.getElementById('notePopoverTextarea');
        if (textarea && (document.activeElement === textarea || selection.anchorNode?.parentElement?.closest('.note-popover'))) {
          return; // Don't close - user is selecting text
        }
      }
      saveNoteFromPopover(true);
    }
  });
  
  // Cancel appointment edit
  function cancelApptEdit(apptId, fieldKey) {
    const opportunityId = document.getElementById('appointmentsContainer')?.dataset.opportunityId;
    if (opportunityId) loadAppointmentsForOpportunity(opportunityId);
  }
  
  function loadAppointmentsForOpportunity(opportunityId) {
    const container = document.getElementById('appointmentsContainer');
    if (!container) {
      console.log('Appointments container not found');
      return;
    }
    
    console.log('Loading appointments for opportunity:', opportunityId);
    
    google.script.run
      .withSuccessHandler(function(appointments) {
        console.log('Appointments received:', appointments);
        if (!appointments || appointments.length === 0) {
          // Check if this is a legacy record that needs backfill
          // Helper to detect truthy checkbox values (handles various Airtable serializations)
          const isTruthyCheckbox = (val) => {
            if (val === true || val === 1) return true;
            if (typeof val === 'string') {
              const lower = val.trim().toLowerCase();
              return ['true', 'yes', 'checked', '1'].includes(lower);
            }
            return Boolean(val);
          };
          
          // currentPanelData may contain objects like {value: true} or primitives
          const rawConverted = currentPanelData['Taco: Converted to Appt'];
          let convertedVal = (typeof rawConverted === 'object' && rawConverted !== null) 
            ? rawConverted.value 
            : rawConverted;
          
          if (isTruthyCheckbox(convertedVal)) {
            console.log('Legacy backfill: Converted to Appt is true but no appointments exist - creating from Taco fields');
            container.innerHTML = '<div style="color:#888; padding:16px 16px 4px 16px; font-style:italic;">Migrating appointment data...</div>';
            
            // Helper to extract primitive values from currentPanelData
            const getVal = (key) => {
              const v = currentPanelData[key];
              if (v === undefined || v === null) return null;
              if (typeof v === 'object' && v.value !== undefined) return v.value;
              return v;
            };
            const getBool = (key) => {
              const v = getVal(key);
              return isTruthyCheckbox(v);
            };
            
            // Build appointment fields from Taco data (server will parse the date)
            const fields = {
              "Appointment Time": getVal('Taco: Appointment Time') || null,
              "Type of Appointment": getVal('Taco: Type of Appointment') || "Phone",
              "How Booked": getVal('Taco: How appt booked') || "Calendly",
              "How Booked Other": getVal('Taco: How Appt Booked Other') || null,
              "Phone Number": getVal('Taco: Appt Phone Number') || null,
              "Video Meet URL": getVal('Taco: Appt Meet URL') || null,
              "Need Evidence in Advance": getBool('Taco: Need Evidence in Advance'),
              "Need Appt Reminder": getBool('Taco: Need Appt Reminder'),
              "Conf Email Sent": getBool('Taco: Appt Conf Email Sent'),
              "Conf Text Sent": getBool('Taco: Appt Conf Text Sent'),
              "Appointment Status": getVal('Taco: Appt Status') || null,
              "Notes": getVal('Taco: Appt Notes') || null
            };
            
            console.log('Backfill fields:', fields);
            
            google.script.run
              .withSuccessHandler(function() {
                console.log('Legacy appointment backfilled successfully');
                loadAppointmentsForOpportunity(opportunityId);
              })
              .withFailureHandler(function(err) {
                console.error('Failed to backfill legacy appointment:', err);
                container.innerHTML = '';
              })
              .createAppointment(opportunityId, fields);
            return;
          }
          
          container.innerHTML = '';
          window.currentOpportunityAppointments = [];
          return;
        }
        
        // Cache appointments for use in evidence email generation
        window.currentOpportunityAppointments = appointments;
        
        // Sort appointments oldest-first (ascending by date)
        appointments.sort((a, b) => {
          const dateA = a.appointmentTime ? new Date(a.appointmentTime).getTime() : 0;
          const dateB = b.appointmentTime ? new Date(b.appointmentTime).getTime() : 0;
          return dateA - dateB;
        });
        
        let html = '';
        appointments.forEach(appt => {
          const status = appt.appointmentStatus || '';
          const isPast = appt.appointmentTime && new Date(appt.appointmentTime) < new Date();
          const needsUpdate = isPast && (status === 'Scheduled' || status === '');
          const statusClass = needsUpdate ? 'status-needs-update' :
                             status === 'Completed' ? 'status-completed' : 
                             status === 'Cancelled' ? 'status-cancelled' : 
                             status === 'No Show' ? 'status-noshow' : 
                             status === 'Scheduled' ? 'status-scheduled' : 'status-blank';
          const statusDisplay = needsUpdate ? 'Please Update Status' : (status || 'Not Set');
          
          // Expand Scheduled appointments by default, collapse others
          const isExpanded = status === 'Scheduled';
          const expandedClass = isExpanded ? 'expanded' : '';
          
          html += `<div class="appointment-item subsequent-appt ${expandedClass}" data-appt-id="${appt.id}">`;
          
          // Collapsible header - nice text flow with status badge on right
          html += `<div class="appointment-item-header" onclick="toggleAppointmentExpand('${appt.id}')">`;
          html += `<div class="appt-header-left">`;
          html += `<span class="appointment-item-chevron">▶</span>`;
          html += `<span class="appt-header-label">Appointment:</span>`;
          html += `<span class="appt-header-time">${formatDatetimeForDisplay(appt.appointmentTime)}</span>`;
          html += `<span class="appt-header-type">${appt.typeOfAppointment || '-'}</span>`;
          html += `</div>`;
          const statusTooltip = needsUpdate ? ' title="This appointment time has passed but the status is still Scheduled or blank. Please update to Completed, Cancelled, or No Show."' : '';
          html += `<span class="appointment-status ${statusClass}"${statusTooltip}>${statusDisplay}</span>`;
          html += `</div>`;
          
          // Expandable body with editable fields
          html += `<div class="appointment-item-body">`;
          html += `<div class="appointment-item-divider"></div>`;
          
          // Audit info above both sections
          let auditParts = [];
          if (appt.createdTime) {
            const createdDate = new Date(appt.createdTime).toLocaleString('en-AU', { day: '2-digit', month: '2-digit', year: '2-digit', hour: 'numeric', minute: '2-digit', hour12: true });
            let createdText = `Created ${createdDate}`;
            if (appt.createdByName) createdText += ` by ${appt.createdByName}`;
            auditParts.push(createdText);
          }
          if (appt.modifiedTime && appt.modifiedTime !== appt.createdTime) {
            const modifiedDate = new Date(appt.modifiedTime).toLocaleString('en-AU', { day: '2-digit', month: '2-digit', year: '2-digit', hour: 'numeric', minute: '2-digit', hour12: true });
            let modifiedText = `Modified ${modifiedDate}`;
            if (appt.modifiedByName) modifiedText += ` by ${appt.modifiedByName}`;
            auditParts.push(modifiedText);
          }
          if (auditParts.length > 0) {
            html += `<div class="appt-audit-info">${auditParts.join(' · ')}</div>`;
          }
          
          // Section 1: Appointment details and preparation
          html += `<div class="appt-section appt-section-1">`;
          
          // Row 1: Appointment Time, Type of Appointment, How Booked (editable)
          html += `<div class="taco-row">`;
          html += renderApptField(appt.id, 'Appointment Time', 'appointmentTime', appt.appointmentTime, 'datetime');
          html += renderApptField(appt.id, 'Type of Appointment', 'typeOfAppointment', appt.typeOfAppointment, 'select', ['Phone', 'Video', 'Office']);
          html += renderApptField(appt.id, 'How Booked', 'howBooked', appt.howBooked, 'select', ['Calendly', 'Email', 'Phone', 'Podium', 'Other']);
          html += `</div>`;
          
          // Row 2: Phone Number (if Phone), Video Meet URL (if Video), How Booked Other (if Other)
          html += `<div class="taco-row">`;
          const phoneStyle = appt.typeOfAppointment === 'Phone' ? '' : 'display:none;';
          const videoStyle = appt.typeOfAppointment === 'Video' ? '' : 'display:none;';
          const otherStyle = appt.howBooked === 'Other' ? '' : 'display:none;';
          html += `<div id="appt_field_wrap_${appt.id}_phone" style="${phoneStyle}">${renderApptFieldNoIcon(appt.id, 'Phone Number', 'phoneNumber', appt.phoneNumber, 'text')}</div>`;
          html += `<div id="appt_field_wrap_${appt.id}_video" style="${videoStyle}">${renderApptFieldNoIcon(appt.id, 'Video Meet URL', 'videoMeetUrl', appt.videoMeetUrl, 'text')}</div>`;
          html += `<div id="appt_field_wrap_${appt.id}_other" style="${otherStyle}">${renderApptField(appt.id, 'How Booked Other', 'howBookedOther', appt.howBookedOther, 'text')}</div>`;
          html += `</div>`;
          
          // Row 3: Need Evidence, Need Reminder
          html += `<div class="taco-row">`;
          html += renderApptCheckbox(appt.id, 'Need Evidence in Advance', 'needEvidenceInAdvance', appt.needEvidenceInAdvance);
          html += renderApptCheckbox(appt.id, 'Need Appt Reminder', 'needApptReminder', appt.needApptReminder);
          html += `</div>`;
          
          // Send Confirmation Email button
          html += `<div style="margin:15px 0;"><button type="button" class="btn-confirm btn-inline" onclick="openEmailComposerFromPanel('${opportunityId}')">Send Confirmation Email</button></div>`;
          
          html += `</div>`; // close section 1
          
          // Section 2: Confirmation status and outcome
          html += `<div class="appt-section appt-section-2">`;
          
          // Row 4: Conf Email Sent, Conf Text Sent, Appointment Status
          html += `<div class="taco-row">`;
          html += renderApptCheckbox(appt.id, 'Conf Email Sent', 'confEmailSent', appt.confEmailSent);
          html += renderApptCheckbox(appt.id, 'Conf Text Sent', 'confTextSent', appt.confTextSent);
          html += renderApptField(appt.id, 'Appointment Status', 'appointmentStatus', status, 'select', ['', 'Scheduled', 'Completed', 'Cancelled', 'No Show']);
          html += `</div>`;
          
          // Row 5: Notes (full width, auto-resize)
          html += `<div style="margin-top:12px;">`;
          html += renderApptFieldNoIcon(appt.id, 'Notes', 'notes', appt.notes, 'textarea');
          html += `</div>`;
          
          html += `</div>`; // close section 2
          html += `</div>`; // close body
          html += `</div>`; // close item
        });
        
        container.innerHTML = html;
      })
      .withFailureHandler(function(err) {
        console.error('Error loading appointments:', err);
        container.innerHTML = '<div style="color:#C00; padding:10px;">Error loading appointments: ' + (err.message || err) + '</div>';
      })
      .getAppointmentsForOpportunity(opportunityId);
  }
  
  function formatDatetimeForInput(dateStr) {
    if (!dateStr) return '';
    try {
      const d = new Date(dateStr);
      if (isNaN(d.getTime())) return '';
      return d.toISOString().slice(0, 16);
    } catch (e) {
      return '';
    }
  }
  
  function formatDatetimeForDisplay(dateStr) {
    if (!dateStr) return 'Time not set';
    try {
      const d = new Date(dateStr);
      if (isNaN(d.getTime())) return dateStr;
      return d.toLocaleString('en-AU', { 
        weekday: 'short', 
        day: '2-digit', 
        month: '2-digit', 
        year: '2-digit',
        hour: 'numeric',
        minute: '2-digit',
        hour12: true 
      });
    } catch (e) {
      return dateStr;
    }
  }
  
  function openAppointmentForm(opportunityId, appointment = null) {
    currentAppointmentOpportunityId = opportunityId;
    editingAppointmentId = appointment ? appointment.id : null;
    
    const modal = document.getElementById('appointmentFormModal');
    const title = document.getElementById('appointmentFormTitle');
    title.textContent = appointment ? 'Edit Appointment' : 'New Appointment';
    
    // Reset form - format datetime for datetime-local input
    document.getElementById('apptFormTime').value = formatDatetimeForInput(appointment?.appointmentTime);
    document.getElementById('apptFormType').value = appointment?.typeOfAppointment || 'Phone';
    document.getElementById('apptFormHowBooked').value = appointment?.howBooked || 'Calendly';
    document.getElementById('apptFormHowBookedOther').value = appointment?.howBookedOther || '';
    document.getElementById('apptFormPhone').value = appointment?.phoneNumber || '';
    document.getElementById('apptFormMeetUrl').value = appointment?.videoMeetUrl || '';
    document.getElementById('apptFormNeedEvidence').checked = appointment?.needEvidenceInAdvance || false;
    document.getElementById('apptFormNeedReminder').checked = appointment?.needApptReminder || false;
    document.getElementById('apptFormNotes').value = appointment?.notes || '';
    document.getElementById('apptFormStatus').value = appointment?.appointmentStatus || '';
    
    updateAppointmentFormVisibility();
    modal.classList.add('visible');
    setTimeout(() => modal.classList.add('showing'), 10);
    
    // Scroll modal to top and focus first field
    modal.scrollTop = 0;
    const modalContent = modal.querySelector('.modal-content');
    if (modalContent) modalContent.scrollTop = 0;
    setTimeout(() => {
      const firstInput = document.getElementById('apptFormTime');
      if (firstInput) firstInput.focus();
    }, 100);
  }
  
  function updateAppointmentFormVisibility() {
    const type = document.getElementById('apptFormType').value;
    const howBooked = document.getElementById('apptFormHowBooked').value;
    
    document.getElementById('apptFormPhoneRow').style.display = type === 'Phone' ? 'block' : 'none';
    document.getElementById('apptFormMeetRow').style.display = type === 'Video' ? 'block' : 'none';
    document.getElementById('apptFormHowBookedOtherRow').style.display = howBooked === 'Other' ? 'block' : 'none';
  }
  
  function closeAppointmentForm() {
    const modal = document.getElementById('appointmentFormModal');
    modal.classList.remove('showing');
    setTimeout(() => modal.classList.remove('visible'), 200);
    currentAppointmentOpportunityId = null;
    editingAppointmentId = null;
  }
  
  function saveAppointment() {
    if (!currentAppointmentOpportunityId) return;
    
    const fields = {
      "Appointment Time": document.getElementById('apptFormTime').value,
      "Type of Appointment": document.getElementById('apptFormType').value,
      "How Booked": document.getElementById('apptFormHowBooked').value,
      "How Booked Other": document.getElementById('apptFormHowBookedOther').value,
      "Phone Number": document.getElementById('apptFormPhone').value,
      "Video Meet URL": document.getElementById('apptFormMeetUrl').value,
      "Need Evidence in Advance": document.getElementById('apptFormNeedEvidence').checked,
      "Need Appt Reminder": document.getElementById('apptFormNeedReminder').checked,
      "Notes": document.getElementById('apptFormNotes').value,
      "Appointment Status": document.getElementById('apptFormStatus').value
    };
    
    const saveBtn = document.getElementById('apptFormSaveBtn');
    saveBtn.disabled = true;
    saveBtn.textContent = 'Saving...';
    
    const oppId = currentAppointmentOpportunityId;
    
    function onSaveComplete() {
      saveBtn.disabled = false;
      saveBtn.textContent = 'Save';
      closeAppointmentForm();
      loadAppointmentsForOpportunity(oppId);
    }
    
    function onSaveError(err) {
      console.error('Error saving appointment:', err);
      alert('Error saving appointment: ' + (err.message || err));
      saveBtn.disabled = false;
      saveBtn.textContent = 'Save';
    }
    
    if (editingAppointmentId) {
      // For updates, we need to update each field sequentially or as a batch
      // Using a single update with all fields would be cleaner
      google.script.run
        .withSuccessHandler(onSaveComplete)
        .withFailureHandler(onSaveError)
        .updateAppointmentFields(editingAppointmentId, fields);
    } else {
      google.script.run
        .withSuccessHandler(onSaveComplete)
        .withFailureHandler(onSaveError)
        .createAppointment(currentAppointmentOpportunityId, fields);
    }
  }
  
  function toggleAppointmentExpand(appointmentId) {
    const item = document.querySelector(`.appointment-item[data-appt-id="${appointmentId}"]`);
    if (item) {
      item.classList.toggle('expanded');
    }
  }
  
  
  function editAppointment(appointmentId, opportunityId) {
    google.script.run
      .withSuccessHandler(function(appointments) {
        const appt = appointments.find(a => a.id === appointmentId);
        if (appt) {
          openAppointmentForm(opportunityId, appt);
        }
      })
      .withFailureHandler(function(err) {
        console.error('Error loading appointment for edit:', err);
      })
      .getAppointmentsForOpportunity(opportunityId);
  }
  
  function deleteAppointment(appointmentId, opportunityId) {
    if (!confirm('Are you sure you want to delete this appointment?')) return;
    
    google.script.run
      .withSuccessHandler(function() {
        loadAppointmentsForOpportunity(opportunityId);
      })
      .withFailureHandler(function(err) {
        console.error('Error deleting appointment:', err);
        alert('Error deleting appointment: ' + (err.message || err));
      })
      .deleteAppointment(appointmentId);
  }

  // ==================== ADDRESS HISTORY SYSTEM ====================
  let currentContactAddresses = [];
  let editingAddressId = null;
  
  function loadAddressHistory(contactId) {
    const container = document.getElementById('addressHistoryList');
    if (!container || !contactId) return;
    
    container.innerHTML = '<div style="padding:10px; color:#888; font-size:12px;">Loading addresses...</div>';
    
    google.script.run
      .withSuccessHandler(function(addresses) {
        currentContactAddresses = addresses || [];
        renderAddressHistory();
      })
      .withFailureHandler(function(err) {
        console.error('Error loading addresses:', err);
        container.innerHTML = '<div style="padding:10px; color:#A00; font-size:12px;">Error loading addresses</div>';
      })
      .getAddressesForContact(contactId);
  }
  
  function renderAddressHistory() {
    const container = document.getElementById('addressHistoryList');
    if (!container) return;
    
    if (currentContactAddresses.length === 0) {
      container.innerHTML = '<div style="padding:10px; color:#888; font-size:12px; font-style:italic;">No addresses recorded</div>';
      return;
    }
    
    // Separate residential and postal addresses
    const residential = currentContactAddresses.filter(a => !a.isPostal);
    const postal = currentContactAddresses.filter(a => a.isPostal);
    
    let html = '';
    
    // Render residential addresses
    residential.forEach((addr, idx) => {
      const isCurrent = !addr.to;
      const dateRange = formatAddressDateRange(addr.from, addr.to);
      html += `
        <div class="address-item${isCurrent ? ' is-current' : ''}" onclick="editAddress('${addr.id}')">
          <div class="address-name">${escapeHtml(addr.calculatedName) || 'No address'}</div>
          <div class="address-meta-row">
            ${addr.status ? `<span class="address-status-badge">${escapeHtml(addr.status)}</span>` : ''}
            <span class="address-date-range">${dateRange}</span>
          </div>
        </div>
      `;
    });
    
    // Render postal addresses
    postal.forEach(addr => {
      html += `
        <div class="address-item is-postal" onclick="editAddress('${addr.id}')">
          <div class="address-name">${escapeHtml(addr.calculatedName) || 'No address'}</div>
          <div class="address-meta-row">
            <span class="address-postal-badge">POSTAL</span>
            ${addr.status ? `<span class="address-status-badge">${escapeHtml(addr.status)}</span>` : ''}
          </div>
        </div>
      `;
    });
    
    // Add expand/collapse if more than 2 addresses
    const totalAddresses = currentContactAddresses.length;
    if (totalAddresses > 2) {
      const isExpanded = container.classList.contains('expanded');
      html += `<span class="address-expand-link" onclick="toggleAddressExpand(event)">${isExpanded ? 'Show less' : `Show all ${totalAddresses} addresses`}</span>`;
    }
    
    container.innerHTML = html;
  }
  
  function formatAddressDateRange(from, to) {
    const fromStr = from ? formatDateDisplay(from) : '?';
    const toStr = to ? formatDateDisplay(to) : 'Present';
    return `${fromStr} - ${toStr}`;
  }
  
  function formatDateDisplay(isoDate) {
    if (!isoDate) return '';
    const parts = isoDate.split('-');
    if (parts.length !== 3) return isoDate;
    return `${parts[2]}/${parts[1]}/${parts[0]}`;
  }
  
  function parseDateInput(value) {
    // Convert DD/MM/YYYY to YYYY-MM-DD for Airtable
    if (!value) return null;
    const match = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (match) {
      return `${match[3]}-${match[2].padStart(2, '0')}-${match[1].padStart(2, '0')}`;
    }
    return value; // Return as-is if already ISO format
  }
  
  window.toggleAddressExpand = function(event) {
    event.stopPropagation();
    const container = document.getElementById('addressHistoryList');
    container.classList.toggle('expanded');
    renderAddressHistory();
  };
  
  window.openAddressModal = function(isPostal = false) {
    const recordId = currentContactRecord?.id;
    if (!recordId) {
      alert('Please save the contact first');
      return;
    }
    
    editingAddressId = null;
    document.getElementById('addressFormId').value = '';
    document.getElementById('addressFormIsPostal').value = isPostal ? 'true' : 'false';
    document.getElementById('addressFormTitle').textContent = isPostal ? 'Add Postal Address' : 'Add Address';
    document.getElementById('addressDeleteBtn').style.display = 'none';
    
    // Reset form fields
    document.querySelector('input[name="addressFormat"][value="Standard"]').checked = true;
    document.getElementById('addressFloor').value = '';
    document.getElementById('addressBuilding').value = '';
    document.getElementById('addressUnit').value = '';
    document.getElementById('addressStreetNo').value = '';
    document.getElementById('addressStreetName').value = '';
    document.getElementById('addressStreetType').value = '';
    document.getElementById('addressCity').value = '';
    document.getElementById('addressState').value = '';
    document.getElementById('addressPostcode').value = '';
    document.getElementById('addressCountry').value = 'Australia';
    document.getElementById('addressLabel').value = '';
    document.getElementById('addressStatus').value = '';
    document.getElementById('addressFrom').value = '';
    document.getElementById('addressTo').value = '';
    
    updateAddressFormatFields();
    
    const modal = document.getElementById('addressFormModal');
    modal.style.display = 'flex';
    setTimeout(() => modal.classList.add('showing'), 10);
  };
  
  window.openPostalAddressModal = function() {
    // Check if there are existing residential addresses to copy from
    const residential = currentContactAddresses.filter(a => !a.isPostal);
    
    if (residential.length === 0) {
      // No addresses to copy from, open new postal address form directly
      openAddressModal(true);
      return;
    }
    
    // Show copy selection modal
    const listContainer = document.getElementById('postalAddressCopyList');
    let html = '';
    residential.forEach(addr => {
      html += `
        <div class="postal-copy-item" onclick="copyAddressAsPostal('${addr.id}')">
          ${escapeHtml(addr.calculatedName) || 'No address'}
        </div>
      `;
    });
    listContainer.innerHTML = html;
    
    const modal = document.getElementById('postalAddressCopyModal');
    modal.style.display = 'flex';
    setTimeout(() => modal.classList.add('showing'), 10);
  };
  
  window.closePostalCopyModal = function() {
    const modal = document.getElementById('postalAddressCopyModal');
    modal.classList.remove('showing');
    setTimeout(() => modal.style.display = 'none', 250);
  };
  
  window.openPostalAddressNew = function() {
    closePostalCopyModal();
    openAddressModal(true);
  };
  
  window.copyAddressAsPostal = function(addressId) {
    closePostalCopyModal();
    
    const addr = currentContactAddresses.find(a => a.id === addressId);
    if (!addr) {
      openAddressModal(true);
      return;
    }
    
    // Open modal pre-filled with copied address data
    editingAddressId = null;
    document.getElementById('addressFormId').value = '';
    document.getElementById('addressFormIsPostal').value = 'true';
    document.getElementById('addressFormTitle').textContent = 'Add Postal Address';
    document.getElementById('addressDeleteBtn').style.display = 'none';
    
    // Fill form with copied data
    const formatRadio = document.querySelector(`input[name="addressFormat"][value="${addr.format || 'Standard'}"]`);
    if (formatRadio) formatRadio.checked = true;
    
    document.getElementById('addressFloor').value = addr.floor || '';
    document.getElementById('addressBuilding').value = addr.building || '';
    document.getElementById('addressUnit').value = addr.unit || '';
    document.getElementById('addressStreetNo').value = addr.streetNo || '';
    document.getElementById('addressStreetName').value = addr.streetName || '';
    document.getElementById('addressStreetType').value = addr.streetType || '';
    document.getElementById('addressCity').value = addr.city || '';
    document.getElementById('addressState').value = addr.state || '';
    document.getElementById('addressPostcode').value = addr.postcode || '';
    document.getElementById('addressCountry').value = addr.country || 'Australia';
    document.getElementById('addressLabel').value = addr.label || '';
    document.getElementById('addressStatus').value = ''; // Don't copy status
    document.getElementById('addressFrom').value = ''; // Don't copy dates
    document.getElementById('addressTo').value = '';
    
    updateAddressFormatFields();
    
    const modal = document.getElementById('addressFormModal');
    modal.style.display = 'flex';
    setTimeout(() => modal.classList.add('showing'), 10);
  };
  
  window.editAddress = function(addressId) {
    const addr = currentContactAddresses.find(a => a.id === addressId);
    if (!addr) return;
    
    editingAddressId = addressId;
    document.getElementById('addressFormId').value = addressId;
    document.getElementById('addressFormIsPostal').value = addr.isPostal ? 'true' : 'false';
    document.getElementById('addressFormTitle').textContent = 'Edit Address';
    document.getElementById('addressDeleteBtn').style.display = 'block';
    
    // Fill form
    const formatRadio = document.querySelector(`input[name="addressFormat"][value="${addr.format || 'Standard'}"]`);
    if (formatRadio) formatRadio.checked = true;
    
    document.getElementById('addressFloor').value = addr.floor || '';
    document.getElementById('addressBuilding').value = addr.building || '';
    document.getElementById('addressUnit').value = addr.unit || '';
    document.getElementById('addressStreetNo').value = addr.streetNo || '';
    document.getElementById('addressStreetName').value = addr.streetName || '';
    document.getElementById('addressStreetType').value = addr.streetType || '';
    document.getElementById('addressCity').value = addr.city || '';
    document.getElementById('addressState').value = addr.state || '';
    document.getElementById('addressPostcode').value = addr.postcode || '';
    document.getElementById('addressCountry').value = addr.country || 'Australia';
    document.getElementById('addressLabel').value = addr.label || '';
    document.getElementById('addressStatus').value = addr.status || '';
    document.getElementById('addressFrom').value = addr.from ? formatDateDisplay(addr.from) : '';
    document.getElementById('addressTo').value = addr.to ? formatDateDisplay(addr.to) : '';
    
    updateAddressFormatFields();
    
    const modal = document.getElementById('addressFormModal');
    modal.style.display = 'flex';
    setTimeout(() => modal.classList.add('showing'), 10);
  };
  
  window.closeAddressForm = function() {
    const modal = document.getElementById('addressFormModal');
    modal.classList.remove('showing');
    setTimeout(() => modal.style.display = 'none', 250);
    editingAddressId = null;
  };
  
  window.updateAddressFormatFields = function() {
    const format = document.querySelector('input[name="addressFormat"]:checked')?.value || 'Standard';
    
    document.getElementById('addressNonStandardFields').style.display = format === 'Non-Standard' ? 'block' : 'none';
    document.getElementById('addressPOBoxFields').style.display = format === 'PO Box' ? 'block' : 'none';
    document.getElementById('addressStreetFields').style.display = format === 'PO Box' ? 'none' : 'block';
  };
  
  window.saveAddress = function() {
    const recordId = currentContactRecord?.id;
    if (!recordId) {
      alert('Please save the contact first');
      return;
    }
    
    const format = document.querySelector('input[name="addressFormat"]:checked')?.value || 'Standard';
    const isPostal = document.getElementById('addressFormIsPostal').value === 'true';
    
    const fields = {
      format: format,
      floor: document.getElementById('addressFloor').value,
      building: document.getElementById('addressBuilding').value,
      unit: document.getElementById('addressUnit').value,
      streetNo: document.getElementById('addressStreetNo').value,
      streetName: document.getElementById('addressStreetName').value,
      streetType: document.getElementById('addressStreetType').value,
      city: document.getElementById('addressCity').value,
      state: document.getElementById('addressState').value,
      postcode: document.getElementById('addressPostcode').value,
      country: document.getElementById('addressCountry').value,
      label: document.getElementById('addressLabel').value,
      status: document.getElementById('addressStatus').value,
      from: parseDateInput(document.getElementById('addressFrom').value),
      to: parseDateInput(document.getElementById('addressTo').value),
      isPostal: isPostal
    };
    
    if (editingAddressId) {
      // Update existing address
      google.script.run
        .withSuccessHandler(function(result) {
          if (result.success) {
            closeAddressForm();
            loadAddressHistory(recordId);
          } else {
            alert('Error saving address: ' + (result.error || 'Unknown error'));
          }
        })
        .withFailureHandler(function(err) {
          console.error('Error updating address:', err);
          alert('Error saving address: ' + (err.message || err));
        })
        .updateAddress(editingAddressId, fields);
    } else {
      // Create new address
      google.script.run
        .withSuccessHandler(function(result) {
          if (result.success) {
            closeAddressForm();
            loadAddressHistory(recordId);
          } else {
            alert('Error creating address: ' + (result.error || 'Unknown error'));
          }
        })
        .withFailureHandler(function(err) {
          console.error('Error creating address:', err);
          alert('Error creating address: ' + (err.message || err));
        })
        .createAddress(recordId, fields);
    }
  };
  
  window.deleteAddress = function() {
    if (!editingAddressId) return;
    
    const recordId = currentContactRecord?.id;
    const addressId = editingAddressId;
    
    showConfirmModal('Are you sure you want to delete this address?', function() {
      google.script.run
        .withSuccessHandler(function(result) {
          if (result.success) {
            closeAddressForm();
            loadAddressHistory(recordId);
          } else {
            alert('Error deleting address: ' + (result.error || 'Unknown error'));
          }
        })
        .withFailureHandler(function(err) {
          console.error('Error deleting address:', err);
          alert('Error deleting address: ' + (err.message || err));
        })
        .deleteAddress(addressId);
    });
  };

  // ==================== EVIDENCE MODAL SYSTEM ====================
  let currentEvidenceOpportunityId = null;
  let currentEvidenceOpportunityName = '';
  let currentEvidenceOpportunityType = '';
  let currentEvidenceLender = '';
  let currentEvidenceItems = [];

  window.openEvidenceModal = function(opportunityId, opportunityName, opportunityType, lender) {
    currentEvidenceOpportunityId = opportunityId;
    currentEvidenceOpportunityName = opportunityName || 'Opportunity';
    currentEvidenceOpportunityType = opportunityType || '';
    currentEvidenceLender = lender || '';
    
    document.getElementById('evidenceOppName').textContent = currentEvidenceOpportunityName + (lender ? ` - ${lender}` : '');
    
    const modal = document.getElementById('evidenceModal');
    modal.classList.add('visible');
    setTimeout(() => modal.classList.add('showing'), 10);
    
    loadEvidenceItems();
  };

  window.closeEvidenceModal = function() {
    const modal = document.getElementById('evidenceModal');
    modal.classList.remove('showing');
    setTimeout(() => modal.classList.remove('visible'), 300);
    currentEvidenceOpportunityId = null;
    currentEvidenceItems = [];
  };

  function loadEvidenceItems() {
    const loading = document.getElementById('evidenceLoading');
    const emptyState = document.getElementById('evidenceEmptyState');
    const container = document.getElementById('evidenceItemsContainer');
    
    loading.style.display = 'block';
    emptyState.style.display = 'none';
    container.innerHTML = '';
    
    google.script.run
      .withSuccessHandler(function(items) {
        loading.style.display = 'none';
        currentEvidenceItems = items || [];
        
        if (currentEvidenceItems.length === 0) {
          emptyState.style.display = 'block';
        } else {
          renderEvidenceItems();
        }
      })
      .withFailureHandler(function(err) {
        loading.style.display = 'none';
        console.error('Error loading evidence items:', err);
        container.innerHTML = '<p style="color:#A00; text-align:center;">Error loading evidence items</p>';
      })
      .getEvidenceItemsForOpportunity(currentEvidenceOpportunityId);
  }

  function renderEvidenceItems() {
    const container = document.getElementById('evidenceItemsContainer');
    const filter = document.getElementById('evidenceStatusFilter').value;
    const showNA = document.getElementById('evidenceShowNA').checked;
    const outstandingFirst = document.getElementById('evidenceOutstandingFirst').checked;
    
    // Sort items if "Outstanding First" is checked
    let itemsToRender = [...currentEvidenceItems];
    if (outstandingFirst) {
      // Outstanding (and any other status) = 1, Received = 2, N/A = 3
      const getPriority = (status) => status === 'Received' ? 2 : status === 'N/A' ? 3 : 1;
      itemsToRender.sort((a, b) => getPriority(a.status) - getPriority(b.status));
    }
    
    // Group items by category
    const categoryOrder = ['Identification', 'Income', 'Assets', 'Liabilities', 'Refinance', 'Purchase & Property', 'Construction', 'Expenses', 'Other'];
    const grouped = {};
    
    itemsToRender.forEach(item => {
      const cat = item.category || 'Other';
      if (!grouped[cat]) grouped[cat] = [];
      grouped[cat].push(item);
    });
    
    // Calculate progress
    let received = 0, total = 0;
    currentEvidenceItems.forEach(item => {
      if (item.status !== 'N/A') {
        total++;
        if (item.status === 'Received') received++;
      }
    });
    const pct = total > 0 ? Math.round((received / total) * 100) : 0;
    document.getElementById('evidenceProgressFill').style.width = pct + '%';
    document.getElementById('evidenceProgressText').textContent = `${pct}% (${received}/${total})`;
    
    // Render categories
    let html = '';
    categoryOrder.forEach(cat => {
      if (!grouped[cat]) return;
      
      // Filter items
      let items = grouped[cat].filter(item => {
        if (filter === 'outstanding' && item.status !== 'Outstanding') return false;
        if (filter === 'received' && item.status !== 'Received') return false;
        if (!showNA && item.status === 'N/A') return false;
        return true;
      });
      
      if (items.length === 0) return;
      
      html += `
        <div class="evidence-category" data-category="${cat}">
          <div class="evidence-category-header" onclick="toggleEvidenceCategory('${cat}')">
            <h3>${cat}</h3>
            <span class="evidence-category-toggle" id="evidence-cat-toggle-${cat.replace(/[^a-zA-Z]/g, '')}">▼</span>
          </div>
          <div class="evidence-category-items" id="evidence-cat-items-${cat.replace(/[^a-zA-Z]/g, '')}">
            ${items.map(item => renderEvidenceItem(item)).join('')}
          </div>
        </div>
      `;
    });
    
    container.innerHTML = html || '<p style="text-align:center; color:#888; padding:40px;">No items match the current filter.</p>';
  }

  function renderEvidenceItem(item) {
    const statusIcon = item.status === 'Received' ? '☑' : item.status === 'N/A' ? '─' : '○';
    const statusClass = item.status === 'Received' ? 'received' : item.status === 'N/A' ? 'na' : 'outstanding';
    const selectClass = item.status.toLowerCase().replace('/', '');
    
    const formatPerthDate = (dateStr) => {
      if (!dateStr) return '';
      try {
        const d = new Date(dateStr);
        const perthOffset = 8 * 60;
        const localOffset = d.getTimezoneOffset();
        const perthTime = new Date(d.getTime() + (perthOffset + localOffset) * 60000);
        const hours = String(perthTime.getHours()).padStart(2, '0');
        const mins = String(perthTime.getMinutes()).padStart(2, '0');
        const day = String(perthTime.getDate()).padStart(2, '0');
        const month = String(perthTime.getMonth() + 1).padStart(2, '0');
        const year = perthTime.getFullYear();
        return `${day}/${month}/${year} at ${hours}:${mins}`;
      } catch (e) { return ''; }
    };
    
    let metaHtml = '';
    if (item.status === 'Received' && item.dateReceived) {
      metaHtml = `<div class="evidence-item-meta">Received ${formatPerthDate(item.dateReceived)}</div>`;
    } else if (item.requestedOn) {
      metaHtml = `<div class="evidence-item-meta">Requested by ${item.requestedByName || 'Unknown'} on ${formatPerthDate(item.requestedOn)}</div>`;
    }
    
    // Helper to check if rich text has actual content (not just empty tags)
    const hasRealContent = (html) => {
      if (!html) return false;
      const stripped = html.replace(/<[^>]*>/g, '').replace(/&nbsp;/g, ' ').replace(/\s+/g, ' ').trim();
      return stripped.length > 0;
    };
    
    // Notes with who/when metadata - only show if there's actual content
    let notesHtml = '';
    if (hasRealContent(item.notes)) {
      const notesMeta = item.modifiedByName && item.modifiedOn 
        ? `<span class="evidence-notes-meta">${item.modifiedByName} - ${formatPerthDate(item.modifiedOn)}</span>`
        : '';
      notesHtml = `<div class="evidence-item-notes"><strong>Internal Notes:</strong> ${item.notes}${notesMeta}</div>`;
    }
    
    // Build name and description - keep HTML for links to work
    const itemName = item.name || 'Unnamed Item';
    const hasDesc = hasRealContent(item.description);
    
    return `
      <div class="evidence-item status-${statusClass}" data-item-id="${item.id}">
        <div class="evidence-item-row">
          <div class="evidence-item-status ${statusClass}">${statusIcon}</div>
          <div class="evidence-item-text">
            <strong>${itemName}</strong>${hasDesc ? ' – ' : ''}<span class="evidence-item-desc-inline">${hasDesc ? item.description : ''}</span>
          </div>
          <div class="evidence-item-actions">
            <select class="evidence-status-select ${selectClass}" onchange="updateEvidenceItemStatus('${item.id}', this.value)">
              <option value="Outstanding" ${item.status === 'Outstanding' ? 'selected' : ''}>Outstanding</option>
              <option value="Received" ${item.status === 'Received' ? 'selected' : ''}>Received</option>
              <option value="N/A" ${item.status === 'N/A' ? 'selected' : ''}>N/A</option>
            </select>
            <button type="button" class="evidence-item-edit-btn" onclick="editEvidenceItem('${item.id}')">✎</button>
          </div>
        </div>
        ${notesHtml || metaHtml ? `<div class="evidence-item-details">${notesHtml}${metaHtml}</div>` : ''}
      </div>
    `;
  }

  window.toggleEvidenceCategory = function(cat) {
    const catId = cat.replace(/[^a-zA-Z]/g, '');
    const items = document.getElementById('evidence-cat-items-' + catId);
    const toggle = document.getElementById('evidence-cat-toggle-' + catId);
    if (items && toggle) {
      const isHidden = items.style.display === 'none';
      items.style.display = isHidden ? 'block' : 'none';
      toggle.textContent = isHidden ? '▼' : '▶';
    }
  };

  window.filterEvidenceItems = function() {
    renderEvidenceItems();
  };

  window.updateEvidenceItemStatus = function(itemId, newStatus) {
    google.script.run
      .withSuccessHandler(function() {
        // Update local state
        const item = currentEvidenceItems.find(i => i.id === itemId);
        if (item) {
          item.status = newStatus;
          if (newStatus === 'Received') {
            item.dateReceived = new Date().toISOString();
          }
        }
        renderEvidenceItems();
      })
      .withFailureHandler(function(err) {
        console.error('Error updating evidence item:', err);
        alert('Error updating status');
      })
      .updateEvidenceItem(itemId, { status: newStatus });
  };

  // Edit Evidence Item Modal
  let editEvidenceDescQuill = null;

  window.editEvidenceItem = function(itemId) {
    const item = currentEvidenceItems.find(i => i.id === itemId);
    if (!item) return;
    
    // Populate the edit modal
    document.getElementById('editEvidenceItemId').value = itemId;
    document.getElementById('editEvidenceCategory').value = item.category || 'Other';
    document.getElementById('editEvidenceName').value = item.name || '';
    document.getElementById('editEvidenceNotes').value = item.notes || '';
    
    const modal = document.getElementById('editEvidenceItemModal');
    modal.classList.add('visible');
    setTimeout(() => modal.classList.add('showing'), 10);
    
    // Initialize Quill editor if not already
    if (!editEvidenceDescQuill) {
      editEvidenceDescQuill = new Quill('#editEvidenceDescEditor', {
        theme: 'snow',
        modules: {
          toolbar: '#editEvidenceDescToolbar'
        },
        placeholder: 'Description...'
      });
      
      // Auto-prepend https:// to links without protocol
      const toolbar = editEvidenceDescQuill.getModule('toolbar');
      toolbar.addHandler('link', function(value) {
        if (value) {
          let href = prompt('Enter the link URL:');
          if (href) {
            if (!/^https?:\/\//i.test(href) && !/^mailto:/i.test(href)) {
              href = 'https://' + href;
            }
            const range = editEvidenceDescQuill.getSelection();
            if (range && range.length > 0) {
              editEvidenceDescQuill.format('link', href);
            } else {
              editEvidenceDescQuill.insertText(range ? range.index : 0, href, 'link', href);
            }
          }
        } else {
          editEvidenceDescQuill.format('link', false);
        }
      });
    }
    
    // Set the description content
    if (item.description) {
      editEvidenceDescQuill.root.innerHTML = item.description;
    } else {
      editEvidenceDescQuill.setContents([]);
    }
  };

  window.closeEditEvidenceItemModal = function() {
    const modal = document.getElementById('editEvidenceItemModal');
    modal.classList.remove('showing');
    setTimeout(() => modal.classList.remove('visible'), 200);
  };

  window.saveEditedEvidenceItem = function() {
    const itemId = document.getElementById('editEvidenceItemId').value;
    const name = document.getElementById('editEvidenceName').value.trim();
    const category = document.getElementById('editEvidenceCategory').value;
    const description = editEvidenceDescQuill ? editEvidenceDescQuill.root.innerHTML : '';
    const notes = document.getElementById('editEvidenceNotes').value;
    
    if (!name) {
      alert('Please enter a name for the item.');
      return;
    }
    
    google.script.run
      .withSuccessHandler(function(result) {
        if (result.success) {
          closeEditEvidenceItemModal();
          loadEvidenceItems();
        } else {
          alert('Error: ' + (result.error || 'Unknown error'));
        }
      })
      .withFailureHandler(function(err) {
        console.error('Error updating evidence item:', err);
        alert('Error saving changes');
      })
      .updateEvidenceItem(itemId, {
        name: name,
        description: description,
        category: category,
        notes: notes
      });
  };

  let pendingDeleteEvidenceItemId = null;
  
  window.deleteEvidenceItem = function() {
    const itemId = document.getElementById('editEvidenceItemId').value;
    const item = currentEvidenceItems.find(i => i.id === itemId);
    
    pendingDeleteEvidenceItemId = itemId;
    document.getElementById('deleteEvidenceConfirmMessage').innerText = `Are you sure you want to delete "${item?.name || 'this item'}"? This action cannot be undone.`;
    openModal('deleteEvidenceConfirmModal');
  };
  
  window.closeDeleteEvidenceConfirmModal = function() {
    closeModal('deleteEvidenceConfirmModal');
    pendingDeleteEvidenceItemId = null;
  };
  
  window.executeDeleteEvidenceItem = function() {
    if (!pendingDeleteEvidenceItemId) return;
    
    const itemId = pendingDeleteEvidenceItemId;
    closeDeleteEvidenceConfirmModal();
    
    google.script.run
      .withSuccessHandler(function(result) {
        if (result.success) {
          closeEditEvidenceItemModal();
          loadEvidenceItems();
        } else {
          showAlert('error', 'Delete Failed', 'Error: ' + (result.error || 'Unknown error'));
        }
      })
      .withFailureHandler(function(err) {
        console.error('Error deleting evidence item:', err);
        showAlert('error', 'Error', 'Error deleting item');
      })
      .deleteEvidenceItem(itemId);
  };

  window.populateEvidenceFromTemplates = function() {
    // Show loading state in appropriate place
    const emptyState = document.getElementById('evidenceEmptyState');
    const hasExistingItems = currentEvidenceItems.length > 0;
    
    if (!hasExistingItems) {
      emptyState.innerHTML = '<p>Populating evidence list...</p>';
    }
    
    google.script.run
      .withSuccessHandler(function(result) {
        if (result.success) {
          if (result.itemsCreated === 0) {
            showAlert('info', 'No Templates Found', 'No templates found for this opportunity type/lender. You can add custom items using the "+ Add Custom" button, or create templates in Airtable\'s "Evidence Templates" table.');
            if (!hasExistingItems) {
              emptyState.innerHTML = '<p>No evidence items yet.</p><button type="button" class="evidence-btn-primary" onclick="populateEvidenceFromTemplates()">Populate from Templates</button>';
            }
          } else {
            showAlert('success', 'Templates Added', result.itemsCreated + ' item(s) added from templates.');
            loadEvidenceItems();
          }
        } else {
          showAlert('error', 'Error', result.error || 'Unknown error');
          if (!hasExistingItems) {
            emptyState.innerHTML = '<p>No evidence items yet.</p><button type="button" class="evidence-btn-primary" onclick="populateEvidenceFromTemplates()">Populate from Templates</button>';
          }
        }
      })
      .withFailureHandler(function(err) {
        console.error('Error populating evidence:', err);
        showAlert('error', 'Error', 'Error populating evidence list');
        if (!hasExistingItems) {
          emptyState.innerHTML = '<p>No evidence items yet.</p><button type="button" class="evidence-btn-primary" onclick="populateEvidenceFromTemplates()">Populate from Templates</button>';
        }
      })
      .populateEvidenceForOpportunity(currentEvidenceOpportunityId, currentEvidenceOpportunityType, currentEvidenceLender);
  };

  window.toggleEvidenceEmailMenu = function() {
    const menu = document.getElementById('evidenceEmailMenu');
    menu.classList.toggle('show');
    // Close on outside click
    setTimeout(() => {
      document.addEventListener('click', closeEvidenceEmailMenuOnOutside, { once: true });
    }, 10);
  };

  function closeEvidenceEmailMenuOnOutside(e) {
    const menu = document.getElementById('evidenceEmailMenu');
    const btn = document.querySelector('.evidence-email-btn');
    if (!menu.contains(e.target) && !btn.contains(e.target)) {
      menu.classList.remove('show');
    }
  }

  // Build client-facing HTML for evidence list (no internal notes, no meta, no edit buttons)
  function buildClientEvidenceMarkup() {
    // Split items by status - exclude N/A entirely
    const outstanding = currentEvidenceItems.filter(i => i.status === 'Outstanding');
    const received = currentEvidenceItems.filter(i => i.status === 'Received');
    
    // Calculate progress
    const total = outstanding.length + received.length;
    const pct = total > 0 ? Math.round((received.length / total) * 100) : 0;
    
    // Helper to strip block HTML but keep inline formatting
    const cleanDesc = (desc) => {
      if (!desc) return '';
      return desc
        .replace(/<p>/gi, '')
        .replace(/<\/p>/gi, ' ')
        .replace(/<br\s*\/?>/gi, ' ')
        .replace(/<div>/gi, '')
        .replace(/<\/div>/gi, ' ')
        .replace(/\s+/g, ' ')
        .trim();
    };
    
    const renderItem = (item) => {
      const statusIcon = item.status === 'Received' ? '✓' : '○';
      const statusColor = item.status === 'Received' ? '#7B8B64' : '#2C2622';
      const descText = cleanDesc(item.description);
      
      let itemText = `<strong>${item.name || 'Item'}</strong>`;
      if (descText) {
        itemText += ` – <span style="font-weight:normal;">${descText}</span>`;
      }
      
      return `<li style="margin-bottom:6px; color:${statusColor};">${statusIcon} ${itemText}</li>`;
    };
    
    let html = '<div style="font-size:13px; font-family:Arial,sans-serif; line-height:1.5;">';
    
    // Stellaris leaf icon
    const leafIcon = `<img src="https://img1.wsimg.com/isteam/ip/2c5f94ee-4964-4e9b-9b9c-a55121f8611b/favicon/31eb51a1-8979-4194-bfa2-e4b30ee1178d/2437d5de-854d-40b2-86b2-fd879f3469f0.png" width="18" height="18" style="width:18px; height:18px; flex-shrink:0;">`;
    
    // Progress bar
    html += `<div style="margin-bottom:15px;">`;
    html += `<div style="display:flex; align-items:center; gap:10px; max-width:50%;">`;
    html += leafIcon;
    html += `<div style="flex:1; height:10px; background:#E0E0E0; border-radius:5px; overflow:hidden;">`;
    html += `<div style="width:${pct}%; height:100%; background:#7B8B64; border-radius:5px;"></div>`;
    html += `</div>`;
    html += `<span style="font-weight:bold; color:#2C2622; white-space:nowrap;">${pct}% (${received.length}/${total})</span>`;
    html += `</div></div>`;
    
    // Outstanding section
    if (outstanding.length > 0) {
      html += `<div style="margin-bottom:15px;">`;
      html += `<h4 style="margin:0 0 8px 0; font-size:13px; font-weight:bold; color:#2C2622;">Outstanding</h4>`;
      html += `<ul style="margin:0; padding-left:0; list-style:none;">`;
      outstanding.forEach(item => { html += renderItem(item); });
      html += `</ul></div>`;
    }
    
    // Received section
    if (received.length > 0) {
      html += `<div style="margin-bottom:15px;">`;
      html += `<h4 style="margin:0 0 8px 0; font-size:13px; font-weight:bold; color:#7B8B64;">Received</h4>`;
      html += `<ul style="margin:0; padding-left:0; list-style:none;">`;
      received.forEach(item => { html += renderItem(item); });
      html += `</ul></div>`;
    }
    
    html += '</div>';
    
    if (outstanding.length === 0 && received.length === 0) {
      return '<p style="color:#888;">No items to display.</p>';
    }
    
    return html;
  }

  window.openEvidenceClientView = function() {
    const content = buildClientEvidenceMarkup();
    document.getElementById('evidenceClientViewContent').innerHTML = content;
    
    const modal = document.getElementById('evidenceClientViewModal');
    modal.classList.add('visible');
    setTimeout(() => modal.classList.add('showing'), 10);
  };

  window.closeEvidenceClientView = function() {
    const modal = document.getElementById('evidenceClientViewModal');
    modal.classList.remove('showing');
    setTimeout(() => modal.classList.remove('visible'), 200);
  };

  window.copyEvidenceClientView = function() {
    const html = buildClientEvidenceMarkup(); // Use full rich format for email clients
    const plainText = html.replace(/<[^>]*>/g, '').replace(/\s+/g, ' ').trim();
    
    // Copy as rich HTML for email clients
    const blob = new Blob([html], { type: 'text/html' });
    const clipboardItem = new ClipboardItem({
      'text/html': blob,
      'text/plain': new Blob([plainText], { type: 'text/plain' })
    });
    
    navigator.clipboard.write([clipboardItem]).then(() => {
      // Show brief "Copied" feedback
      const btn = document.querySelector('#evidenceClientViewModal .btn-primary');
      const originalText = btn.textContent;
      btn.textContent = 'Copied!';
      btn.style.background = '#7B8B64';
      setTimeout(() => {
        btn.textContent = originalText;
        btn.style.background = '';
        closeEvidenceClientView();
      }, 800);
    }).catch(err => {
      console.error('Clipboard error:', err);
      // Fallback to plain text
      navigator.clipboard.writeText(plainText).then(() => {
        const btn = document.querySelector('#evidenceClientViewModal .btn-primary');
        btn.textContent = 'Copied (plain text)';
        btn.style.background = '#7B8B64';
        setTimeout(() => {
          btn.textContent = 'Copy to Clipboard';
          btn.style.background = '';
          closeEvidenceClientView();
        }, 800);
      }).catch(() => {
        const btn = document.querySelector('#evidenceClientViewModal .btn-primary');
        btn.textContent = 'Copy failed';
        btn.style.background = '#C44';
        setTimeout(() => {
          btn.textContent = 'Copy to Clipboard';
          btn.style.background = '';
        }, 1500);
      });
    });
  };

  window.copyEvidenceToClipboard = function() {
    // Legacy function - now redirects to client view
    openEvidenceClientView();
  };

  let evidenceEmailQuill = null;
  let pendingEvidenceEmailItemIds = [];
  let currentEvidenceEmailType = null;
  
  window.generateEvidenceEmail = function(type) {
    document.getElementById('evidenceEmailMenu').classList.remove('show');
    
    const outstanding = currentEvidenceItems.filter(i => i.status === 'Outstanding');
    if (outstanding.length === 0 && type !== 'appointment') {
      showAlert('info', 'No Items', 'No outstanding items to request!');
      return;
    }
    
    currentEvidenceEmailType = type;
    pendingEvidenceEmailItemIds = outstanding.map(i => i.id);
    
    // Get contact info
    const contactName = currentContactRecord?.fields?.PreferredName || currentContactRecord?.fields?.FirstName || 'there';
    const contactEmail = currentContactRecord?.fields?.EmailAddress1 || '';
    
    // Build rich HTML items list - matching client view format exactly
    const received = currentEvidenceItems.filter(i => i.status === 'Received');
    const total = outstanding.length + received.length;
    const pct = total > 0 ? Math.round((received.length / total) * 100) : 0;
    
    // Helper to clean descriptions (keep inline links)
    const cleanDesc = (desc) => {
      if (!desc) return '';
      return desc.replace(/<p>/gi, '').replace(/<\/p>/gi, ' ').replace(/<br\s*\/?>/gi, ' ').replace(/<div>/gi, '').replace(/<\/div>/gi, ' ').replace(/\s+/g, ' ').trim();
    };
    
    // Build evidence list HTML using table-based layout for Quill/email compatibility
    // Note: Quill strips flexbox and many CSS styles, so we use tables with inline styles
    const barWidth = Math.max(pct * 2, 10); // Scale to reasonable pixel width (max 200px for 100%)
    const barBgWidth = 200;
    
    let evidenceListHtml = '';
    
    // Progress bar - simple text format that Quill preserves reliably
    // Quill strips complex HTML/tables, so we use a clean text representation
    evidenceListHtml += `<p style="margin:8px 0 16px 0;"><strong style="color:#7B8B64;">Progress: ${pct}% (${received.length}/${total})</strong></p>`;
    
    // Outstanding section with bold header
    if (outstanding.length > 0) {
      evidenceListHtml += `<p style="margin:16px 0 8px 0;"><strong style="color:#2C2622;">Outstanding</strong></p>`;
      outstanding.forEach(item => {
        const desc = cleanDesc(item.description);
        let line = `○ <strong>${item.name || 'Item'}</strong>`;
        if (desc) line += ` – ${desc}`;
        evidenceListHtml += `<p style="margin:6px 0 6px 12px; color:#2C2622;">${line}</p>`;
      });
      evidenceListHtml += `<p style="margin:0;"></p>`; // spacer after list
    }
    
    // Received section with bold green header
    if (received.length > 0) {
      evidenceListHtml += `<p style="margin:16px 0 8px 0;"><strong style="color:#7B8B64;">Received</strong></p>`;
      received.forEach(item => {
        const desc = cleanDesc(item.description);
        let line = `✓ <strong>${item.name || 'Item'}</strong>`;
        if (desc) line += ` – ${desc}`;
        evidenceListHtml += `<p style="margin:6px 0 6px 12px; color:#7B8B64;">${line}</p>`;
      });
      evidenceListHtml += `<p style="margin:0;"></p>`; // spacer after list
    }
    
    // Build email content based on type
    let subject, body, title;
    
    if (type === 'initial') {
      title = 'Initial Request Email';
      subject = `Documents needed for your ${currentEvidenceOpportunityName}`;
      body = `<p>Hi ${contactName},</p>
<p>Thank you for choosing Stellaris Finance! To get your application moving, we need the following documents:</p>
${evidenceListHtml}
<p>Simply reply to this email with the documents attached. If you have any questions, don't hesitate to reach out!</p>
<p>Kind regards,</p>`;
    } else if (type === 'subsequent') {
      title = 'Subsequent Request Email';
      subject = `Quick follow-up: Documents still needed for ${currentEvidenceOpportunityName}`;
      body = `<p>Hi ${contactName},</p>
<p>Just a quick follow-up on your application. We're still waiting on a few items:</p>
${evidenceListHtml}
<p>Once we have these, we can move to the next stage. Let me know if you need any help!</p>
<p>Kind regards,</p>`;
    } else if (type === 'appointment') {
      title = 'Appointment Confirmation Email';
      subject = `Your upcoming appointment – ${currentEvidenceOpportunityName}`;
      
      // Format appointment details from cached opportunity appointments
      let apptDetails = '';
      if (window.currentOpportunityAppointments && window.currentOpportunityAppointments.length > 0) {
        // Find next upcoming appointment
        const now = new Date();
        const upcoming = window.currentOpportunityAppointments
          .filter(a => a.appointmentTime && new Date(a.appointmentTime) > now)
          .sort((a, b) => new Date(a.appointmentTime) - new Date(b.appointmentTime))[0];
        
        if (upcoming) {
          const apptDate = new Date(upcoming.appointmentTime);
          const perthOffset = 8 * 60;
          const localOffset = apptDate.getTimezoneOffset();
          const perthTime = new Date(apptDate.getTime() + (perthOffset + localOffset) * 60000);
          
          const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
          const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
          const dayName = days[perthTime.getDay()];
          const monthName = months[perthTime.getMonth()];
          const dayNum = perthTime.getDate();
          const hours = perthTime.getHours();
          const mins = String(perthTime.getMinutes()).padStart(2, '0');
          const ampm = hours >= 12 ? 'PM' : 'AM';
          const hour12 = hours % 12 || 12;
          
          apptDetails = `<p><strong>Date:</strong> ${dayName}, ${dayNum} ${monthName}<br>`;
          apptDetails += `<strong>Time:</strong> ${hour12}:${mins} ${ampm} (Perth time)<br>`;
          apptDetails += `<strong>Type:</strong> ${upcoming.typeOfAppointment || 'Meeting'}`;
          
          if (upcoming.typeOfAppointment === 'Phone' && upcoming.phoneNumber) {
            apptDetails += `<br><strong>Phone:</strong> ${upcoming.phoneNumber}`;
          } else if (upcoming.typeOfAppointment === 'Video' && upcoming.videoMeetUrl) {
            apptDetails += `<br><strong>Join Link:</strong> <a href="${upcoming.videoMeetUrl}">${upcoming.videoMeetUrl}</a>`;
          } else if (upcoming.typeOfAppointment === 'Office') {
            apptDetails += `<br><strong>Location:</strong> Stellaris Finance Office`;
          }
          apptDetails += '</p>';
        } else {
          apptDetails = '<p><em>[No upcoming appointments found - please add appointment details]</em></p>';
        }
      } else {
        apptDetails = '<p><em>[No appointments on record - please add appointment details]</em></p>';
      }
      
      body = `<p>Hi ${contactName},</p>
<p>This is to confirm your upcoming appointment with Stellaris Finance.</p>
${apptDetails}`;
      
      if (outstanding.length > 0) {
        body += `<p>To make the most of our meeting, please send the following items beforehand if possible:</p>
${evidenceListHtml}`;
      }
      
      body += `<p>We look forward to speaking with you!</p>
<p>Kind regards,</p>`;
    }
    
    // Store the base email body HTML for reset functionality
    window.pendingEvidenceEmailBodyBase = body;
    
    // Populate modal
    document.getElementById('evidenceEmailModalTitle').textContent = title;
    document.getElementById('evidenceEmailTo').value = contactEmail;
    document.getElementById('evidenceEmailSubject').value = subject;
    document.getElementById('evidenceEmailItemCount').textContent = outstanding.length;
    
    // Open modal first so iframe is rendered
    const modal = document.getElementById('evidenceEmailModal');
    modal.classList.add('visible');
    setTimeout(() => modal.classList.add('showing'), 10);
    
    // Initialize the iframe editor with the email content
    setTimeout(() => initEmailEditorIframe(body), 100);
  };
  
  function initEmailEditorIframe(htmlContent) {
    const iframe = document.getElementById('evidenceEmailEditor');
    if (!iframe) return;
    
    const doc = iframe.contentDocument || iframe.contentWindow.document;
    doc.open();
    doc.write(`
      <!DOCTYPE html>
      <html>
      <head>
        <style>
          body {
            font-family: Arial, sans-serif;
            font-size: 14px;
            line-height: 1.6;
            padding: 20px;
            margin: 0;
            color: #2C2622;
          }
          p { margin: 0 0 12px 0; }
          a { color: #19414C; }
          strong { font-weight: bold; }
        </style>
      </head>
      <body>${htmlContent}</body>
      </html>
    `);
    doc.close();
    doc.designMode = 'on';
  }
  
  window.emailEditorCommand = function(command) {
    const iframe = document.getElementById('evidenceEmailEditor');
    if (!iframe) return;
    
    const doc = iframe.contentDocument || iframe.contentWindow.document;
    
    if (command === 'createLink') {
      const url = prompt('Enter URL:', 'https://');
      if (url) {
        doc.execCommand('createLink', false, url);
      }
    } else {
      doc.execCommand(command, false, null);
    }
    
    iframe.contentWindow.focus();
  };
  
  window.resetEmailToTemplate = function() {
    if (confirm('Reset email to the original template? Your edits will be lost.')) {
      initEmailEditorIframe(window.pendingEvidenceEmailBodyBase || '');
    }
  };
  
  function getEmailEditorContent() {
    const iframe = document.getElementById('evidenceEmailEditor');
    if (!iframe) {
      console.warn('Email editor iframe not found, using stored template');
      return window.pendingEvidenceEmailBodyBase || '';
    }
    
    try {
      const doc = iframe.contentDocument || iframe.contentWindow.document;
      if (doc && doc.body) {
        const content = doc.body.innerHTML;
        // If content is empty or just whitespace, fall back to stored template
        if (content && content.trim() && content.trim() !== '<br>') {
          return content;
        }
      }
    } catch (e) {
      console.warn('Could not read iframe content:', e);
    }
    
    // Fallback to stored template
    return window.pendingEvidenceEmailBodyBase || '';
  }
  
  window.closeEvidenceEmailModal = function() {
    const modal = document.getElementById('evidenceEmailModal');
    modal.classList.remove('showing');
    setTimeout(() => modal.classList.remove('visible'), 200);
    pendingEvidenceEmailItemIds = [];
    currentEvidenceEmailType = null;
    window.pendingEvidenceEmailBody = null;
    window.pendingEvidenceEmailBodyBase = null;
  };
  
  window.sendEvidenceEmail = function() {
    const to = document.getElementById('evidenceEmailTo').value.trim();
    const subject = document.getElementById('evidenceEmailSubject').value.trim();
    const body = getEmailEditorContent();
    
    if (!to) {
      showAlert('warning', 'Missing Recipient', 'Please enter an email address.');
      return;
    }
    
    if (!body || body.trim() === '<br>' || body.trim() === '') {
      showAlert('warning', 'Empty Message', 'Please add some content to the email.');
      return;
    }
    
    const sendBtn = document.getElementById('evidenceEmailSendBtn');
    sendBtn.textContent = 'Sending...';
    sendBtn.disabled = true;
    
    // Send the email
    google.script.run
      .withSuccessHandler(function(result) {
        if (result && result.success) {
          // Mark items as requested
          if (pendingEvidenceEmailItemIds.length > 0) {
            google.script.run
              .withSuccessHandler(function() {
                // Update local data
                pendingEvidenceEmailItemIds.forEach(itemId => {
                  const item = currentEvidenceItems.find(i => i.id === itemId);
                  if (item) {
                    item.requestedOn = new Date().toISOString();
                  }
                });
                renderEvidenceItems();
              })
              .markEvidenceItemsAsRequested(pendingEvidenceEmailItemIds);
          }
          
          closeEvidenceEmailModal();
          showAlert('success', 'Email Sent', 'Email sent successfully! Outstanding items have been marked as requested.');
        } else {
          showAlert('error', 'Send Failed', result?.error || 'Failed to send email.');
        }
        sendBtn.textContent = 'Send Email';
        sendBtn.disabled = false;
      })
      .withFailureHandler(function(err) {
        showAlert('error', 'Error', 'Failed to send email: ' + (err.message || 'Unknown error'));
        sendBtn.textContent = 'Send Email';
        sendBtn.disabled = false;
      })
      .sendEmail(to, subject, body);
  };

  // Add Evidence Item Modal
  // Quill editor instance for evidence description
  let newEvidenceDescQuill = null;

  window.openAddEvidenceItemModal = function() {
    document.getElementById('newEvidenceName').value = '';
    document.getElementById('newEvidenceCategory').value = 'Other';
    
    const modal = document.getElementById('addEvidenceItemModal');
    modal.classList.add('visible');
    setTimeout(() => modal.classList.add('showing'), 10);
    
    // Initialize Quill editor if not already
    if (!newEvidenceDescQuill) {
      newEvidenceDescQuill = new Quill('#newEvidenceDescEditor', {
        theme: 'snow',
        modules: {
          toolbar: '#newEvidenceDescToolbar'
        },
        placeholder: 'Describe what you need - supports bold, links, bullet points...'
      });
      
      // Auto-prepend https:// to links without protocol
      const toolbar = newEvidenceDescQuill.getModule('toolbar');
      toolbar.addHandler('link', function(value) {
        if (value) {
          let href = prompt('Enter the link URL:');
          if (href) {
            // Auto-prepend https:// if no protocol specified
            if (!/^https?:\/\//i.test(href) && !/^mailto:/i.test(href)) {
              href = 'https://' + href;
            }
            const range = newEvidenceDescQuill.getSelection();
            if (range && range.length > 0) {
              newEvidenceDescQuill.format('link', href);
            } else {
              // Insert the URL as text if no selection
              newEvidenceDescQuill.insertText(range ? range.index : 0, href, 'link', href);
            }
          }
        } else {
          newEvidenceDescQuill.format('link', false);
        }
      });
    } else {
      newEvidenceDescQuill.setContents([]);
    }
  };

  window.closeAddEvidenceItemModal = function() {
    const modal = document.getElementById('addEvidenceItemModal');
    modal.classList.remove('showing');
    setTimeout(() => modal.classList.remove('visible'), 200);
  };

  window.submitNewEvidenceItem = function() {
    const name = document.getElementById('newEvidenceName').value.trim();
    const description = newEvidenceDescQuill ? newEvidenceDescQuill.root.innerHTML : '';
    const category = document.getElementById('newEvidenceCategory').value;
    
    if (!name) {
      alert('Please enter a name for the item.');
      return;
    }
    
    google.script.run
      .withSuccessHandler(function(result) {
        if (result.success) {
          closeAddEvidenceItemModal();
          loadEvidenceItems();
        } else {
          alert('Error: ' + (result.error || 'Unknown error'));
        }
      })
      .withFailureHandler(function(err) {
        console.error('Error creating evidence item:', err);
        alert('Error creating item');
      })
      .createEvidenceItem(currentEvidenceOpportunityId, {
        name: name,
        description: description,
        category: category
      });
  };

  // --- EVIDENCE TEMPLATES MANAGEMENT ---
  let allEvidenceTemplates = [];
  let editingEvidenceTemplateId = null;
  let evTplDescQuill = null;

  window.openEvidenceTemplatesModal = function() {
    const modal = document.getElementById('evidenceTemplatesModal');
    if (!modal) return;
    
    loadAllEvidenceTemplates();
    modal.style.display = 'flex';
    setTimeout(() => modal.classList.add('showing'), 10);
  };

  window.closeEvidenceTemplatesModal = function() {
    const modal = document.getElementById('evidenceTemplatesModal');
    if (modal) {
      modal.classList.remove('showing');
      setTimeout(() => modal.style.display = 'none', 250);
    }
  };

  function loadAllEvidenceTemplates() {
    const container = document.getElementById('evidenceTemplatesContainer');
    if (container) container.innerHTML = '<div style="text-align:center; padding:20px; color:#888;">Loading templates...</div>';
    
    google.script.run
      .withSuccessHandler(function(templates) {
        allEvidenceTemplates = templates || [];
        renderEvidenceTemplatesList();
        populateEvidenceTemplatesFilter();
      })
      .withFailureHandler(function(err) {
        console.error('Error loading evidence templates:', err);
        if (container) container.innerHTML = '<div style="text-align:center; padding:20px; color:#c44;">Failed to load templates</div>';
      })
      .getAllEvidenceTemplates();
  }

  function populateEvidenceTemplatesFilter() {
    const filter = document.getElementById('evidenceTemplatesFilter');
    if (!filter) return;
    
    const categories = [...new Set(allEvidenceTemplates.map(t => t.categoryName).filter(Boolean))].sort();
    let html = '<option value="">All Categories</option>';
    categories.forEach(cat => {
      html += `<option value="${cat}">${cat}</option>`;
    });
    filter.innerHTML = html;
    
    filter.onchange = renderEvidenceTemplatesList;
    document.getElementById('evidenceTemplatesSearch').oninput = renderEvidenceTemplatesList;
  }

  function renderEvidenceTemplatesList() {
    const container = document.getElementById('evidenceTemplatesContainer');
    if (!container) return;
    
    const searchTerm = (document.getElementById('evidenceTemplatesSearch')?.value || '').toLowerCase();
    const categoryFilter = document.getElementById('evidenceTemplatesFilter')?.value || '';
    
    let filtered = allEvidenceTemplates.filter(t => {
      if (categoryFilter && t.categoryName !== categoryFilter) return false;
      if (searchTerm && !t.name.toLowerCase().includes(searchTerm) && !(t.description || '').toLowerCase().includes(searchTerm)) return false;
      return true;
    });
    
    if (filtered.length === 0) {
      container.innerHTML = '<div style="text-align:center; padding:30px; color:#888;">No templates match your search.</div>';
      return;
    }
    
    const grouped = {};
    filtered.forEach(t => {
      const cat = t.categoryName || 'Uncategorized';
      if (!grouped[cat]) grouped[cat] = [];
      grouped[cat].push(t);
    });
    
    let html = '';
    Object.keys(grouped).sort().forEach(cat => {
      html += `<div style="margin-bottom:20px;">`;
      html += `<h4 style="margin:0 0 10px; color:var(--color-midnight); font-size:13px; border-bottom:1px solid #ddd; padding-bottom:5px;">${cat}</h4>`;
      grouped[cat].forEach(t => {
        const desc = (t.description || '').replace(/<[^>]*>/g, '').substring(0, 80);
        html += `<div class="evidence-template-item" onclick="openEvidenceTemplateEdit('${t.id}')" style="display:flex; justify-content:space-between; align-items:center; padding:10px 12px; background:#fff; border:1px solid #eee; border-radius:4px; margin-bottom:6px; cursor:pointer; transition:background 0.15s;">`;
        html += `<div><strong style="color:#2C2622;">${t.name}</strong>`;
        if (desc) html += `<div style="font-size:11px; color:#888; margin-top:2px;">${desc}${t.description && t.description.length > 80 ? '...' : ''}</div>`;
        html += `</div>`;
        html += `<span style="font-size:11px; color:#888; flex-shrink:0;">#${t.displayOrder}</span>`;
        html += `</div>`;
      });
      html += `</div>`;
    });
    
    container.innerHTML = html;
  }

  window.openNewEvidenceTemplateForm = function() {
    editingEvidenceTemplateId = null;
    document.getElementById('evidenceTemplateEditTitle').textContent = 'New Evidence Template';
    document.getElementById('evTplEditName').value = '';
    document.getElementById('evTplEditOrder').value = 100;
    document.getElementById('evTplEditLenderSpecific').checked = false;
    document.getElementById('evTplDeleteBtn').style.display = 'none';
    
    loadCategoriesForTemplateEdit();
    loadOppTypesForTemplateEdit([]);
    
    const modal = document.getElementById('evidenceTemplateEditModal');
    modal.style.display = 'flex';
    setTimeout(() => modal.classList.add('showing'), 10);
    
    initEvTplQuillEditor('');
  };

  window.openEvidenceTemplateEdit = function(templateId) {
    const template = allEvidenceTemplates.find(t => t.id === templateId);
    if (!template) return;
    
    editingEvidenceTemplateId = templateId;
    document.getElementById('evidenceTemplateEditTitle').textContent = 'Edit Template';
    document.getElementById('evTplEditName').value = template.name || '';
    document.getElementById('evTplEditOrder').value = template.displayOrder || 0;
    document.getElementById('evTplEditLenderSpecific').checked = template.isLenderSpecific || false;
    document.getElementById('evTplDeleteBtn').style.display = 'block';
    
    loadCategoriesForTemplateEdit(template.categoryId);
    loadOppTypesForTemplateEdit(template.opportunityTypes || []);
    
    const modal = document.getElementById('evidenceTemplateEditModal');
    modal.style.display = 'flex';
    setTimeout(() => modal.classList.add('showing'), 10);
    
    initEvTplQuillEditor(template.description || '');
  };

  function loadCategoriesForTemplateEdit(selectedId = null) {
    const select = document.getElementById('evTplEditCategory');
    if (!select) return;
    
    google.script.run
      .withSuccessHandler(function(categories) {
        let html = '<option value="">-- Select Category --</option>';
        (categories || []).forEach(c => {
          const selected = c.id === selectedId ? ' selected' : '';
          html += `<option value="${c.id}"${selected}>${c.name}</option>`;
        });
        select.innerHTML = html;
      })
      .getEvidenceCategories();
  }

  function loadOppTypesForTemplateEdit(selectedTypes) {
    const container = document.getElementById('evTplEditOppTypes');
    if (!container) return;
    
    // Use actual opportunity types from the system
    const oppTypes = [
      'Home Loans',
      'Commercial Loans', 
      'Deposit Bonds',
      'Insurance (General)',
      'Insurance (Life)',
      'Personal Loans',
      'Asset Finance',
      'Tax Depreciation Schedule',
      'Financial Planning'
    ];
    let html = '';
    oppTypes.forEach(type => {
      const checked = selectedTypes.includes(type) ? ' checked' : '';
      html += `<label style="display:flex; align-items:center; gap:8px; font-size:12px; cursor:pointer; white-space:nowrap;"><input type="checkbox" class="evTplOppType" value="${type}" style="width:14px; height:14px; flex-shrink:0;"${checked}> ${type}</label>`;
    });
    container.innerHTML = html;
  }

  function initEvTplQuillEditor(content) {
    const editorEl = document.getElementById('evTplEditDescEditor');
    if (!editorEl) return;
    
    if (evTplDescQuill) {
      evTplDescQuill.setText('');
      evTplDescQuill.clipboard.dangerouslyPasteHTML(content || '');
    } else {
      evTplDescQuill = new Quill(editorEl, {
        theme: 'snow',
        modules: {
          toolbar: [['bold', 'italic'], ['link'], [{ 'list': 'bullet' }]]
        }
      });
      evTplDescQuill.clipboard.dangerouslyPasteHTML(content || '');
    }
  }

  window.closeEvidenceTemplateEditModal = function() {
    const modal = document.getElementById('evidenceTemplateEditModal');
    if (modal) {
      modal.classList.remove('showing');
      setTimeout(() => modal.style.display = 'none', 250);
    }
    editingEvidenceTemplateId = null;
  };

  window.saveEvidenceTemplate = function() {
    const name = document.getElementById('evTplEditName').value.trim();
    if (!name) {
      showAlert('Error', 'Please enter a template name', 'error');
      return;
    }
    
    const fields = {
      name: name,
      description: evTplDescQuill ? evTplDescQuill.root.innerHTML : '',
      categoryId: document.getElementById('evTplEditCategory').value || null,
      displayOrder: parseInt(document.getElementById('evTplEditOrder').value, 10) || 100,
      isLenderSpecific: document.getElementById('evTplEditLenderSpecific').checked,
      opportunityTypes: Array.from(document.querySelectorAll('.evTplOppType:checked')).map(cb => cb.value)
    };
    
    const btn = document.getElementById('evTplSaveBtn');
    btn.textContent = 'Saving...';
    btn.disabled = true;
    
    const handler = function(result) {
      btn.textContent = 'Save';
      btn.disabled = false;
      if (result.success) {
        showAlert('Success', editingEvidenceTemplateId ? 'Template updated' : 'Template created', 'success');
        closeEvidenceTemplateEditModal();
        loadAllEvidenceTemplates();
      } else {
        showAlert('Error', result.error || 'Failed to save template', 'error');
      }
    };
    
    const errorHandler = function(err) {
      btn.textContent = 'Save';
      btn.disabled = false;
      showAlert('Error', 'Failed to save: ' + (err.message || 'Unknown error'), 'error');
    };
    
    if (editingEvidenceTemplateId) {
      google.script.run.withSuccessHandler(handler).withFailureHandler(errorHandler).updateEvidenceTemplate(editingEvidenceTemplateId, fields);
    } else {
      google.script.run.withSuccessHandler(handler).withFailureHandler(errorHandler).createEvidenceTemplate(fields);
    }
  };

  window.deleteEvidenceTemplate = function() {
    if (!editingEvidenceTemplateId) return;
    
    showConfirmModal('Delete Template', 'Are you sure you want to delete this template? This cannot be undone.', function() {
      google.script.run
        .withSuccessHandler(function(result) {
          if (result.success) {
            showAlert('Deleted', 'Template deleted', 'success');
            closeEvidenceTemplateEditModal();
            loadAllEvidenceTemplates();
          } else {
            showAlert('Error', result.error || 'Failed to delete', 'error');
          }
        })
        .withFailureHandler(function(err) {
          showAlert('Error', 'Failed to delete: ' + (err.message || 'Unknown error'), 'error');
        })
        .deleteEvidenceTemplate(editingEvidenceTemplateId);
    });
  };