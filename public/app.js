  let searchTimeout;
  let spouseSearchTimeout;
  let linkedSearchTimeout;
  let loadingTimer;
  let pollInterval;
  let pollAttempts = 0;
  let panelHistory = []; 
  let currentContactRecord = null; 
  let currentOppRecords = []; 
  let currentOppSortDirection = 'desc'; 
  let pendingLinkedEdits = {}; 
  let currentPanelData = {}; 
  let pendingRemovals = {}; 

  window.onload = function() { 
    loadContacts(); 
    checkUserIdentity(); 
    initKeyboardShortcuts();
    initDarkMode();
    initScreensaver();
  };

  // --- KEYBOARD SHORTCUTS ---
  function initKeyboardShortcuts() {
    document.addEventListener('keydown', function(e) {
      const activeEl = document.activeElement;
      const isTyping = activeEl.tagName === 'INPUT' || activeEl.tagName === 'TEXTAREA' || activeEl.isContentEditable;
      
      if (e.key === 'Escape') {
        closeOppPanel();
        closeSpouseModal();
        closeNewOppModal();
        closeShortcutsModal();
        closeDeleteConfirmModal();
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
      document.getElementById('formDivider').style.display = 'block';
    } else {
      document.getElementById('emptyState').style.display = 'flex';
      document.getElementById('profileContent').style.display = 'none';
      document.getElementById('formDivider').style.display = 'none';
      document.getElementById('formTitle').innerText = "Contact";
      document.getElementById('formSubtitle').innerText = '';
      document.getElementById('editBtn').style.visibility = 'hidden';
      document.getElementById('refreshBtn').style.display = 'none'; 
      document.getElementById('auditSection').style.display = 'none';
      document.getElementById('duplicateWarningBox').style.display = 'none'; 
    }
  }

  function enableEditMode() {
    const inputs = document.querySelectorAll('#contactForm input, #contactForm textarea');
    inputs.forEach(input => { input.classList.remove('locked'); input.readOnly = false; });
    document.getElementById('actionRow').style.display = 'flex';
    document.getElementById('cancelBtn').style.display = 'inline-block';
    document.getElementById('editBtn').style.visibility = 'hidden';
    updateHeaderTitle(true); 
  }

  function disableEditMode() {
    const inputs = document.querySelectorAll('#contactForm input, #contactForm textarea');
    inputs.forEach(input => { input.classList.add('locked'); input.readOnly = true; });
    document.getElementById('actionRow').style.display = 'none';
    document.getElementById('cancelBtn').style.display = 'none';
    updateHeaderTitle(false); 
  }
  
  function cancelEditMode() {
    const recordId = document.getElementById('recordId').value;
    document.getElementById('cancelBtn').style.display = 'none';
    if (recordId && currentContactRecord) {
      selectContact(currentContactRecord);
    } else {
      document.getElementById('contactForm').reset();
      toggleProfileView(false);
    }
    disableEditMode();
  }

  function selectContact(record) {
    document.getElementById('cancelBtn').style.display = 'none';
    toggleProfileView(true);
    currentContactRecord = record; 
    const f = record.fields;

    document.getElementById('recordId').value = record.id;
    document.getElementById('firstName').value = f.FirstName || "";
    document.getElementById('middleName').value = f.MiddleName || "";
    document.getElementById('lastName').value = f.LastName || "";
    document.getElementById('preferredName').value = f.PreferredName || "";
    document.getElementById('mobilePhone').value = f.Mobile || "";
    document.getElementById('email1').value = f.EmailAddress1 || "";
    document.getElementById('description').value = f.Description || "";

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

    document.getElementById('formSubtitle').innerText = formatSubtitle(f);
    renderHistory(f);
    loadOpportunities(f);
    renderSpouseSection(f); 
    closeOppPanel();
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
    
    currentEmailContext = {
      opportunity: opportunityData,
      contact: contactData,
      greeting: contactData.PreferredName || contactData.FirstName || 'there',
      broker: opportunityData['Taco: Broker'] || 'our Mortgage Broker',
      brokerFirst: (opportunityData['Taco: Broker'] || '').split(' ')[0] || 'the broker',
      appointmentTime: opportunityData['Taco: Appointment Time'] || '[appointment time]',
      phoneNumber: opportunityData['Taco: Appt Phone Number'] || '[phone number]',
      meetUrl: opportunityData['Taco: Appt Meet URL'] || '[Google Meet URL]',
      emails: emails,
      sender: 'Shae'
    };
    
    const apptType = opportunityData['Taco: Type of Appointment'] || 'Phone';
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
    
    // Set content in Quill editor if available, otherwise fallback to innerHTML
    if (emailQuill) {
      emailQuill.clipboard.dangerouslyPasteHTML(body);
    } else {
      document.getElementById('emailPreviewBody').innerHTML = body;
    }
  }
  
  function replaceVariables(template, variables) {
    return template.replace(/\{\{(\w+)\}\}/g, (match, key) => variables[key] || match);
  }
  
  function calculateDaysUntil(appointmentTimeStr) {
    if (!appointmentTimeStr) return '?';
    
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
    let apptDate = new Date(year, monthIdx, day);
    
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const apptDateOnly = new Date(year, monthIdx, day);
    
    if (apptDateOnly < today) {
      apptDate = new Date(year + 1, monthIdx, day);
    }
    
    const diffTime = apptDate - today;
    const diffDays = Math.round(diffTime / (1000 * 60 * 60 * 24));
    return diffDays >= 0 ? diffDays : '?';
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
      
      // Fetch emails from Primary Applicant and Applicants
      const applicantIds = [];
      if (fields['Primary Applicant'] && fields['Primary Applicant'].length > 0) {
        applicantIds.push(...fields['Primary Applicant']);
      }
      if (fields['Applicants'] && fields['Applicants'].length > 0) {
        applicantIds.push(...fields['Applicants']);
      }
      
      if (applicantIds.length > 0) {
        // Fetch email addresses for all applicants
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
    }).getRecordById('Opportunities', opportunityId);
  }

  // --- SPOUSE LOGIC ---
  function renderSpouseSection(f) {
     const statusEl = document.getElementById('spouseStatusText');
     const historyList = document.getElementById('spouseHistoryList');
     const linkEl = document.getElementById('spouseEditLink');

     const spouseName = (f['Spouse Name'] && f['Spouse Name'].length > 0) ? f['Spouse Name'][0] : null;
     const spouseId = (f['Spouse'] && f['Spouse'].length > 0) ? f['Spouse'][0] : null;

     if (spouseName && spouseId) {
        statusEl.innerHTML = `Spouse: <a class="data-link" onclick="loadPanelRecord('Contacts', '${spouseId}')">${spouseName}</a>`;
        linkEl.innerText = "Edit"; 
     } else {
        statusEl.innerHTML = "Single";
        linkEl.innerText = "Edit"; 
     }

     linkEl.style.display = 'inline'; 

     historyList.innerHTML = '';
     const rawLogs = f['Spouse History Text']; 

     if (rawLogs && Array.isArray(rawLogs) && rawLogs.length > 0) {
        const parsedLogs = rawLogs.map(parseSpouseHistoryEntry).filter(Boolean);
        parsedLogs.sort((a, b) => b.timestamp - a.timestamp);
        
        const showLimit = 3;
        const initialSet = parsedLogs.slice(0, showLimit);
        initialSet.forEach(entry => { renderHistoryItem(entry, historyList); });
        if (parsedLogs.length > showLimit) {
           const remaining = parsedLogs.slice(showLimit);
           const expandLink = document.createElement('div');
           expandLink.className = 'expand-link';
           expandLink.innerText = `Show ${remaining.length} older records...`;
           expandLink.onclick = function() {
              remaining.forEach(entry => { renderHistoryItem(entry, historyList); });
              expandLink.style.display = 'none'; 
           };
           historyList.appendChild(expandLink);
        }
     } else {
        historyList.innerHTML = '<li class="spouse-history-item" style="border:none;">No history recorded.</li>';
     }
  }
  
  function parseSpouseHistoryEntry(logString) {
     const match = logString.match(/^(\d{4})-(\d{2})-(\d{2})\s+(\d{2}):(\d{2}):(\d{2}):\s*(connected as spouse to|disconnected as spouse from)\s+(.+)$/);
     if (!match) return null;
     const [, year, month, day, hours, mins, secs, action, spouseName] = match;
     const timestamp = new Date(parseInt(year), parseInt(month) - 1, parseInt(day), parseInt(hours), parseInt(mins), parseInt(secs));
     const displayDate = `${day}/${month}/${year}`;
     const displayText = `${action} ${spouseName}`;
     return { timestamp, displayDate, displayText };
  }

  function renderHistoryItem(entry, container) {
     const li = document.createElement('li');
     li.className = 'spouse-history-item';
     li.innerText = `${entry.displayDate}: ${entry.displayText}`;
     const expandLink = container.querySelector('.expand-link');
     if(expandLink) { container.insertBefore(li, expandLink); } else { container.appendChild(li); }
  }

  // --- INLINE EDIT LOGIC ---
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
    document.getElementById('recordId').value = ""; enableEditMode();
    document.getElementById('formTitle').innerText = "New Contact";
    document.getElementById('formSubtitle').innerText = '';
    document.getElementById('submitBtn').innerText = "Save Contact";
    document.getElementById('cancelBtn').style.display = 'inline-block';
    document.getElementById('editBtn').style.visibility = 'hidden';
    document.getElementById('oppList').innerHTML = '<li style="color:#CCC; font-size:12px; font-style:italic;">No opportunities linked.</li>';
    document.getElementById('auditSection').style.display = 'none';
    document.getElementById('duplicateWarningBox').style.display = 'none';
    document.getElementById('spouseStatusText').innerHTML = "Single";
    document.getElementById('spouseHistoryList').innerHTML = "";
    document.getElementById('spouseEditLink').style.display = 'inline';
    document.getElementById('refreshBtn').style.display = 'none';
    closeOppPanel();
  }
  function handleSearch(event) {
    const query = event.target.value; const status = document.getElementById('searchStatus');
    clearTimeout(loadingTimer); 
    if(query.length === 0) { status.innerText = ""; loadContacts(); return; }
    clearTimeout(searchTimeout); status.innerText = "Typing...";
    searchTimeout = setTimeout(() => {
      status.innerText = "Searching...";
      google.script.run.withSuccessHandler(function(records) {
         status.innerText = records.length > 0 ? `Found ${records.length} matches` : "No matches found";
         renderList(records);
      }).searchContacts(query);
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

    google.script.run.withSuccessHandler(function(records) {
         clearTimeout(loadingTimer); document.getElementById('loading').style.display = 'none';
         if (!records || records.length === 0) { 
           list.innerHTML = '<li style="padding:20px; color:#999; text-align:center; font-size:13px;">No contacts found</li>'; 
           return; 
         }
         renderList(records);
      }).getRecentContacts();
  }
  function renderList(records) {
    const list = document.getElementById('contactList'); 
    document.getElementById('loading').style.display = 'none'; 
    list.innerHTML = '';
    records.forEach(record => {
      const f = record.fields; const item = document.createElement('li'); item.className = 'contact-item';
      const fullName = formatName(f);
      const initials = getInitials(f.FirstName, f.LastName);
      const avatarColor = getAvatarColor(fullName);
      const modifiedTooltip = formatModifiedTooltip(f);
      const modifiedShort = formatModifiedShort(f);
      item.innerHTML = `<div class="contact-avatar" style="background-color:${avatarColor}">${initials}</div><div class="contact-info"><span class="contact-name">${fullName}</span><div class="contact-details-row">${formatDetailsRow(f)}</div></div>${modifiedShort ? `<span class="contact-modified" title="${modifiedTooltip || ''}">${modifiedShort}</span>` : ''}`;
      item.onclick = function() { selectContact(record); }; list.appendChild(item);
    });
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
  function formatModifiedTooltip(f) {
    const modifiedOn = f['Modified On'];
    const modifiedBy = f['Last Site User Name'];
    if (!modifiedOn) return null;
    
    const dateMatch = modifiedOn.match(/(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/);
    if (!dateMatch) return null;
    
    const modDate = new Date(dateMatch[1], dateMatch[2] - 1, dateMatch[3], dateMatch[4], dateMatch[5]);
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
    const modifiedOn = f['Modified On'];
    if (!modifiedOn) return null;
    
    const dateMatch = modifiedOn.match(/(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/);
    if (!dateMatch) return null;
    
    const modDate = new Date(dateMatch[1], dateMatch[2] - 1, dateMatch[3], dateMatch[4], dateMatch[5]);
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
    const preferredName = f.PreferredName || f.FirstName || '';
    const tenure = calculateTenure(f.Created);
    if (!preferredName && !tenure) return '';
    const parts = [];
    if (preferredName) parts.push(`prefers ${preferredName}`);
    if (tenure) parts.push(`in our database for ${tenure}`);
    return parts.join(' · ');
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
    return "today";
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

  // --- CORRECTED HISTORY DATE LOGIC ---
  function renderHistory(f) {
    const section = document.getElementById('auditSection');
    section.innerHTML = ''; section.style.display = 'block';
    const createdStr = f.Created || "";
    const dateMatch = createdStr.match(/(\d{2}):(\d{2})\s+(\d{2})\/(\d{2})\/(\d{4})/);
    let durationText = "unavailable";

    if (dateMatch) {
       const hours = parseInt(dateMatch[1], 10);
       const minutes = parseInt(dateMatch[2], 10);
       const day = parseInt(dateMatch[3], 10);
       const month = parseInt(dateMatch[4], 10) - 1; 
       const year = parseInt(dateMatch[5], 10);
       const createdDate = new Date(year, month, day, hours, minutes);
       const now = new Date();
       const diffMs = now - createdDate;
       const diffMinsTotal = Math.floor(diffMs / (1000 * 60));
       const diffHoursTotal = Math.floor(diffMs / (1000 * 60 * 60));
       const diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));

       if (diffDays > 730) { const years = Math.floor(diffDays / 365); durationText = `over ${years} years`; } 
       else if (diffDays > 60) { const months = Math.floor(diffDays / 30); durationText = `over ${months} months`; } 
       else if (diffDays >= 1) { durationText = (diffDays === 1) ? "1 day" : `${diffDays} days`; } 
       else {
          const mins = diffMinsTotal % 60;
          const hStr = (diffHoursTotal === 1) ? "hr" : "hrs";
          durationText = `${diffHoursTotal} ${hStr} and ${mins} minutes`;
       }
    }
    if (f.Created) { const line1 = document.createElement('div'); line1.className = 'audit-modified'; line1.innerText = f.Created; section.appendChild(line1); }
    if (f.Modified) { const line2 = document.createElement('div'); line2.className = 'audit-modified'; line2.innerText = f.Modified; section.appendChild(line2); }
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
         li.innerHTML = `<span class="opp-title">${name}${typeLabel}</span><span class="opp-role-wrapper">${statusBadge}<span class="opp-role">${role}</span></span>`;
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
      mobilePhone: formObject.mobilePhone.value,
      email1: formObject.email1.value,
      description: formObject.description.value
    };
    google.script.run.withSuccessHandler(function(response) {
         status.innerText = "✅ " + response; status.className = "status-success";
         loadContacts(); if(!formData.recordId) resetForm();
         btn.disabled = false; btn.innerText = "Update Contact"; disableEditMode(); 
         setTimeout(() => { status.innerText = ""; status.className = ""; }, 3000);
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
          else item.value.forEach(link => { linkHtml += `<a class="data-link" onclick="event.stopPropagation(); loadPanelRecord('${link.table}', '${link.id}')">${link.name}</a>`; });
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
      
      // Audit section
      if (response.audit && (response.audit.Created || response.audit.Modified)) {
        let auditHtml = '<div class="panel-audit-section">';
        if (response.audit.Created) auditHtml += `<div>${response.audit.Created}</div>`;
        if (response.audit.Modified) auditHtml += `<div>${response.audit.Modified}</div>`;
        auditHtml += '</div>';
        html += auditHtml;
      }
      
      // For Opportunities, use smart layout
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
        html += `<div style="padding:8px 0;"><button type="button" class="btn-add-appointment" style="padding:6px 14px; background:#7B8B64; color:#F2F0E9; border:none; border-radius:4px; cursor:pointer; font-size:13px; font-weight:600;" onclick="openAppointmentForm('${id}')">+ Add Appointment</button></div>`;
        html += `</div>`;
        
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
  
  // Helper function to render editable appointment fields without edit icon (for Notes)
  function renderApptFieldNoIcon(apptId, label, fieldKey, value, type) {
    const displayValue = value || '';
    const escaped = (value || '').replace(/"/g, '&quot;');
    const valueHtml = `<div class="detail-value appt-editable appt-notes-field" onclick="editApptField('${apptId}', '${fieldKey}', '${type}')" data-appt-id="${apptId}" data-field="${fieldKey}" data-value="${escaped}" style="white-space:pre-wrap; min-height:60px; padding:8px; border:1px solid #ddd; border-radius:4px; cursor:text;">${displayValue}</div>`;
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
      inputHtml = `<textarea class="inline-edit-input" rows="3" onblur="saveApptField('${apptId}', '${fieldKey}', this.value, '${type}')" onkeydown="if(event.key==='Escape'){cancelApptEdit('${apptId}', '${fieldKey}');}">${currentValue === '-' ? '' : currentValue}</textarea>`;
    } else {
      inputHtml = `<input type="text" class="inline-edit-input" value="${currentValue === '-' ? '' : currentValue}" onblur="saveApptField('${apptId}', '${fieldKey}', this.value, '${type}')" onkeydown="if(event.key==='Enter'){this.blur();}if(event.key==='Escape'){cancelApptEdit('${apptId}', '${fieldKey}');}">`;
    }
    
    valueSpan.outerHTML = inputHtml;
    const input = parent.querySelector('.inline-edit-input');
    if (input) input.focus();
  }
  
  // Save appointment field
  function saveApptField(apptId, fieldKey, value, type) {
    const opportunityId = document.getElementById('appointmentsContainer')?.dataset.opportunityId;
    
    // If setting appointment time and status is currently blank, auto-set to Scheduled
    if (fieldKey === 'appointmentTime' && value) {
      const statusEl = document.querySelector(`[data-appt-id="${apptId}"][data-field="appointmentStatus"]`);
      const currentStatus = statusEl?.querySelector('span')?.textContent || '';
      if (!currentStatus || currentStatus === 'Not Set' || currentStatus === '-') {
        // Also update status to Scheduled
        google.script.run.updateAppointment(apptId, 'appointmentStatus', 'Scheduled');
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
                container.innerHTML = '<div style="color:#888; padding:16px 16px 4px 16px; font-style:italic;">No appointments scheduled</div>';
              })
              .createAppointment(opportunityId, fields);
            return;
          }
          
          container.innerHTML = '<div style="color:#888; padding:16px 16px 4px 16px; font-style:italic;">No appointments scheduled</div>';
          return;
        }
        
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
          html += `<div id="appt_field_wrap_${appt.id}_phone" style="${phoneStyle}">${renderApptField(appt.id, 'Phone Number', 'phoneNumber', appt.phoneNumber, 'text')}</div>`;
          html += `<div id="appt_field_wrap_${appt.id}_video" style="${videoStyle}">${renderApptField(appt.id, 'Video Meet URL', 'videoMeetUrl', appt.videoMeetUrl, 'text')}</div>`;
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
          
          // Row 4: Conf Email Sent, Conf Text Sent
          html += `<div class="taco-row">`;
          html += renderApptCheckbox(appt.id, 'Conf Email Sent', 'confEmailSent', appt.confEmailSent);
          html += renderApptCheckbox(appt.id, 'Conf Text Sent', 'confTextSent', appt.confTextSent);
          html += `</div>`;
          
          // Row 5: Status (1/3) and Notes (2/3)
          html += `<div class="taco-row taco-row-status-notes" style="margin-top:15px;">`;
          html += renderApptField(appt.id, 'Status', 'appointmentStatus', status, 'select', ['', 'Scheduled', 'Completed', 'Cancelled', 'No Show']);
          html += `<div style="grid-column: span 2;">${renderApptFieldNoIcon(appt.id, 'Notes', 'notes', appt.notes, 'textarea')}</div>`;
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