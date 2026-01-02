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
  
  // Email template links (editable, saved to localStorage)
  const DEFAULT_EMAIL_LINKS = {
    officeMap: 'https://maps.app.goo.gl/qm2ohJP2j1t6GqCt9',
    ourTeam: 'https://stellaris.loans/our-team',
    factFind: 'https://drive.google.com/file/d/1_U6kKck5IA3TBtFdJEzyxs_XpvcKvg9s/view?usp=sharing',
    myGov: 'https://my.gov.au/',
    myGovVideo: 'https://www.youtube.com/watch?v=bSMs2XO1V7Y',
    incomeStatementInstructions: 'https://drive.google.com/file/d/1Y8B4zPLb_DTkV2GZnlGztm-HMfA3OWYP/view?usp=sharing'
  };
  
  function loadEmailLinks() {
    const saved = localStorage.getItem('emailLinks');
    if (saved) {
      try { return { ...DEFAULT_EMAIL_LINKS, ...JSON.parse(saved) }; }
      catch (e) { return { ...DEFAULT_EMAIL_LINKS }; }
    }
    return { ...DEFAULT_EMAIL_LINKS };
  }
  
  let EMAIL_LINKS = loadEmailLinks();
  
  function openEmailSettings() {
    document.getElementById('settingOfficeMap').value = EMAIL_LINKS.officeMap || '';
    document.getElementById('settingOurTeam').value = EMAIL_LINKS.ourTeam || '';
    document.getElementById('settingFactFind').value = EMAIL_LINKS.factFind || '';
    document.getElementById('settingMyGov').value = EMAIL_LINKS.myGov || '';
    document.getElementById('settingMyGovVideo').value = EMAIL_LINKS.myGovVideo || '';
    document.getElementById('settingIncomeInstructions').value = EMAIL_LINKS.incomeStatementInstructions || '';
    openModal('emailSettingsModal');
  }
  
  function closeEmailSettings() {
    closeModal('emailSettingsModal');
  }
  
  function saveEmailSettings() {
    EMAIL_LINKS.officeMap = document.getElementById('settingOfficeMap').value;
    EMAIL_LINKS.ourTeam = document.getElementById('settingOurTeam').value;
    EMAIL_LINKS.factFind = document.getElementById('settingFactFind').value;
    EMAIL_LINKS.myGov = document.getElementById('settingMyGov').value;
    EMAIL_LINKS.myGovVideo = document.getElementById('settingMyGovVideo').value;
    EMAIL_LINKS.incomeStatementInstructions = document.getElementById('settingIncomeInstructions').value;
    localStorage.setItem('emailLinks', JSON.stringify(EMAIL_LINKS));
    closeEmailSettings();
    updateEmailPreview();
    showAlert('Saved', 'Email template links updated', 'success');
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
    
    document.getElementById('emailPreviewBody').innerHTML = body;
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
    if (label) label.innerText = 'Saving...';
    google.script.run.withSuccessHandler(function(res) {
      if (label) label.innerText = isChecked ? 'Yes' : 'No';
      // Toggle appointment fields visibility
      if (fieldKey === 'Taco: Converted to Appt') {
        const section = document.getElementById('apptFieldsSection');
        const notice = document.getElementById('apptCollapsedNotice');
        if (section) section.style.display = isChecked ? '' : 'none';
        if (notice) notice.style.display = isChecked ? '' : 'none';
      }
    }).withFailureHandler(function(err) {
      const input = document.getElementById('input_' + fieldKey);
      if (input) input.checked = !isChecked;
      if (label) label.innerText = !isChecked ? 'Yes' : 'No';
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
     document.getElementById('view_' + key).style.display = 'block';
     document.getElementById('edit_' + key).style.display = 'none';
     pendingLinkedEdits[key] = [];
     pendingRemovals = {};
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

    // --- RESTORED CORRECT MESSAGE ---
    loadingTimer = setTimeout(() => { 
       loadingDiv.innerHTML = `
         <div style="margin-top:15px; text-align:center;">
           <button onclick="loadContacts()" class="wake-btn">Wake up Google!</button>
           <p class="wake-note">Google isn't constantly awake in the background waiting for us to use this site, so it goes to sleep if we haven't used it for a little while. You might need to hit the button a few times to get it to pay attention. One day we'll make it work differently so that this problem goes away.</p>
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
      response.data.forEach(item => { if(item.type === 'link') currentPanelData[item.key] = item.value; });

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
          const noLabel = item.noLabel || 'No';
          return `<div class="detail-group${tacoClass}"><div class="detail-label">${item.label}</div><div class="checkbox-field"><input type="checkbox" id="input_${item.key}" ${checkedAttr} onchange="saveCheckboxField('${tbl}', '${recId}', '${item.key}', this.checked)"><label for="input_${item.key}">${isChecked ? 'Yes' : noLabel}</label></div></div>`;
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
          return `<div class="detail-group${tacoClass}"><div class="detail-label">${item.label}</div><div id="view_${item.key}" onclick="toggleLinkedEdit('${item.key}')" class="editable-field"><div class="detail-value" style="display:flex; justify-content:space-between; align-items:center;"><span>${linkHtml}</span><span class="edit-field-icon">✎</span></div></div><div id="edit_${item.key}" style="display:none;"><div class="edit-wrapper"><div id="chip_container_${item.key}" class="link-chip-container"></div><input type="text" placeholder="Add contact..." class="link-search-input" onkeyup="handleLinkedSearch(event, '${item.key}')"><div id="error_${item.key}" class="input-error"></div><div id="results_${item.key}" class="link-results"></div><div class="edit-btn-row" style="margin-top:10px;"><button onclick="closeLinkedEdit('${item.key}')" class="btn-cancel-field">Done</button></div></div></div></div>`;
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
          html += `<div class="taco-section-header"><img src="https://taco.insightprocessing.com.au/static/images/taco.jpg" alt="Taco"><span>Taco fields</span></div>`;
          html += '<div id="tacoFieldsContainer">';
          
          // Get current values for conditional logic
          const convertedToAppt = dataMap['Taco: Converted to Appt']?.value === true || dataMap['Taco: Converted to Appt']?.value === 'true';
          const typeOfAppt = dataMap['Taco: Type of Appointment']?.value || '';
          const howBooked = dataMap['Taco: How appt booked']?.value || '';
          
          // Check if appointment is in the past (at least 1 day after)
          const apptTimeStr = dataMap['Taco: Appointment Time']?.value || '';
          let apptIsPast = false;
          if (apptTimeStr && convertedToAppt) {
            const match = apptTimeStr.match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
            if (match) {
              const day = parseInt(match[1]);
              const month = parseInt(match[2]) - 1;
              let year = parseInt(match[3]);
              if (year < 100) year += 2000;
              const apptDate = new Date(year, month, day);
              apptDate.setHours(23, 59, 59, 999);
              const today = new Date();
              today.setHours(0, 0, 0, 0);
              apptIsPast = today > apptDate;
            }
          }
          
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
          
          // Appointment fields (only if Converted to Appt is checked)
          const showApptFields = convertedToAppt && !apptIsPast;
          
          // Show collapsed notice ABOVE appointment fields if appointment is past
          if (convertedToAppt && apptIsPast) {
            const collapsedText = `Appointment details hidden (${apptTimeStr} has passed)`;
            html += `<div id="apptCollapsedNotice" class="appt-collapsed-notice" onclick="togglePastApptFields()">
              <span class="chevron">▶</span>
              <span id="apptNoticeText" data-collapsed-text="${collapsedText.replace(/"/g, '&quot;')}">${collapsedText}</span>
            </div>`;
          }
          
          html += `<div id="apptFieldsSection" style="${showApptFields ? '' : 'display:none;'}">`;
          
          // Row 5: Appointment Time, Type of Appointment, How Appt Booked
          html += '<div class="taco-row">';
          if (dataMap['Taco: Appointment Time']) html += renderField(dataMap['Taco: Appointment Time'], table, id);
          if (dataMap['Taco: Type of Appointment']) html += renderField(dataMap['Taco: Type of Appointment'], table, id);
          if (dataMap['Taco: How appt booked']) html += renderField(dataMap['Taco: How appt booked'], table, id);
          html += '</div>';
          
          // Row 6: Appt Phone Number (if Phone), Appt Meet URL (if Video), How Appt Booked Other (if Other)
          html += '<div class="taco-row">';
          const phoneDisplay = typeOfAppt === 'Phone' ? '' : 'display:none;';
          const videoDisplay = typeOfAppt === 'Video' ? '' : 'display:none;';
          const otherDisplay = howBooked === 'Other' ? '' : 'display:none;';
          if (dataMap['Taco: Appt Phone Number']) html += `<div id="field_wrap_Taco: Appt Phone Number" style="${phoneDisplay}">${renderField(dataMap['Taco: Appt Phone Number'], table, id)}</div>`;
          if (dataMap['Taco: Appt Meet URL']) html += `<div id="field_wrap_Taco: Appt Meet URL" style="${videoDisplay}">${renderField(dataMap['Taco: Appt Meet URL'], table, id)}</div>`;
          if (dataMap['Taco: How Appt Booked Other']) html += `<div id="field_wrap_Taco: How Appt Booked Other" style="${otherDisplay}">${renderField(dataMap['Taco: How Appt Booked Other'], table, id)}</div>`;
          html += '</div>';
          
          // Row 7: Need Evidence in Advance, Need Appt Reminder
          html += '<div class="taco-row">';
          if (dataMap['Taco: Need Evidence in Advance']) html += renderField(dataMap['Taco: Need Evidence in Advance'], table, id);
          if (dataMap['Taco: Need Appt Reminder']) {
            // Custom label if Calendly
            const reminderField = { ...dataMap['Taco: Need Appt Reminder'] };
            if (howBooked === 'Calendly') {
              reminderField.noLabel = 'Not required as Calendly will do it automatically';
            }
            html += renderField(reminderField, table, id);
          }
          html += '</div>';
          
          // Row 8: Appt Conf Email Sent, Appt Conf Text Sent
          html += '<div class="taco-row">';
          if (dataMap['Taco: Appt Conf Email Sent']) html += renderField(dataMap['Taco: Appt Conf Email Sent'], table, id);
          if (dataMap['Taco: Appt Conf Text Sent']) html += renderField(dataMap['Taco: Appt Conf Text Sent'], table, id);
          html += '</div>';
          
          html += '</div>'; // close apptFieldsSection
          
          html += '</div></div>'; // close tacoFieldsContainer and taco-section-box
        }
        
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
        
        // Action buttons
        const safeName = (response.title || '').replace(/'/g, "\\'");
        html += `<div style="margin-top:30px; padding-top:20px; border-top:1px solid #EEE;">`;
        html += `<button type="button" class="btn-confirm" style="width:100%; margin-bottom:10px;" onclick="openEmailComposerFromPanel('${id}')">Send Confirmation Email</button>`;
        html += `<button type="button" class="btn-delete" onclick="confirmDeleteOpportunity('${id}', '${safeName}')">Delete Opportunity</button>`;
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