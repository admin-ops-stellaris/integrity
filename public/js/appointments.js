(function() {
  'use strict';
  
  const state = window.IntegrityState;
  
  // Helper function to render editable appointment fields (label above value like Taco fields)
  function renderApptField(apptId, label, fieldKey, value, type, options = []) {
    const displayValue = value || '-';
    let valueHtml = '';
    
    if (type === 'datetime') {
      const formatted = formatDatetimeForDisplay(value);
      const escapedValue = (value || '').replace(/'/g, "\\'");
      // Datetime editing opens mini-popover with smart-date and smart-time fields
      valueHtml = `<div class="detail-value appt-editable" onclick="editApptTimeInline('${apptId}', '${escapedValue}')" data-appt-id="${apptId}" data-field="${fieldKey}" data-value="${value || ''}" style="display:flex; justify-content:space-between; align-items:center;"><span>${formatted}</span><span class="edit-field-icon">✎</span></div>`;
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
      const apptTime = new Date(value);
      const now = new Date();
      if (apptTime > now) {
        const statusEl = document.querySelector(`[data-appt-id="${apptId}"][data-field="appointmentStatus"]`);
        const statusSpan = statusEl?.querySelector('span');
        const currentStatus = statusSpan?.textContent?.trim() || '';
        if (!currentStatus || currentStatus === 'Not Set' || currentStatus === '-' || currentStatus === '') {
          google.script.run.updateAppointment(apptId, 'appointmentStatus', 'Scheduled');
        }
      }
    }
    
    google.script.run
      .withSuccessHandler(function() {
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
  
  // Cancel appointment edit
  function cancelApptEdit(apptId, fieldKey) {
    const opportunityId = document.getElementById('appointmentsContainer')?.dataset.opportunityId;
    if (opportunityId) loadAppointmentsForOpportunity(opportunityId);
  }
  
  // Edit appointment date/time inline with mini-popover
  function editApptTimeInline(apptId, currentValue) {
    // Close any existing popover
    closeApptTimePopover();
    
    const triggerEl = document.querySelector(`[data-appt-id="${apptId}"][data-field="appointmentTime"]`);
    if (!triggerEl) return;
    
    // Parse current datetime value using local time to match existing display logic
    let dateDisplay = '';
    let timeDisplay = '';
    let isoDate = '';
    let time24 = '';
    
    if (currentValue) {
      // Use formatDatetimeForInput to get consistent ISO format
      const isoInput = formatDatetimeForInput(currentValue);
      if (isoInput) {
        // isoInput is YYYY-MM-DDTHH:MM format
        const parts = isoInput.split('T');
        if (parts.length === 2) {
          isoDate = parts[0]; // YYYY-MM-DD
          time24 = parts[1];  // HH:MM
          
          // Convert to display formats
          const [year, month, day] = isoDate.split('-');
          dateDisplay = `${day}/${month}/${year}`;
          
          const [h, m] = time24.split(':').map(Number);
          const ampm = h >= 12 ? 'PM' : 'AM';
          const displayHours = h % 12 || 12;
          timeDisplay = `${displayHours}:${String(m).padStart(2, '0')} ${ampm}`;
        }
      }
    }
    
    // Create popover HTML
    const popover = document.createElement('div');
    popover.className = 'appt-time-popover';
    popover.id = 'apptTimePopover';
    popover.dataset.apptId = apptId;
    popover.innerHTML = `
      <div class="appt-time-popover-content">
        <div class="appt-time-popover-row">
          <label>Date</label>
          <input type="text" class="smart-date appt-popover-date" 
                 value="${dateDisplay}" 
                 data-iso-date="${isoDate}"
                 placeholder="DD/MM/YYYY">
        </div>
        <div class="appt-time-popover-row">
          <label>Time</label>
          <input type="text" class="smart-time appt-popover-time" 
                 value="${timeDisplay}" 
                 data-time24="${time24}"
                 placeholder="e.g. 1300">
        </div>
        <div class="appt-time-popover-actions">
          <button type="button" class="appt-popover-save" onclick="saveApptTimePopover()">✓</button>
          <button type="button" class="appt-popover-cancel" onclick="closeApptTimePopover()">✕</button>
        </div>
      </div>
    `;
    
    // Position near the trigger element
    document.body.appendChild(popover);
    const rect = triggerEl.getBoundingClientRect();
    popover.style.position = 'fixed';
    popover.style.top = (rect.bottom + 4) + 'px';
    popover.style.left = rect.left + 'px';
    popover.style.zIndex = '10000';
    
    // Adjust if off-screen
    const popoverRect = popover.getBoundingClientRect();
    if (popoverRect.right > window.innerWidth) {
      popover.style.left = (window.innerWidth - popoverRect.width - 10) + 'px';
    }
    if (popoverRect.bottom > window.innerHeight) {
      popover.style.top = (rect.top - popoverRect.height - 4) + 'px';
    }
    
    // Focus the date field
    const dateInput = popover.querySelector('.appt-popover-date');
    if (dateInput) {
      dateInput.focus();
      dateInput.select();
    }
    
    // Add event listeners for keyboard, outside click, and focus tracking
    setTimeout(() => {
      document.addEventListener('click', handlePopoverOutsideClick);
      document.addEventListener('keydown', handlePopoverKeydown);
      popover.addEventListener('focusout', handlePopoverFocusOut);
    }, 10);
  }
  
  function handlePopoverOutsideClick(e) {
    const popover = document.getElementById('apptTimePopover');
    if (popover && !popover.contains(e.target)) {
      closeApptTimePopover();
    }
  }
  
  function handlePopoverFocusOut(e) {
    const popover = document.getElementById('apptTimePopover');
    if (!popover) return;
    
    // Check if the new focus target is still within the popover
    // relatedTarget is where focus is going to
    if (e.relatedTarget && popover.contains(e.relatedTarget)) {
      // Focus is moving within the popover (e.g., Tab from date to time)
      return;
    }
    
    // Small delay to allow button clicks to register before closing
    setTimeout(() => {
      if (document.activeElement && popover.contains(document.activeElement)) {
        return;
      }
      // Focus moved outside popover - close it
      closeApptTimePopover();
    }, 100);
  }
  
  function handlePopoverKeydown(e) {
    const popover = document.getElementById('apptTimePopover');
    if (!popover) return;
    
    if (e.key === 'Escape') {
      closeApptTimePopover();
    } else if (e.key === 'Enter') {
      e.preventDefault();
      saveApptTimePopover();
    } else if (e.key === 'Tab' && !e.shiftKey) {
      // Allow tab between date and time fields
      const timeInput = popover.querySelector('.appt-popover-time');
      if (document.activeElement === timeInput) {
        e.preventDefault();
        saveApptTimePopover();
      }
    }
  }
  
  function saveApptTimePopover() {
    const popover = document.getElementById('apptTimePopover');
    if (!popover) return;
    
    const apptId = popover.dataset.apptId;
    const dateInput = popover.querySelector('.appt-popover-date');
    const timeInput = popover.querySelector('.appt-popover-time');
    
    // Parse values directly in case change event hasn't fired yet
    let isoDate = dateInput?.dataset.isoDate || '';
    let time24 = timeInput?.dataset.time24 || '';
    
    // If dataset not set, try parsing the raw value
    if (!isoDate && dateInput?.value) {
      const parsed = window.parseFlexibleDate(dateInput.value);
      if (parsed) {
        isoDate = parsed.iso;
        dateInput.value = parsed.display;
        dateInput.dataset.isoDate = parsed.iso;
      }
    }
    if (!time24 && timeInput?.value) {
      const parsed = window.parseFlexibleTime(timeInput.value);
      if (parsed) {
        time24 = parsed.value24;
        timeInput.value = parsed.display;
        timeInput.dataset.time24 = parsed.value24;
      }
    }
    
    // Validate: require at least a date
    if (!isoDate) {
      if (dateInput?.value) {
        dateInput.style.borderColor = '#dc2626';
        dateInput.focus();
        return; // Keep popover open for correction
      }
      closeApptTimePopover();
      return;
    }
    
    // Build ISO datetime with timezone offset to preserve local time
    // Perth is UTC+8, but we get the actual local offset dynamically
    const now = new Date();
    const offsetMinutes = -now.getTimezoneOffset(); // getTimezoneOffset returns minutes behind UTC (negative for ahead)
    const offsetHours = Math.floor(Math.abs(offsetMinutes) / 60);
    const offsetMins = Math.abs(offsetMinutes) % 60;
    const offsetSign = offsetMinutes >= 0 ? '+' : '-';
    const tzOffset = `${offsetSign}${String(offsetHours).padStart(2, '0')}:${String(offsetMins).padStart(2, '0')}`;
    
    let appointmentTimeISO = '';
    if (isoDate && time24) {
      appointmentTimeISO = `${isoDate}T${time24}:00${tzOffset}`;
    } else if (isoDate) {
      appointmentTimeISO = `${isoDate}T00:00:00${tzOffset}`;
    }
    
    // Delegate to saveApptField for consistent status auto-update behavior
    closeApptTimePopover();
    saveApptField(apptId, 'appointmentTime', appointmentTimeISO, 'datetime');
  }
  
  function closeApptTimePopover() {
    const popover = document.getElementById('apptTimePopover');
    if (popover) {
      popover.remove();
    }
    document.removeEventListener('click', handlePopoverOutsideClick);
    document.removeEventListener('keydown', handlePopoverKeydown);
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
          const isTruthyCheckbox = (val) => {
            if (val === true || val === 1) return true;
            if (typeof val === 'string') {
              const lower = val.trim().toLowerCase();
              return ['true', 'yes', 'checked', '1'].includes(lower);
            }
            return Boolean(val);
          };
          
          const rawConverted = state.currentPanelData['Taco: Converted to Appt'];
          let convertedVal = (typeof rawConverted === 'object' && rawConverted !== null) 
            ? rawConverted.value 
            : rawConverted;
          
          if (isTruthyCheckbox(convertedVal)) {
            console.log('Legacy backfill: Converted to Appt is true but no appointments exist - creating from Taco fields');
            container.innerHTML = '<div style="color:#888; padding:16px 16px 4px 16px; font-style:italic;">Migrating appointment data...</div>';
            
            const getVal = (key) => {
              const v = state.currentPanelData[key];
              if (v === undefined || v === null) return null;
              if (typeof v === 'object' && v.value !== undefined) return v.value;
              return v;
            };
            const getBool = (key) => {
              const v = getVal(key);
              return isTruthyCheckbox(v);
            };
            
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
        
        window.currentOpportunityAppointments = appointments;
        
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
          
          const isExpanded = status === 'Scheduled';
          const expandedClass = isExpanded ? 'expanded' : '';
          
          html += `<div class="appointment-item subsequent-appt ${expandedClass}" data-appt-id="${appt.id}">`;
          
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
          
          html += `<div class="appointment-item-body">`;
          html += `<div class="appointment-item-divider"></div>`;
          
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
          
          html += `<div class="appt-section appt-section-1">`;
          
          html += `<div class="taco-row">`;
          html += renderApptField(appt.id, 'Appointment Time', 'appointmentTime', appt.appointmentTime, 'datetime');
          html += renderApptField(appt.id, 'Type of Appointment', 'typeOfAppointment', appt.typeOfAppointment, 'select', ['Phone', 'Video', 'Office']);
          html += renderApptField(appt.id, 'How Booked', 'howBooked', appt.howBooked, 'select', ['Calendly', 'Email', 'Phone', 'Podium', 'Other']);
          html += `</div>`;
          
          html += `<div class="taco-row">`;
          const phoneStyle = appt.typeOfAppointment === 'Phone' ? '' : 'display:none;';
          const videoStyle = appt.typeOfAppointment === 'Video' ? '' : 'display:none;';
          const otherStyle = appt.howBooked === 'Other' ? '' : 'display:none;';
          html += `<div id="appt_field_wrap_${appt.id}_phone" style="${phoneStyle}">${renderApptFieldNoIcon(appt.id, 'Phone Number', 'phoneNumber', appt.phoneNumber, 'text')}</div>`;
          html += `<div id="appt_field_wrap_${appt.id}_video" style="${videoStyle}">${renderApptFieldNoIcon(appt.id, 'Video Meet URL', 'videoMeetUrl', appt.videoMeetUrl, 'text')}</div>`;
          html += `<div id="appt_field_wrap_${appt.id}_other" style="${otherStyle}">${renderApptField(appt.id, 'How Booked Other', 'howBookedOther', appt.howBookedOther, 'text')}</div>`;
          html += `</div>`;
          
          html += `<div class="taco-row">`;
          html += renderApptCheckbox(appt.id, 'Need Evidence in Advance', 'needEvidenceInAdvance', appt.needEvidenceInAdvance);
          html += renderApptCheckbox(appt.id, 'Need Appt Reminder', 'needApptReminder', appt.needApptReminder);
          html += `</div>`;
          
          html += `<div style="margin:15px 0;"><button type="button" class="btn-confirm btn-inline" onclick="openEmailComposerFromPanel('${opportunityId}')">Send Confirmation Email</button></div>`;
          
          html += `</div>`;
          
          html += `<div class="appt-section appt-section-2">`;
          
          html += `<div class="taco-row">`;
          html += renderApptCheckbox(appt.id, 'Conf Email Sent', 'confEmailSent', appt.confEmailSent);
          html += renderApptCheckbox(appt.id, 'Conf Text Sent', 'confTextSent', appt.confTextSent);
          html += renderApptField(appt.id, 'Appointment Status', 'appointmentStatus', status, 'select', ['', 'Scheduled', 'Completed', 'Cancelled', 'No Show']);
          html += `</div>`;
          
          html += `<div style="margin-top:12px;">`;
          html += renderApptFieldNoIcon(appt.id, 'Notes', 'notes', appt.notes, 'textarea');
          html += `</div>`;
          
          html += `</div>`;
          html += `</div>`;
          html += `</div>`;
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
    state.currentAppointmentOpportunityId = opportunityId;
    state.editingAppointmentId = appointment ? appointment.id : null;
    
    const modal = document.getElementById('appointmentFormModal');
    const title = document.getElementById('appointmentFormTitle');
    title.textContent = appointment ? 'Edit Appointment' : 'New Appointment';
    
    // Split datetime into separate date and time fields
    const dateEl = document.getElementById('apptFormDate');
    const timeEl = document.getElementById('apptFormTime');
    
    if (appointment?.appointmentTime) {
      const dt = new Date(appointment.appointmentTime);
      if (!isNaN(dt.getTime())) {
        // Format date as DD/MM/YYYY for display
        const dd = String(dt.getDate()).padStart(2, '0');
        const mm = String(dt.getMonth() + 1).padStart(2, '0');
        const yyyy = dt.getFullYear();
        dateEl.value = `${dd}/${mm}/${yyyy}`;
        dateEl.dataset.isoDate = `${yyyy}-${mm}-${dd}`;
        
        // Format time as h:mm AM/PM for display
        let hours = dt.getHours();
        const mins = String(dt.getMinutes()).padStart(2, '0');
        const ampm = hours < 12 ? 'AM' : 'PM';
        const displayHour = hours % 12 || 12;
        timeEl.value = `${displayHour}:${mins} ${ampm}`;
        timeEl.dataset.time24 = `${String(hours).padStart(2, '0')}:${mins}`;
      } else {
        dateEl.value = '';
        timeEl.value = '';
        delete dateEl.dataset.isoDate;
        delete timeEl.dataset.time24;
      }
    } else {
      dateEl.value = '';
      timeEl.value = '';
      delete dateEl.dataset.isoDate;
      delete timeEl.dataset.time24;
    }
    
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
    
    modal.scrollTop = 0;
    const modalContent = modal.querySelector('.modal-content');
    if (modalContent) modalContent.scrollTop = 0;
    setTimeout(() => {
      const firstInput = document.getElementById('apptFormDate');
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
    state.currentAppointmentOpportunityId = null;
    state.editingAppointmentId = null;
  }
  
  function saveAppointment() {
    if (!state.currentAppointmentOpportunityId) return;
    
    // Combine date and time into ISO string
    const dateEl = document.getElementById('apptFormDate');
    const timeEl = document.getElementById('apptFormTime');
    
    let appointmentTimeISO = '';
    const isoDate = dateEl.dataset.isoDate || parseDateInput(dateEl.value, dateEl);
    const time24 = timeEl.dataset.time24 || (function() {
      const parsed = window.parseFlexibleTime(timeEl.value);
      return parsed ? parsed.value24 : null;
    })();
    
    // Build ISO datetime with timezone offset to preserve local time
    const now = new Date();
    const offsetMinutes = -now.getTimezoneOffset();
    const offsetHours = Math.floor(Math.abs(offsetMinutes) / 60);
    const offsetMins = Math.abs(offsetMinutes) % 60;
    const offsetSign = offsetMinutes >= 0 ? '+' : '-';
    const tzOffset = `${offsetSign}${String(offsetHours).padStart(2, '0')}:${String(offsetMins).padStart(2, '0')}`;
    
    if (isoDate && time24) {
      appointmentTimeISO = `${isoDate}T${time24}:00${tzOffset}`;
    } else if (isoDate) {
      appointmentTimeISO = `${isoDate}T00:00:00${tzOffset}`;
    }
    
    const fields = {
      "Appointment Time": appointmentTimeISO,
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
    
    const oppId = state.currentAppointmentOpportunityId;
    
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
    
    if (state.editingAppointmentId) {
      google.script.run
        .withSuccessHandler(onSaveComplete)
        .withFailureHandler(onSaveError)
        .updateAppointmentFields(state.editingAppointmentId, fields);
    } else {
      google.script.run
        .withSuccessHandler(onSaveComplete)
        .withFailureHandler(onSaveError)
        .createAppointment(state.currentAppointmentOpportunityId, fields);
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
  
  // Expose functions to window for onclick handlers and opportunities.js usage
  window.renderApptField = renderApptField;
  window.renderApptFieldNoIcon = renderApptFieldNoIcon;
  window.renderApptCheckbox = renderApptCheckbox;
  window.editApptField = editApptField;
  window.saveApptField = saveApptField;
  window.updateApptCheckbox = updateApptCheckbox;
  window.autoResizeTextarea = autoResizeTextarea;
  window.cancelApptEdit = cancelApptEdit;
  window.editApptTimeInline = editApptTimeInline;
  window.saveApptTimePopover = saveApptTimePopover;
  window.closeApptTimePopover = closeApptTimePopover;
  window.loadAppointmentsForOpportunity = loadAppointmentsForOpportunity;
  window.formatDatetimeForInput = formatDatetimeForInput;
  window.formatDatetimeForDisplay = formatDatetimeForDisplay;
  window.openAppointmentForm = openAppointmentForm;
  window.updateAppointmentFormVisibility = updateAppointmentFormVisibility;
  window.closeAppointmentForm = closeAppointmentForm;
  window.saveAppointment = saveAppointment;
  window.toggleAppointmentExpand = toggleAppointmentExpand;
  window.editAppointment = editAppointment;
  window.deleteAppointment = deleteAppointment;
  
})();
