/**
 * Appointments Module
 * Appointment management within opportunities
 */
(function() {
  'use strict';
  
  const state = window.IntegrityState;
  
  // ============================================================
  // Load Appointments for Opportunity
  // ============================================================
  
  window.loadAppointmentsForOpportunity = function(opportunityId) {
    google.script.run.withSuccessHandler(function(appointments) {
      renderAppointments(appointments || [], opportunityId);
    }).getAppointmentsForOpportunity(opportunityId);
  };
  
  // ============================================================
  // Render Appointments
  // ============================================================
  
  function renderAppointments(appointments, opportunityId) {
    const container = document.getElementById('appointmentsSection');
    if (!container) return;
    
    const addBtn = `<span class="add-appointment-btn" onclick="openAppointmentForm('${opportunityId}')">+ Add Appointment</span>`;
    
    if (!appointments || appointments.length === 0) {
      container.innerHTML = `
        <div class="appointments-header">
          <span class="appointments-title">Appointments</span>
          ${addBtn}
        </div>
        <div class="appointments-empty">No appointments scheduled</div>
      `;
      return;
    }
    
    // Sort by date
    appointments.sort((a, b) => {
      const aDate = new Date(a.appointmentTime || 0);
      const bDate = new Date(b.appointmentTime || 0);
      return aDate - bDate;
    });
    
    container.innerHTML = `
      <div class="appointments-header">
        <span class="appointments-title">Appointments</span>
        ${addBtn}
      </div>
      <div class="appointments-list">
        ${appointments.map(appt => renderAppointmentItem(appt, opportunityId)).join('')}
      </div>
    `;
  }
  
  function renderAppointmentItem(appt, opportunityId) {
    const timeDisplay = formatAppointmentTime(appt.appointmentTime);
    const daysUntil = calculateDaysUntil(appt.appointmentTime);
    const typeIcon = getAppointmentTypeIcon(appt.type);
    const statusClass = appt.status ? `appt-status-${appt.status.toLowerCase().replace(/\s+/g, '-')}` : '';
    
    return `
      <div class="appointment-item ${statusClass}" data-appt-id="${appt.id}">
        <div class="appointment-main" onclick="toggleAppointmentExpand('${appt.id}')">
          <span class="appointment-icon">${typeIcon}</span>
          <span class="appointment-time">${timeDisplay}</span>
          ${daysUntil ? `<span class="appointment-days">${daysUntil}</span>` : ''}
          <span class="appointment-expand-icon">&#9654;</span>
        </div>
        <div class="appointment-details" id="apptDetails_${appt.id}" style="display:none;">
          ${renderApptField(appt.id, 'Type', 'type', appt.type, 'select', ['Office', 'Phone', 'Video'])}
          ${renderApptField(appt.id, 'How Booked', 'howBooked', appt.howBooked, 'select', ['Calendly', 'Email', 'Phone', 'Podium', 'Other'])}
          ${appt.type === 'Phone' ? renderApptField(appt.id, 'Phone Number', 'phoneNumber', appt.phoneNumber, 'text') : ''}
          ${appt.type === 'Video' ? renderApptField(appt.id, 'Video Meet URL', 'videoMeetUrl', appt.videoMeetUrl, 'text') : ''}
          ${renderApptCheckbox(appt.id, 'Evidence Needed', 'evidenceNeeded', appt.evidenceNeeded)}
          ${renderApptCheckbox(appt.id, 'Reminder Sent', 'reminderSent', appt.reminderSent)}
          ${renderApptField(appt.id, 'Status', 'status', appt.status, 'select', ['Scheduled', 'Completed', 'No Show', 'Cancelled'])}
          ${renderApptField(appt.id, 'Notes', 'notes', appt.notes, 'textarea')}
          <div class="appointment-actions">
            <button class="btn-small btn-edit" onclick="editAppointment('${appt.id}', '${opportunityId}')">Edit</button>
            <button class="btn-small btn-delete" onclick="deleteAppointment('${appt.id}', '${opportunityId}')">Delete</button>
          </div>
        </div>
      </div>
    `;
  }
  
  // ============================================================
  // Render Appointment Fields
  // ============================================================
  
  function renderApptField(apptId, label, fieldKey, value, type, options = []) {
    const displayValue = value || '-';
    
    return `
      <div class="appt-field" data-appt-id="${apptId}" data-field-key="${fieldKey}">
        <label>${escapeHtml(label)}</label>
        <span class="appt-field-value" onclick="editApptField('${apptId}', '${fieldKey}', '${type}', ${escapeHtmlForAttr(JSON.stringify(options))})">${escapeHtml(displayValue)}</span>
      </div>
    `;
  }
  
  function renderApptFieldNoIcon(apptId, label, fieldKey, value, type) {
    return renderApptField(apptId, label, fieldKey, value, type, []);
  }
  
  function renderApptCheckbox(apptId, label, fieldKey, checked) {
    return `
      <div class="appt-field appt-checkbox" data-appt-id="${apptId}" data-field-key="${fieldKey}">
        <label>
          <input type="checkbox" ${checked ? 'checked' : ''} onchange="updateApptCheckbox('${apptId}', '${fieldKey}', this.checked)">
          ${escapeHtml(label)}
        </label>
      </div>
    `;
  }
  
  // ============================================================
  // Edit Appointment Field
  // ============================================================
  
  window.editApptField = function(apptId, fieldKey, type, options) {
    const fieldEl = document.querySelector(`[data-appt-id="${apptId}"][data-field-key="${fieldKey}"] .appt-field-value`);
    if (!fieldEl) return;
    
    const currentValue = fieldEl.textContent.trim();
    let inputHtml = '';
    
    switch (type) {
      case 'select':
        inputHtml = `<select onchange="saveApptField('${apptId}', '${fieldKey}', this.value, '${type}')" onblur="cancelApptEdit('${apptId}', '${fieldKey}')">
          <option value="">-</option>
          ${options.map(o => `<option value="${escapeHtml(o)}" ${currentValue === o ? 'selected' : ''}>${escapeHtml(o)}</option>`).join('')}
        </select>`;
        break;
      case 'textarea':
        inputHtml = `<textarea onblur="saveApptField('${apptId}', '${fieldKey}', this.value, '${type}')">${currentValue === '-' ? '' : escapeHtml(currentValue)}</textarea>`;
        break;
      default:
        inputHtml = `<input type="text" value="${currentValue === '-' ? '' : escapeHtml(currentValue)}" onblur="saveApptField('${apptId}', '${fieldKey}', this.value, '${type}')">`;
    }
    
    fieldEl.innerHTML = inputHtml;
    const input = fieldEl.querySelector('input, select, textarea');
    if (input) input.focus();
  };
  
  window.saveApptField = function(apptId, fieldKey, value, type) {
    const updateData = {};
    updateData[fieldKey] = value || null;
    
    google.script.run.withSuccessHandler(function(result) {
      if (result.success && state.currentPanelData.record) {
        loadAppointmentsForOpportunity(state.currentPanelData.record.id);
      }
    }).updateAppointment(apptId, updateData);
  };
  
  window.updateApptCheckbox = function(apptId, fieldKey, checked) {
    const updateData = {};
    updateData[fieldKey] = checked;
    
    google.script.run.withSuccessHandler(function(result) {
      // Checkbox updates silently
    }).updateAppointment(apptId, updateData);
  };
  
  window.cancelApptEdit = function(apptId, fieldKey) {
    if (state.currentPanelData.record) {
      loadAppointmentsForOpportunity(state.currentPanelData.record.id);
    }
  };
  
  // ============================================================
  // Toggle Appointment Expand
  // ============================================================
  
  window.toggleAppointmentExpand = function(appointmentId) {
    const details = document.getElementById(`apptDetails_${appointmentId}`);
    const icon = document.querySelector(`[data-appt-id="${appointmentId}"] .appointment-expand-icon`);
    
    if (details) {
      const isHidden = details.style.display === 'none';
      details.style.display = isHidden ? 'block' : 'none';
      if (icon) icon.classList.toggle('expanded', isHidden);
    }
  };
  
  // ============================================================
  // Appointment Form
  // ============================================================
  
  window.openAppointmentForm = function(opportunityId, appointment = null) {
    const modal = document.getElementById('appointmentFormModal');
    const title = document.getElementById('appointmentFormTitle');
    
    document.getElementById('appointmentFormOppId').value = opportunityId;
    document.getElementById('appointmentFormId').value = appointment?.id || '';
    
    // Reset or populate fields
    document.getElementById('appointmentTime').value = appointment ? formatDatetimeForInput(appointment.appointmentTime) : '';
    document.getElementById('appointmentType').value = appointment?.type || 'Office';
    document.getElementById('appointmentHowBooked').value = appointment?.howBooked || '';
    document.getElementById('appointmentPhoneNumber').value = appointment?.phoneNumber || '';
    document.getElementById('appointmentVideoUrl').value = appointment?.videoMeetUrl || '';
    document.getElementById('appointmentStatus').value = appointment?.status || 'Scheduled';
    document.getElementById('appointmentNotes').value = appointment?.notes || '';
    
    title.textContent = appointment ? 'Edit Appointment' : 'Add Appointment';
    
    updateAppointmentFormVisibility();
    
    openModal('appointmentFormModal');
  };
  
  window.updateAppointmentFormVisibility = function() {
    const type = document.getElementById('appointmentType').value;
    document.getElementById('appointmentPhoneRow').style.display = type === 'Phone' ? 'flex' : 'none';
    document.getElementById('appointmentVideoRow').style.display = type === 'Video' ? 'flex' : 'none';
  };
  
  window.closeAppointmentForm = function() {
    closeModal('appointmentFormModal');
  };
  
  // ============================================================
  // Save Appointment
  // ============================================================
  
  window.saveAppointment = function() {
    const oppId = document.getElementById('appointmentFormOppId').value;
    const apptId = document.getElementById('appointmentFormId').value;
    
    const apptData = {
      appointmentTime: document.getElementById('appointmentTime').value || null,
      type: document.getElementById('appointmentType').value,
      howBooked: document.getElementById('appointmentHowBooked').value,
      phoneNumber: document.getElementById('appointmentPhoneNumber').value,
      videoMeetUrl: document.getElementById('appointmentVideoUrl').value,
      status: document.getElementById('appointmentStatus').value,
      notes: document.getElementById('appointmentNotes').value
    };
    
    closeAppointmentForm();
    
    if (apptId) {
      google.script.run.withSuccessHandler(function(result) {
        if (result.success) {
          loadAppointmentsForOpportunity(oppId);
        }
      }).updateAppointment(apptId, apptData);
    } else {
      google.script.run.withSuccessHandler(function(result) {
        if (result.success) {
          loadAppointmentsForOpportunity(oppId);
        }
      }).createAppointment(oppId, apptData);
    }
  };
  
  // ============================================================
  // Edit/Delete Appointment
  // ============================================================
  
  window.editAppointment = function(appointmentId, opportunityId) {
    google.script.run.withSuccessHandler(function(appointment) {
      if (appointment) {
        openAppointmentForm(opportunityId, appointment);
      }
    }).getAppointmentById(appointmentId);
  };
  
  window.deleteAppointment = function(appointmentId, opportunityId) {
    showConfirmModal('Are you sure you want to delete this appointment?', function() {
      google.script.run.withSuccessHandler(function(result) {
        if (result.success) {
          loadAppointmentsForOpportunity(opportunityId);
        }
      }).deleteAppointment(appointmentId);
    });
  };
  
  // ============================================================
  // Formatting Helpers
  // ============================================================
  
  window.formatAppointmentTime = function(dateStr) {
    if (!dateStr) return 'No time set';
    
    const d = new Date(dateStr);
    if (isNaN(d.getTime())) return dateStr;
    
    const days = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    
    const dayName = days[d.getDay()];
    const dayNum = d.getDate();
    const suffix = getOrdinalSuffix(dayNum);
    const month = months[d.getMonth()];
    
    let hours = d.getHours();
    const mins = String(d.getMinutes()).padStart(2, '0');
    const ampm = hours >= 12 ? 'pm' : 'am';
    hours = hours % 12 || 12;
    
    return `${dayName} ${dayNum}${suffix} ${month}, ${hours}:${mins}${ampm}`;
  };
  
  function getOrdinalSuffix(day) {
    if (day > 3 && day < 21) return 'th';
    switch (day % 10) {
      case 1: return 'st';
      case 2: return 'nd';
      case 3: return 'rd';
      default: return 'th';
    }
  }
  
  window.calculateDaysUntil = function(appointmentTimeStr) {
    if (!appointmentTimeStr) return '';
    
    const apptDate = new Date(appointmentTimeStr);
    if (isNaN(apptDate.getTime())) return '';
    
    const now = new Date();
    const diffMs = apptDate - now;
    const diffDays = Math.ceil(diffMs / (1000 * 60 * 60 * 24));
    
    if (diffDays < 0) return `${Math.abs(diffDays)} days ago`;
    if (diffDays === 0) return 'Today';
    if (diffDays === 1) return 'Tomorrow';
    return `in ${diffDays} days`;
  };
  
  function getAppointmentTypeIcon(type) {
    switch (type) {
      case 'Phone': return '&#128222;';
      case 'Video': return '&#128249;';
      case 'Office': return '&#127970;';
      default: return '&#128197;';
    }
  }
  
  window.togglePastApptFields = function() {
    const pastFieldsContainer = document.getElementById('pastApptFields');
    if (pastFieldsContainer) {
      pastFieldsContainer.classList.toggle('expanded');
    }
  };
  
})();
