/**
 * Opportunities Module
 * Opportunity list, panel, and appointments
 */
(function() {
  'use strict';
  
  const state = window.IntegrityState;
  
  // ============================================================
  // Load Opportunities
  // ============================================================
  
  window.loadOpportunities = function(f) {
    const oppIds = [];
    if (f['Opportunities - Primary Applicant']) oppIds.push(...f['Opportunities - Primary Applicant']);
    if (f['Opportunities - Applicant']) oppIds.push(...f['Opportunities - Applicant']);
    
    const uniqueIds = [...new Set(oppIds)];
    
    if (uniqueIds.length === 0) {
      state.currentOppRecords = [];
      renderOppList();
      return;
    }
    
    google.script.run.withSuccessHandler(function(opps) {
      state.currentOppRecords = opps || [];
      renderOppList();
    }).getOpportunitiesByIds(uniqueIds);
  };
  
  // ============================================================
  // Toggle Sort Direction
  // ============================================================
  
  window.toggleOppSort = function() {
    state.currentOppSortDirection = state.currentOppSortDirection === 'desc' ? 'asc' : 'desc';
    renderOppList();
  };
  
  // ============================================================
  // Render Opportunity List
  // ============================================================
  
  window.renderOppList = function() {
    const list = document.getElementById('oppList');
    const sortIcon = document.getElementById('oppSortIcon');
    
    if (sortIcon) {
      sortIcon.innerHTML = state.currentOppSortDirection === 'desc' ? '&#9660;' : '&#9650;';
    }
    
    if (!state.currentOppRecords || state.currentOppRecords.length === 0) {
      list.innerHTML = '<li style="color:#CCC; font-size:12px; font-style:italic;">No opportunities linked.</li>';
      return;
    }
    
    const sorted = [...state.currentOppRecords].sort((a, b) => {
      const aDate = new Date(a.fields?.['Created On'] || 0);
      const bDate = new Date(b.fields?.['Created On'] || 0);
      return state.currentOppSortDirection === 'desc' ? bDate - aDate : aDate - bDate;
    });
    
    list.innerHTML = sorted.map(opp => {
      const f = opp.fields || {};
      const name = f['Opportunity Name'] || 'Unnamed';
      const status = f['Status'] || '';
      const statusClass = getOppStatusClass(status);
      
      return `
        <li class="opp-list-item" onclick="loadPanelRecord('Opportunities', '${opp.id}')">
          <span class="opp-name">${escapeHtml(name)}</span>
          <span class="opp-status ${statusClass}">${escapeHtml(status)}</span>
        </li>
      `;
    }).join('');
  };
  
  function getOppStatusClass(status) {
    const statusClasses = {
      'New': 'status-new',
      'In Progress': 'status-in-progress',
      'Submitted': 'status-submitted',
      'Conditionally Approved': 'status-conditionally-approved',
      'Approved': 'status-approved',
      'Settled': 'status-settled',
      'Won': 'status-won',
      'Lost': 'status-lost'
    };
    return statusClasses[status] || '';
  }
  
  // ============================================================
  // Quick Add Opportunity
  // ============================================================
  
  window.quickAddOpportunity = function() {
    openOppComposer();
  };
  
  window.openOppComposer = function() {
    const modal = document.getElementById('newOppModal');
    
    // Reset form
    document.getElementById('newOppType').value = '';
    document.getElementById('newOppAmount').value = '';
    document.getElementById('newOppLender').value = '';
    document.getElementById('tacoImport').value = '';
    document.getElementById('parsedTacoData').innerHTML = '';
    
    updateComposerSpouseLabel();
    
    openModal('newOppModal');
  };
  
  window.closeOppComposer = function() {
    closeModal('newOppModal');
  };
  
  window.closeNewOppModal = function() {
    closeOppComposer();
  };
  
  // ============================================================
  // Update Composer Spouse Label
  // ============================================================
  
  function updateComposerSpouseLabel() {
    const spouseCheckboxLabel = document.getElementById('addSpouseLabel');
    const spouseCheckbox = document.getElementById('addSpouseToOpp');
    
    if (state.currentContactRecord?.fields?.['Spouse Calculated Name']) {
      const spouseName = state.currentContactRecord.fields['Spouse Calculated Name'];
      spouseCheckboxLabel.textContent = `Add ${spouseName} as Applicant`;
      spouseCheckbox.parentElement.style.display = 'block';
    } else {
      spouseCheckbox.parentElement.style.display = 'none';
      spouseCheckbox.checked = false;
    }
  }
  
  // ============================================================
  // Taco Data Parser
  // ============================================================
  
  window.parseTacoData = function() {
    const tacoText = document.getElementById('tacoImport').value;
    const outputDiv = document.getElementById('parsedTacoData');
    
    if (!tacoText.trim()) {
      outputDiv.innerHTML = '';
      return;
    }
    
    const lines = tacoText.split('\n');
    const parsed = {};
    
    lines.forEach(line => {
      const colonIdx = line.indexOf(':');
      if (colonIdx > 0) {
        const key = line.substring(0, colonIdx).trim();
        const value = line.substring(colonIdx + 1).trim();
        if (key && value) {
          parsed[key] = value;
        }
      }
    });
    
    // Map to our fields
    if (parsed['Lender']) document.getElementById('newOppLender').value = parsed['Lender'];
    if (parsed['Loan Amount']) document.getElementById('newOppAmount').value = parsed['Loan Amount'].replace(/[,$]/g, '');
    
    outputDiv.innerHTML = `<div style="color:#7B8B64; font-size:11px;">Parsed ${Object.keys(parsed).length} fields</div>`;
  };
  
  window.clearTacoImport = function() {
    document.getElementById('tacoImport').value = '';
    document.getElementById('parsedTacoData').innerHTML = '';
  };
  
  // ============================================================
  // Submit from Composer
  // ============================================================
  
  window.submitFromComposer = function() {
    const contactId = state.currentContactRecord?.id;
    if (!contactId) {
      showAlert('Error', 'No contact selected', 'error');
      return;
    }
    
    const oppType = document.getElementById('newOppType').value;
    const amount = document.getElementById('newOppAmount').value;
    const lender = document.getElementById('newOppLender').value;
    const addSpouse = document.getElementById('addSpouseToOpp')?.checked || false;
    
    if (!oppType) {
      showAlert('Error', 'Please select an opportunity type', 'error');
      return;
    }
    
    const oppData = {
      'Opportunity Type': oppType,
      'Amount': amount ? parseFloat(amount.replace(/[,$]/g, '')) : null,
      'Lender': lender,
      'Primary Applicant': [contactId],
      'Status': 'New'
    };
    
    if (addSpouse && state.currentContactRecord?.fields?.Spouse?.[0]) {
      oppData['Applicant'] = [state.currentContactRecord.fields.Spouse[0]];
    }
    
    closeOppComposer();
    
    google.script.run.withSuccessHandler(function(result) {
      if (result && result.id) {
        // Reload opportunities
        loadOpportunities(state.currentContactRecord.fields);
        // Open the new opportunity
        loadPanelRecord('Opportunities', result.id);
      }
    }).createOpportunity(oppData);
  };
  
  // ============================================================
  // Load Panel Record (Opportunity Detail)
  // ============================================================
  
  window.loadPanelRecord = function(table, id) {
    const panel = document.getElementById('oppDetailPanel');
    panel.classList.add('open');
    
    state.panelHistory.push({ table, id });
    updateBackButton();
    
    if (table === 'Opportunities') {
      google.script.run.withSuccessHandler(function(record) {
        if (record && record.fields) {
          state.currentPanelData = { table, record };
          renderOpportunityPanel(record);
        }
      }).getOpportunityById(id);
    }
  };
  
  window.popHistory = function() {
    if (state.panelHistory.length <= 1) return;
    state.panelHistory.pop();
    const prev = state.panelHistory[state.panelHistory.length - 1];
    state.panelHistory.pop();
    loadPanelRecord(prev.table, prev.id);
  };
  
  function updateBackButton() {
    const btn = document.getElementById('panelBackBtn');
    if (btn) {
      btn.style.display = state.panelHistory.length > 1 ? 'block' : 'none';
    }
  }
  
  window.closeOppPanel = function() {
    document.getElementById('oppDetailPanel').classList.remove('open');
    state.panelHistory = [];
  };
  
  // ============================================================
  // Render Opportunity Panel
  // ============================================================
  
  function renderOpportunityPanel(record) {
    const f = record.fields;
    
    document.getElementById('panelTitle').textContent = f['Opportunity Name'] || 'Opportunity';
    
    const statusEl = document.getElementById('panelStatus');
    statusEl.textContent = f['Status'] || '';
    statusEl.className = 'panel-status ' + getOppStatusClass(f['Status'] || '');
    
    // Render fields
    const fieldsContainer = document.getElementById('panelFields');
    fieldsContainer.innerHTML = `
      ${renderPanelField(record.id, 'Opportunity Type', f['Opportunity Type'], 'select', getOppTypeOptions())}
      ${renderPanelField(record.id, 'Status', f['Status'], 'select', getStatusOptions())}
      ${renderPanelField(record.id, 'Amount', f['Amount'], 'currency')}
      ${renderPanelField(record.id, 'Lender', f['Lender'], 'text')}
      ${renderPanelField(record.id, 'Settlement Date', f['Settlement Date'], 'date')}
      ${renderPanelField(record.id, 'Notes', f['Notes'], 'textarea')}
    `;
    
    // Load appointments
    loadAppointmentsForOpportunity(record.id);
    
    // Render audit trail
    refreshPanelAudit('Opportunities', record.id);
    
    // Check for "Won" status
    if (f['Status'] === 'Won' || f['Status'] === 'Settled') {
      triggerWonCelebration();
    }
  }
  
  function renderPanelField(recordId, fieldKey, value, type, options = []) {
    const displayValue = formatFieldValue(value, type);
    const fieldId = `panelField_${fieldKey.replace(/\s+/g, '_')}`;
    
    return `
      <div class="panel-field" id="${fieldId}">
        <label class="panel-field-label">${escapeHtml(fieldKey)}</label>
        <div class="panel-field-value" onclick="toggleFieldEdit('${escapeHtmlForAttr(fieldKey)}')" data-record-id="${recordId}" data-field-key="${escapeHtmlForAttr(fieldKey)}" data-field-type="${type}" data-options="${escapeHtmlForAttr(JSON.stringify(options))}">
          ${displayValue || '<span class="empty-value">-</span>'}
        </div>
      </div>
    `;
  }
  
  function formatFieldValue(value, type) {
    if (value === null || value === undefined || value === '') return '';
    
    switch (type) {
      case 'currency':
        const num = parseFloat(value);
        return isNaN(num) ? value : '$' + num.toLocaleString('en-AU', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
      case 'date':
        return formatDateDisplay(value);
      default:
        return escapeHtml(String(value));
    }
  }
  
  function getOppTypeOptions() {
    return ['Home Loan', 'Investment Loan', 'Construction Loan', 'Personal Loan', 'Car Loan', 'Commercial Loan', 'Asset Finance', 'Other'];
  }
  
  function getStatusOptions() {
    return ['New', 'In Progress', 'Submitted', 'Conditionally Approved', 'Approved', 'Docs Out', 'Docs In', 'Settled', 'Won', 'Lost'];
  }
  
  // ============================================================
  // Toggle Field Edit
  // ============================================================
  
  window.toggleFieldEdit = function(fieldKey) {
    const fieldEl = document.querySelector(`[data-field-key="${fieldKey}"]`);
    if (!fieldEl) return;
    
    const recordId = fieldEl.dataset.recordId;
    const type = fieldEl.dataset.fieldType;
    const options = JSON.parse(fieldEl.dataset.options || '[]');
    const currentValue = fieldEl.textContent.trim();
    
    // Replace with input
    let inputHtml = '';
    
    switch (type) {
      case 'select':
        inputHtml = `<select onchange="saveFieldEdit('Opportunities', '${recordId}', '${escapeHtmlForAttr(fieldKey)}')" onblur="cancelFieldEdit('${escapeHtmlForAttr(fieldKey)}')">
          <option value="">-</option>
          ${options.map(o => `<option value="${escapeHtml(o)}" ${currentValue === o ? 'selected' : ''}>${escapeHtml(o)}</option>`).join('')}
        </select>`;
        break;
      case 'textarea':
        inputHtml = `<textarea onblur="saveFieldEdit('Opportunities', '${recordId}', '${escapeHtmlForAttr(fieldKey)}')">${escapeHtml(currentValue === '-' ? '' : currentValue)}</textarea>`;
        break;
      case 'date':
        inputHtml = `<input type="text" placeholder="DD/MM/YYYY" value="${currentValue === '-' ? '' : currentValue}" onblur="saveDateField('Opportunities', '${recordId}', '${escapeHtmlForAttr(fieldKey)}')">`;
        break;
      case 'currency':
        const numValue = currentValue.replace(/[$,]/g, '');
        inputHtml = `<input type="text" value="${numValue === '-' ? '' : numValue}" onblur="saveFieldEdit('Opportunities', '${recordId}', '${escapeHtmlForAttr(fieldKey)}')">`;
        break;
      default:
        inputHtml = `<input type="text" value="${currentValue === '-' ? '' : currentValue}" onblur="saveFieldEdit('Opportunities', '${recordId}', '${escapeHtmlForAttr(fieldKey)}')">`;
    }
    
    fieldEl.innerHTML = inputHtml;
    const input = fieldEl.querySelector('input, select, textarea');
    if (input) {
      input.focus();
      if (input.tagName === 'SELECT') {
        // Keep select open
      }
    }
  };
  
  window.cancelFieldEdit = function(fieldKey) {
    // Re-render panel
    if (state.currentPanelData.record) {
      renderOpportunityPanel(state.currentPanelData.record);
    }
  };
  
  window.saveFieldEdit = function(table, id, fieldKey) {
    const fieldEl = document.querySelector(`[data-field-key="${fieldKey}"]`);
    const input = fieldEl?.querySelector('input, select, textarea');
    
    if (!input) return;
    
    let value = input.value;
    const type = fieldEl.dataset.fieldType;
    
    // Convert currency back to number
    if (type === 'currency' && value) {
      value = parseFloat(value.replace(/[,$]/g, ''));
    }
    
    const updateData = {};
    updateData[fieldKey] = value || null;
    
    google.script.run.withSuccessHandler(function(result) {
      if (result && result.id) {
        state.currentPanelData.record = result;
        renderOpportunityPanel(result);
        
        // Reload opportunity list
        if (state.currentContactRecord) {
          loadOpportunities(state.currentContactRecord.fields);
        }
      }
    }).updateOpportunity(id, updateData);
  };
  
  window.saveDateField = function(table, id, fieldKey) {
    const fieldEl = document.querySelector(`[data-field-key="${fieldKey}"]`);
    const input = fieldEl?.querySelector('input');
    
    if (!input) return;
    
    const value = parseDateInput(input.value);
    
    const updateData = {};
    updateData[fieldKey] = value;
    
    google.script.run.withSuccessHandler(function(result) {
      if (result && result.id) {
        state.currentPanelData.record = result;
        renderOpportunityPanel(result);
      }
    }).updateOpportunity(id, updateData);
  };
  
  // ============================================================
  // Panel Audit Trail
  // ============================================================
  
  window.refreshPanelAudit = function(table, id) {
    const auditEl = document.getElementById('panelAudit');
    if (!auditEl) return;
    
    const record = state.currentPanelData.record;
    if (!record) return;
    
    const f = record.fields;
    const createdOn = f['Created On'];
    const createdBy = f['Created By'];
    const modifiedOn = f['Modified On'];
    const modifiedBy = f['Modified By'];
    
    let html = '';
    if (createdOn) {
      html += `<div class="audit-item">Created: ${formatAuditDate(createdOn)}${createdBy ? ' by ' + createdBy : ''}</div>`;
    }
    if (modifiedOn) {
      html += `<div class="audit-item">Modified: ${formatAuditDate(modifiedOn)}${modifiedBy ? ' by ' + modifiedBy : ''}</div>`;
    }
    
    auditEl.innerHTML = html;
  };
  
  // ============================================================
  // Delete Opportunity
  // ============================================================
  
  window.confirmDeleteOpportunity = function(oppId, oppName) {
    state.deletingOppId = oppId;
    document.getElementById('deleteOppName').textContent = oppName || 'this opportunity';
    openModal('deleteOppConfirmModal');
  };
  
  window.closeDeleteOppConfirmModal = function() {
    closeModal('deleteOppConfirmModal');
    state.deletingOppId = null;
  };
  
  window.executeDeleteOpportunity = function() {
    const oppId = state.deletingOppId;
    if (!oppId) return;
    
    closeDeleteOppConfirmModal();
    closeOppPanel();
    
    google.script.run.withSuccessHandler(function(result) {
      if (result.success) {
        if (state.currentContactRecord) {
          loadOpportunities(state.currentContactRecord.fields);
        }
      } else {
        showAlert('Error', result.error || 'Failed to delete opportunity', 'error');
      }
    }).deleteOpportunity(oppId);
  };
  
})();
