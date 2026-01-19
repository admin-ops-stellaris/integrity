/**
 * Opportunities Module
 * Handles opportunity management: composer, list, panel, and deletion
 * 
 * Depends on: shared-state.js, shared-utils.js, modal-utils.js
 */
(function() {
  'use strict';
  
  const state = window.IntegrityState;

  // ============================================================
  // Quick Add Opportunity (Composer)
  // ============================================================

  window.quickAddOpportunity = function() {
    if (!state.currentContactRecord) { 
      alert('Please select a contact first.'); 
      return; 
    }
    openOppComposer();
  };

  function openOppComposer() {
    const f = state.currentContactRecord.fields;
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

  window.closeOppComposer = function() {
    document.getElementById('oppComposer').classList.remove('open');
    clearTacoImport();
  };

  // ============================================================
  // Taco Data Parsing
  // ============================================================

  window.parseTacoData = function() {
    const rawText = document.getElementById('tacoRawInput').value;
    if (!rawText.trim()) {
      document.getElementById('tacoPreview').style.display = 'none';
      state.parsedTacoFields = {};
      return;
    }
    
    google.script.run.withSuccessHandler(function(result) {
      state.parsedTacoFields = result.parsed || {};
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
  };

  function clearTacoImport() {
    document.getElementById('tacoRawInput').value = '';
    document.getElementById('tacoPreview').style.display = 'none';
    document.getElementById('tacoImportArea').style.display = 'block';
    state.parsedTacoFields = {};
  }
  window.clearTacoImport = clearTacoImport;

  window.updateComposerSpouseLabel = function() {
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
  };

  window.submitFromComposer = function() {
    const oppName = document.getElementById('composerOppName').value.trim();
    if (!oppName) { alert('Please enter an opportunity name.'); return; }
    
    const oppType = document.getElementById('composerOppType').value;
    const f = state.currentContactRecord.fields;
    const spouseId = (f['Spouse'] && f['Spouse'].length > 0) ? f['Spouse'][0] : null;
    const addSpouse = document.getElementById('composerAddSpouse')?.checked && spouseId;
    
    const tacoFieldsCopy = { ...state.parsedTacoFields };
    document.getElementById('oppComposer').classList.remove('open');
    
    google.script.run.withSuccessHandler(function(res) {
      clearTacoImport();
      if (res && res.id) {
        const finishUp = () => {
          google.script.run.withSuccessHandler(function(updatedContact) {
            if (updatedContact) {
              state.currentContactRecord = updatedContact;
              loadOpportunities(updatedContact.fields);
            }
            setTimeout(() => loadPanelRecord('Opportunities', res.id), 300);
          }).getContactById(state.currentContactRecord.id);
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
    }).createOpportunity(oppName, state.currentContactRecord.id, oppType, tacoFieldsCopy);
  };

  // ============================================================
  // Opportunity List
  // ============================================================

  window.loadOpportunities = function(f) {
    const oppList = document.getElementById('oppList');
    const loader = document.getElementById('oppLoading');
    document.getElementById('oppSortBtn').style.display = 'none';
    oppList.innerHTML = '';
    loader.style.display = 'block';
    
    let oppsToFetch = [];
    let roleMap = {};
    
    const addIds = (ids, roleName) => {
      if (!ids) return;
      (Array.isArray(ids) ? ids : [ids]).forEach(id => {
        oppsToFetch.push(id);
        roleMap[id] = roleName;
      });
    };
    
    addIds(f['Opportunities - Primary Applicant'], 'Primary Applicant');
    addIds(f['Opportunities - Applicant'], 'Applicant');
    addIds(f['Opportunities - Guarantor'], 'Guarantor');
    
    if (oppsToFetch.length === 0) {
      loader.style.display = 'none';
      oppList.innerHTML = '<li style="color:#CCC; font-size:12px; font-style:italic;">No opportunities linked.</li>';
      return;
    }
    
    google.script.run.withSuccessHandler(function(oppRecords) {
      loader.style.display = 'none';
      oppRecords.forEach(r => r._role = roleMap[r.id] || "Linked");
      if (oppRecords.length > 1) {
        document.getElementById('oppSortBtn').style.display = 'inline';
      }
      state.currentOppRecords = oppRecords;
      renderOppList();
    }).getLinkedOpportunities(oppsToFetch);
  };

  window.toggleOppSort = function() {
    if (state.currentOppSortDirection === 'asc') {
      state.currentOppSortDirection = 'desc';
    } else {
      state.currentOppSortDirection = 'asc';
    }
    renderOppList();
  };

  function renderOppList() {
    const oppList = document.getElementById('oppList');
    oppList.innerHTML = '';
    
    const sorted = [...state.currentOppRecords].sort((a, b) => {
      const nameA = (a.fields['Opportunity Name'] || "").toLowerCase();
      const nameB = (b.fields['Opportunity Name'] || "").toLowerCase();
      if (state.currentOppSortDirection === 'asc') return nameA.localeCompare(nameB);
      return nameB.localeCompare(nameA);
    });
    
    sorted.forEach(opp => {
      const fields = opp.fields;
      const name = fields['Opportunity Name'] || "Unnamed Opportunity";
      const role = opp._role;
      const status = fields['Status'] || '';
      const oppType = fields['Opportunity Type'] || '';
      const statusClass = status === 'Won' ? 'status-won' : status === 'Lost' ? 'status-lost' : '';
      
      const li = document.createElement('li');
      li.className = `opp-item ${statusClass}`;
      
      const statusBadge = status ? `<span class="opp-status-badge ${statusClass}">${status}</span>` : '';
      const typeLabel = oppType ? `<span class="opp-type">${oppType}</span>` : '';
      const roleLabel = role ? `<span class="opp-role-badge">${role}</span>` : '';
      
      li.innerHTML = `
        <span class="opp-title">${name}</span>
        <div class="opp-meta-row">${statusBadge}${typeLabel}${roleLabel}</div>
      `;
      li.onclick = function() {
        state.panelHistory = [];
        loadPanelRecord('Opportunities', opp.id);
      };
      oppList.appendChild(li);
    });
  }

  // ============================================================
  // Opportunity Panel
  // ============================================================

  window.loadPanelRecord = function(table, id) {
    const panel = document.getElementById('oppDetailPanel');
    const content = document.getElementById('panelContent');
    const titleEl = document.getElementById('panelTitle');
    
    panel.classList.add('open');
    content.innerHTML = `<div style="text-align:center; color:#999; margin-top:50px;">Loading...</div>`;
    
    google.script.run.withSuccessHandler(function(response) {
      if (!response || !response.data) {
        content.innerHTML = "Error loading.";
        return;
      }

      state.currentPanelData = {};
      response.data.forEach(item => {
        state.currentPanelData[item.key] = item.value;
      });

      state.panelHistory.push({ table: table, id: id, title: response.title });
      updateBackButton();
      titleEl.innerText = response.title;
      
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
            if (parts.length === 3) {
              inputVal = `${parts[2].length === 2 ? '20' + parts[2] : parts[2]}-${parts[1]}-${parts[0]}`;
              displayVal = `${parts[0]}/${parts[1]}/${parts[2].slice(-2)}`;
            } else if (rawVal.includes('-')) {
              const isoParts = rawVal.split('-');
              if (isoParts.length === 3) {
                inputVal = rawVal;
                displayVal = `${isoParts[2]}/${isoParts[1]}/${isoParts[0].slice(-2)}`;
              }
            } else {
              displayVal = rawVal;
              inputVal = rawVal;
            }
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
          const links = item.value;
          let linkHtml = '';
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
      
      if (table === 'Opportunities') {
        const dataMap = {};
        response.data.forEach(item => { dataMap[item.key] = item; });
        
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
        if (response.audit && (response.audit.Created || response.audit.Modified)) {
          let auditHtml = '<div class="panel-audit-section">';
          if (response.audit.Created) auditHtml += `<div>${response.audit.Created}</div>`;
          if (response.audit.Modified) auditHtml += `<div>${response.audit.Modified}</div>`;
          auditHtml += '</div>';
          html += auditHtml;
        }
      }
      
      if (table === 'Opportunities') {
        const dataMap = {};
        response.data.forEach(item => { dataMap[item.key] = item; });
        
        html += '<div class="panel-row panel-row-3">';
        ['Opportunity Name', 'Status', 'Opportunity Type'].forEach(key => {
          if (dataMap[key]) html += renderField(dataMap[key], table, id);
        });
        html += '</div>';
        
        const tacoFields = response.data.filter(item => item.tacoField);
        if (tacoFields.length > 0) {
          html += '<div class="taco-section-box">';
          html += `<div class="taco-section-header"><img src="https://taco.insightprocessing.com.au/static/images/taco.jpg" alt="Taco"><span>Fields from Taco Enquiry tab</span></div>`;
          html += '<div id="tacoFieldsContainer">';
          
          html += '<div class="taco-row">';
          if (dataMap['Taco: New or Existing Client']) html += renderField(dataMap['Taco: New or Existing Client'], table, id);
          if (dataMap['Taco: Lead Source']) html += renderField(dataMap['Taco: Lead Source'], table, id);
          html += '<div class="detail-group"></div>';
          html += '</div>';
          
          html += '<div class="taco-row">';
          if (dataMap['Taco: Last thing we did']) html += renderField(dataMap['Taco: Last thing we did'], table, id);
          if (dataMap['Taco: How can we help']) html += renderField(dataMap['Taco: How can we help'], table, id);
          if (dataMap['Taco: CM notes']) html += renderField(dataMap['Taco: CM notes'], table, id);
          html += '</div>';
          
          html += '<div class="taco-row">';
          if (dataMap['Taco: Broker']) html += renderField(dataMap['Taco: Broker'], table, id);
          if (dataMap['Taco: Broker Assistant']) html += renderField(dataMap['Taco: Broker Assistant'], table, id);
          if (dataMap['Taco: Client Manager']) html += renderField(dataMap['Taco: Client Manager'], table, id);
          html += '</div>';
          
          html += '<div class="taco-row">';
          if (dataMap['Taco: Converted to Appt']) html += renderField(dataMap['Taco: Converted to Appt'], table, id);
          html += '</div>';
          
          html += '</div></div>';
        }
        
        html += `<div class="appointments-section" style="margin-top:15px;">`;
        html += `<div id="appointmentsContainer" data-opportunity-id="${id}"><div style="color:#888; padding:10px;">Loading appointments...</div></div>`;
        html += `<div class="opp-action-buttons">`;
        html += `<button type="button" onclick="openAppointmentForm('${id}')">+ ADD APPOINTMENT</button>`;
        html += `</div></div>`;
        
        setTimeout(() => loadAppointmentsForOpportunity(id), 100);
        
        const applicantKeys = ['Primary Applicant', 'Applicants', 'Guarantors', 'Loan Applications'];
        html += '<div class="panel-row panel-row-4" style="margin-top:20px;">';
        applicantKeys.forEach(key => {
          if (dataMap[key]) html += renderField(dataMap[key], table, id);
        });
        html += '</div>';
        
        if (dataMap['Lead Source Major'] || dataMap['Lead Source Minor']) {
          html += '<div class="panel-row panel-row-2">';
          if (dataMap['Lead Source Major']) html += renderField(dataMap['Lead Source Major'], table, id);
          if (dataMap['Lead Source Minor']) html += renderField(dataMap['Lead Source Minor'], table, id);
          html += '</div>';
        }
        
        const usedKeys = new Set(['Opportunity Name', 'Status', 'Opportunity Type', 'Lead Source Major', 'Lead Source Minor', ...applicantKeys]);
        const remaining = response.data.filter(item => !item.tacoField && !usedKeys.has(item.key));
        if (remaining.length > 0) {
          html += '<div style="margin-top:15px; display:grid; grid-template-columns:repeat(3, 1fr); gap:12px 15px;">';
          remaining.forEach(item => { html += renderField(item, table, id); });
          html += '</div>';
        }
        
        const safeName = (response.title || '').replace(/'/g, "\\'");
        html += `<div style="margin-top:30px; padding-top:20px; border-top:1px solid #EEE;">`;
        html += `<button type="button" class="btn-delete btn-inline" onclick="confirmDeleteOpportunity('${id}', '${safeName}')">Delete Opportunity</button>`;
        html += `</div>`;
      } else {
        response.data.forEach(item => { html += renderField(item, table, id); });
      }
      
      content.innerHTML = html;
    }).getRecordDetail(table, id);
  };

  // ============================================================
  // Panel Navigation
  // ============================================================

  window.popHistory = function() {
    if (state.panelHistory.length <= 1) return;
    state.panelHistory.pop();
    const prev = state.panelHistory[state.panelHistory.length - 1];
    state.panelHistory.pop();
    loadPanelRecord(prev.table, prev.id);
  };

  function updateBackButton() {
    const btn = document.getElementById('panelBackBtn');
    if (state.panelHistory.length > 1) {
      btn.style.display = 'block';
    } else {
      btn.style.display = 'none';
    }
  }
  window.updateBackButton = updateBackButton;

  window.closeOppPanel = function() {
    document.getElementById('oppDetailPanel').classList.remove('open');
    state.panelHistory = [];
  };

  // ============================================================
  // Delete Opportunity
  // ============================================================

  window.confirmDeleteOpportunity = function(oppId, oppName) {
    state.currentOppToDelete = { id: oppId, name: oppName };
    document.getElementById('deleteOppConfirmMessage').innerText = `Are you sure you want to delete "${oppName}"? This action cannot be undone.`;
    openModal('deleteOppConfirmModal');
  };

  window.closeDeleteOppConfirmModal = function() {
    closeModal('deleteOppConfirmModal');
    state.currentOppToDelete = null;
  };

  window.executeDeleteOpportunity = function() {
    if (!state.currentOppToDelete) return;
    const { id, name } = state.currentOppToDelete;
    
    closeModal('deleteOppConfirmModal', function() {
      google.script.run
        .withSuccessHandler(function(result) {
          if (result.success) {
            showAlert('Success', `"${name}" has been deleted.`, 'success');
            closeOppPanel();
            if (state.currentContactRecord) {
              google.script.run.withSuccessHandler(function(updatedContact) {
                if (updatedContact) {
                  state.currentContactRecord = updatedContact;
                  loadOpportunities(updatedContact.fields);
                }
              }).getContactById(state.currentContactRecord.id);
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
    state.currentOppToDelete = null;
  };

  // ============================================================
  // Celebration (Won status)
  // ============================================================

  window.triggerWonCelebration = function() {
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
  };

})();
