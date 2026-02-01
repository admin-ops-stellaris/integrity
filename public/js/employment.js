/**
 * Employment Module
 * Employment history management with Primary conflict resolution, conditional visibility,
 * and JSON blob handling for Income and Address data
 */
(function() {
  'use strict';
  
  const state = window.IntegrityState;
  
  let editingEmploymentId = null;
  let conflictingPrimaryId = null;
  let currentIncomes = [];

  window.loadEmployment = function(contactId) {
    google.script.run.withSuccessHandler(function(employment) {
      state.currentEmployment = employment || [];
      renderEmploymentHistory();
    }).withFailureHandler(function(err) {
      console.error('Failed to load employment:', err);
      state.currentEmployment = [];
      renderEmploymentHistory();
    }).getEmploymentForContact(contactId);
  };

  window.renderEmploymentHistory = function() {
    const container = document.getElementById('employmentHistoryList');
    if (!container) return;
    
    const employment = state.currentEmployment || [];
    
    if (employment.length === 0) {
      container.innerHTML = '<div style="padding:10px; color:#888; font-size:12px; font-style:italic;">No employment recorded</div>';
      return;
    }
    
    container.innerHTML = employment.map(emp => {
      const statusClass = emp.status === 'Primary Employment' ? 'is-primary' : 
                         emp.status === 'Secondary Employment' ? 'is-secondary' : 'is-previous';
      const badgeClass = emp.status === 'Primary Employment' ? 'primary' : 
                        emp.status === 'Secondary Employment' ? 'secondary' : 'previous';
      const badgeText = emp.status === 'Primary Employment' ? 'Primary' : 
                       emp.status === 'Secondary Employment' ? 'Secondary' : 'Previous';
      
      const dateRange = formatEmploymentDateRange(emp.startDate, emp.endDate);
      const displayName = emp.employerName || emp.employmentType || 'Employment';
      
      return `
        <div class="employment-item ${statusClass}" onclick="editEmployment('${emp.id}')">
          <div class="employment-main">
            ${escapeHtml(displayName)}
            <span class="employment-status-badge ${badgeClass}">${badgeText}</span>
          </div>
          ${emp.jobTitle ? `<div class="employment-job-title">${escapeHtml(emp.jobTitle)}</div>` : ''}
          ${dateRange ? `<div class="employment-dates">${dateRange}</div>` : ''}
        </div>
      `;
    }).join('');
  };

  function formatEmploymentDateRange(startDate, endDate) {
    const startStr = startDate ? formatDateDisplay(startDate) : '';
    const endStr = endDate ? formatDateDisplay(endDate) : 'Present';
    if (startStr) {
      return `${startStr} - ${endStr}`;
    }
    return '';
  }

  window.openEmploymentModal = function() {
    const modal = document.getElementById('employmentModal');
    const title = document.getElementById('employmentFormTitle');
    const deleteBtn = document.getElementById('employmentDeleteBtn');
    
    document.getElementById('employmentFormId').value = '';
    
    ['employmentStatus', 'employmentType', 'employmentEmployerName', 'employmentAbn',
     'employmentJobTitle', 'employmentBasis', 'employmentPaygType', 'employmentOperatingStructure',
     'employmentStartDate', 'employmentEndDate', 'employmentContactTitle', 'employmentContactFirstName',
     'employmentContactSurname', 'employmentContactPhone', 'employmentContactEmail',
     'employmentAddrStreetNo', 'employmentAddrStreetName', 'employmentAddrStreetType',
     'employmentAddrCity', 'employmentAddrState', 'employmentAddrPostcode'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.value = '';
    });
    
    ['employmentOnProbation', 'employmentOnBenefits', 'employmentStudent'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.checked = false;
    });
    
    currentIncomes = [];
    renderIncomeRows();
    
    title.textContent = 'Add Employment';
    deleteBtn.style.display = 'none';
    editingEmploymentId = null;
    
    updateConditionalSections();
    
    modal.style.display = 'flex';
    setTimeout(() => modal.classList.add('showing'), 10);
  };

  window.editEmployment = function(employmentId) {
    const emp = (state.currentEmployment || []).find(e => e.id === employmentId);
    if (!emp) return;
    
    const modal = document.getElementById('employmentModal');
    const title = document.getElementById('employmentFormTitle');
    const deleteBtn = document.getElementById('employmentDeleteBtn');
    
    document.getElementById('employmentFormId').value = emp.id;
    document.getElementById('employmentStatus').value = emp.status || '';
    document.getElementById('employmentType').value = emp.employmentType || '';
    document.getElementById('employmentEmployerName').value = emp.employerName || '';
    document.getElementById('employmentAbn').value = emp.employerAbn || '';
    document.getElementById('employmentJobTitle').value = emp.jobTitle || '';
    document.getElementById('employmentBasis').value = emp.employmentBasis || '';
    document.getElementById('employmentPaygType').value = emp.paygType || '';
    document.getElementById('employmentOnProbation').checked = emp.onProbation || false;
    document.getElementById('employmentOperatingStructure').value = emp.operatingStructure || '';
    document.getElementById('employmentOnBenefits').checked = emp.onBenefits || false;
    document.getElementById('employmentStudent').checked = emp.student || false;
    
    if (emp.startDate) {
      const parsed = parseDateForEditor(emp.startDate);
      document.getElementById('employmentStartDate').value = parsed.dateDisplay || '';
    } else {
      document.getElementById('employmentStartDate').value = '';
    }
    
    if (emp.endDate) {
      const parsed = parseDateForEditor(emp.endDate);
      document.getElementById('employmentEndDate').value = parsed.dateDisplay || '';
    } else {
      document.getElementById('employmentEndDate').value = '';
    }
    
    document.getElementById('employmentContactTitle').value = emp.contactPersonTitle || '';
    document.getElementById('employmentContactFirstName').value = emp.contactPersonFirstName || '';
    document.getElementById('employmentContactSurname').value = emp.contactPersonSurname || '';
    document.getElementById('employmentContactPhone').value = emp.contactPersonPhone || '';
    document.getElementById('employmentContactEmail').value = emp.contactPersonEmail || '';
    
    try {
      const addrData = JSON.parse(emp.addressData || '{}');
      document.getElementById('employmentAddrStreetNo').value = addrData.streetNo || '';
      document.getElementById('employmentAddrStreetName').value = addrData.streetName || '';
      document.getElementById('employmentAddrStreetType').value = addrData.streetType || '';
      document.getElementById('employmentAddrCity').value = addrData.city || '';
      document.getElementById('employmentAddrState').value = addrData.state || '';
      document.getElementById('employmentAddrPostcode').value = addrData.postcode || '';
    } catch (e) {
      console.error('Error parsing address data:', e);
    }
    
    try {
      currentIncomes = JSON.parse(emp.incomes || '[]');
    } catch (e) {
      currentIncomes = [];
    }
    renderIncomeRows();
    
    title.textContent = 'Edit Employment';
    deleteBtn.style.display = 'block';
    editingEmploymentId = emp.id;
    
    updateConditionalSections();
    
    modal.style.display = 'flex';
    setTimeout(() => modal.classList.add('showing'), 10);
  };

  window.closeEmploymentModal = function() {
    const modal = document.getElementById('employmentModal');
    modal.classList.remove('showing');
    setTimeout(() => {
      modal.style.display = 'none';
      editingEmploymentId = null;
    }, 200);
  };

  window.handleEmploymentStatusChange = function() {
    const newStatus = document.getElementById('employmentStatus').value;
    
    if (newStatus === 'Primary Employment') {
      const existingPrimary = (state.currentEmployment || []).find(
        e => e.status === 'Primary Employment' && e.id !== editingEmploymentId
      );
      
      if (existingPrimary) {
        conflictingPrimaryId = existingPrimary.id;
        
        const message = document.getElementById('employmentConflictMessage');
        message.textContent = `"${existingPrimary.employerName || 'Current employment'}" is already set as Primary Employment.`;
        
        const conflictModal = document.getElementById('employmentConflictModal');
        conflictModal.style.display = 'flex';
        setTimeout(() => conflictModal.classList.add('showing'), 10);
      }
    }
    
    updateConditionalSections();
  };

  window.closeEmploymentConflictModal = function() {
    const modal = document.getElementById('employmentConflictModal');
    modal.classList.remove('showing');
    setTimeout(() => {
      modal.style.display = 'none';
      document.getElementById('employmentStatus').value = '';
      conflictingPrimaryId = null;
    }, 200);
  };

  window.resolveConflictAsSecondary = function() {
    if (!conflictingPrimaryId) return;
    
    google.script.run
      .withSuccessHandler(function() {
        const modal = document.getElementById('employmentConflictModal');
        modal.classList.remove('showing');
        setTimeout(() => modal.style.display = 'none', 200);
        
        const emp = (state.currentEmployment || []).find(e => e.id === conflictingPrimaryId);
        if (emp) emp.status = 'Secondary Employment';
        
        conflictingPrimaryId = null;
      })
      .withFailureHandler(function(err) {
        console.error('Failed to update employment status:', err);
        showAlert('Error', 'Failed to update existing employment: ' + err, 'error');
      })
      .updateEmployment(conflictingPrimaryId, { status: 'Secondary Employment' });
  };

  window.resolveConflictAsPrevious = function() {
    if (!conflictingPrimaryId) return;
    
    const endDateStr = prompt('Enter end date for the previous primary employment (DD/MM/YYYY):');
    if (endDateStr === null) return;
    
    const endDate = endDateStr ? parseSmartDate(endDateStr) : null;
    const updates = { status: 'Previous Employment' };
    if (endDate) {
      updates.endDate = endDate;
    }
    
    google.script.run
      .withSuccessHandler(function() {
        const modal = document.getElementById('employmentConflictModal');
        modal.classList.remove('showing');
        setTimeout(() => modal.style.display = 'none', 200);
        
        const emp = (state.currentEmployment || []).find(e => e.id === conflictingPrimaryId);
        if (emp) {
          emp.status = 'Previous Employment';
          if (endDate) emp.endDate = endDate;
        }
        
        conflictingPrimaryId = null;
      })
      .withFailureHandler(function(err) {
        console.error('Failed to update employment status:', err);
        showAlert('Error', 'Failed to update existing employment: ' + err, 'error');
      })
      .updateEmployment(conflictingPrimaryId, updates);
  };

  window.handleEmploymentTypeChange = function() {
    updateConditionalSections();
  };

  function updateConditionalSections() {
    const type = document.getElementById('employmentType').value;
    const status = document.getElementById('employmentStatus').value;
    
    const employerSection = document.getElementById('employmentEmployerSection');
    const paygSection = document.getElementById('employmentPaygSection');
    const selfEmployedSection = document.getElementById('employmentSelfEmployedSection');
    const benefitsSection = document.getElementById('employmentBenefitsSection');
    const contactSection = document.getElementById('employmentContactPersonSection');
    const addressSection = document.getElementById('employmentAddressSection');
    const incomeSection = document.getElementById('employmentIncomeSection');
    const addIncomeBtn = document.getElementById('employmentAddIncomeBtn');
    
    const showEmployer = status === 'Primary Employment' || status === 'Secondary Employment';
    employerSection.style.display = showEmployer ? 'block' : 'none';
    contactSection.style.display = showEmployer ? 'block' : 'none';
    addressSection.style.display = showEmployer ? 'block' : 'none';
    
    paygSection.style.display = type === 'PAYG' ? 'block' : 'none';
    selfEmployedSection.style.display = type === 'Self Employed' ? 'block' : 'none';
    benefitsSection.style.display = (type === 'Unemployed' || type === 'Retired') ? 'block' : 'none';
    
    const showIncome = type === 'PAYG' || type === 'Self Employed';
    incomeSection.style.display = showIncome ? 'block' : 'none';
    if (addIncomeBtn) addIncomeBtn.style.display = showIncome ? 'inline' : 'none';
  }

  window.addIncomeRow = function() {
    currentIncomes.push({ type: '', amount: '', frequency: '', comment: '' });
    renderIncomeRows();
  };

  function renderIncomeRows() {
    const container = document.getElementById('employmentIncomeList');
    if (!container) return;
    
    if (currentIncomes.length === 0) {
      container.innerHTML = '<div style="font-size:12px; color:#888; font-style:italic;">No income entries. Click "+ Add Income" to add.</div>';
      return;
    }
    
    container.innerHTML = currentIncomes.map((income, idx) => `
      <div class="income-row" data-idx="${idx}">
        <div class="field-group" style="flex:1;">
          <label>Type</label>
          <select onchange="updateIncomeField(${idx}, 'type', this.value)">
            <option value="" ${income.type === '' ? 'selected' : ''}>Select...</option>
            <option value="Base Salary" ${income.type === 'Base Salary' ? 'selected' : ''}>Base Salary</option>
            <option value="Overtime" ${income.type === 'Overtime' ? 'selected' : ''}>Overtime</option>
            <option value="Commission" ${income.type === 'Commission' ? 'selected' : ''}>Commission</option>
            <option value="Bonus" ${income.type === 'Bonus' ? 'selected' : ''}>Bonus</option>
            <option value="Allowance" ${income.type === 'Allowance' ? 'selected' : ''}>Allowance</option>
            <option value="Rental Income" ${income.type === 'Rental Income' ? 'selected' : ''}>Rental Income</option>
            <option value="Dividends" ${income.type === 'Dividends' ? 'selected' : ''}>Dividends</option>
            <option value="Centrelink" ${income.type === 'Centrelink' ? 'selected' : ''}>Centrelink</option>
            <option value="Child Support" ${income.type === 'Child Support' ? 'selected' : ''}>Child Support</option>
            <option value="Other" ${income.type === 'Other' ? 'selected' : ''}>Other</option>
          </select>
        </div>
        <div class="field-group" style="flex:0 0 100px;">
          <label>Amount ($)</label>
          <input type="text" value="${income.amount || ''}" onchange="updateIncomeField(${idx}, 'amount', this.value)">
        </div>
        <div class="field-group" style="flex:0 0 110px;">
          <label>Frequency</label>
          <select onchange="updateIncomeField(${idx}, 'frequency', this.value)">
            <option value="" ${income.frequency === '' ? 'selected' : ''}>Select...</option>
            <option value="Weekly" ${income.frequency === 'Weekly' ? 'selected' : ''}>Weekly</option>
            <option value="Fortnightly" ${income.frequency === 'Fortnightly' ? 'selected' : ''}>Fortnightly</option>
            <option value="Monthly" ${income.frequency === 'Monthly' ? 'selected' : ''}>Monthly</option>
            <option value="Yearly" ${income.frequency === 'Yearly' ? 'selected' : ''}>Yearly</option>
          </select>
        </div>
        <div class="field-group" style="flex:1;">
          <label>Comment</label>
          <input type="text" value="${escapeHtml(income.comment || '')}" onchange="updateIncomeField(${idx}, 'comment', this.value)">
        </div>
        <button type="button" class="income-delete-btn" onclick="deleteIncomeRow(${idx})">×</button>
      </div>
    `).join('');
  }

  window.updateIncomeField = function(idx, field, value) {
    if (currentIncomes[idx]) {
      currentIncomes[idx][field] = value;
    }
  };

  window.deleteIncomeRow = function(idx) {
    currentIncomes.splice(idx, 1);
    renderIncomeRows();
  };

  window.saveEmployment = function() {
    const contactId = state.currentContactId;
    if (!contactId) {
      showAlert('Error', 'No contact selected', 'error');
      return;
    }
    
    const status = document.getElementById('employmentStatus').value;
    const type = document.getElementById('employmentType').value;
    
    if (!status) {
      showAlert('Error', 'Please select a status', 'error');
      return;
    }
    
    const startDateRaw = document.getElementById('employmentStartDate').value;
    const endDateRaw = document.getElementById('employmentEndDate').value;
    
    const addressData = {
      streetNo: document.getElementById('employmentAddrStreetNo').value,
      streetName: document.getElementById('employmentAddrStreetName').value,
      streetType: document.getElementById('employmentAddrStreetType').value,
      city: document.getElementById('employmentAddrCity').value,
      state: document.getElementById('employmentAddrState').value,
      postcode: document.getElementById('employmentAddrPostcode').value
    };
    
    const fields = {
      status: status,
      employmentType: type,
      employerName: document.getElementById('employmentEmployerName').value,
      employerAbn: document.getElementById('employmentAbn').value,
      jobTitle: document.getElementById('employmentJobTitle').value,
      employmentBasis: document.getElementById('employmentBasis').value,
      paygType: document.getElementById('employmentPaygType').value,
      onProbation: document.getElementById('employmentOnProbation').checked,
      operatingStructure: document.getElementById('employmentOperatingStructure').value,
      onBenefits: document.getElementById('employmentOnBenefits').checked,
      student: document.getElementById('employmentStudent').checked,
      startDate: startDateRaw ? parseSmartDate(startDateRaw) : '',
      endDate: endDateRaw ? parseSmartDate(endDateRaw) : '',
      contactPersonTitle: document.getElementById('employmentContactTitle').value,
      contactPersonFirstName: document.getElementById('employmentContactFirstName').value,
      contactPersonSurname: document.getElementById('employmentContactSurname').value,
      contactPersonPhone: document.getElementById('employmentContactPhone').value,
      contactPersonEmail: document.getElementById('employmentContactEmail').value,
      addressData: JSON.stringify(addressData),
      incomes: JSON.stringify(currentIncomes)
    };
    
    const employmentId = document.getElementById('employmentFormId').value;
    
    if (employmentId) {
      google.script.run
        .withSuccessHandler(function(result) {
          if (result.success) {
            closeEmploymentModal();
            loadEmployment(contactId);
          } else {
            showAlert('Error', result.error || 'Failed to update employment', 'error');
          }
        })
        .withFailureHandler(function(err) {
          showAlert('Error', 'Failed to update employment: ' + err, 'error');
        })
        .updateEmployment(employmentId, fields);
    } else {
      google.script.run
        .withSuccessHandler(function(result) {
          if (result.success) {
            closeEmploymentModal();
            loadEmployment(contactId);
          } else {
            showAlert('Error', result.error || 'Failed to create employment', 'error');
          }
        })
        .withFailureHandler(function(err) {
          showAlert('Error', 'Failed to create employment: ' + err, 'error');
        })
        .createEmployment(contactId, fields);
    }
  };

  window.confirmDeleteEmployment = function() {
    const employmentId = document.getElementById('employmentFormId').value;
    if (!employmentId) return;
    
    showConfirm('Are you sure you want to delete this employment record?', function() {
      google.script.run
        .withSuccessHandler(function(result) {
          if (result.success) {
            closeEmploymentModal();
            loadEmployment(state.currentContactId);
          } else {
            showAlert('Error', result.error || 'Failed to delete employment', 'error');
          }
        })
        .withFailureHandler(function(err) {
          showAlert('Error', 'Failed to delete employment: ' + err, 'error');
        })
        .deleteEmployment(employmentId);
    });
  };

})();
