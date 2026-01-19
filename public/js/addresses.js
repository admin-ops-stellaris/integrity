/**
 * Addresses Module
 * Address history management, modals, and CRUD operations
 */
(function() {
  'use strict';
  
  const state = window.IntegrityState;
  
  // ============================================================
  // Load Address History
  // ============================================================
  
  window.loadAddressHistory = function(contactId) {
    google.script.run.withSuccessHandler(function(addresses) {
      state.currentContactAddresses = addresses || [];
      renderAddressHistory();
    }).getAddressesForContact(contactId);
  };
  
  // ============================================================
  // Render Address History
  // ============================================================
  
  window.renderAddressHistory = function() {
    const container = document.getElementById('addressHistoryList');
    const postalDisplay = document.getElementById('postalAddressDisplay');
    
    if (!container) return;
    
    // Separate residential and postal addresses
    const residential = state.currentContactAddresses.filter(a => !a.isPostal);
    const postal = state.currentContactAddresses.find(a => a.isPostal);
    
    // Render postal address section
    if (postalDisplay) {
      if (postal) {
        postalDisplay.innerHTML = `
          <div class="postal-address-display" onclick="editAddress('${postal.id}')">
            ${escapeHtml(postal.calculatedName) || 'No address'}
          </div>
        `;
      } else {
        postalDisplay.innerHTML = '';
      }
    }
    
    // Update postal button text: +Add when none exists, +Update when one does
    const postalBtn = document.getElementById('postalAddressAddBtn');
    if (postalBtn) {
      postalBtn.textContent = postal ? '+ Update' : '+ Add';
    }
    
    // Render residential addresses
    if (residential.length === 0) {
      container.innerHTML = '<div style="padding:10px; color:#888; font-size:12px; font-style:italic;">No addresses recorded</div>';
      return;
    }
    
    // Sort: current first, then by To date descending
    residential.sort((a, b) => {
      const aIsCurrent = !a.to || a.status === 'Current';
      const bIsCurrent = !b.to || b.status === 'Current';
      if (aIsCurrent && !bIsCurrent) return -1;
      if (!aIsCurrent && bIsCurrent) return 1;
      const aTo = a.to ? new Date(a.to) : new Date();
      const bTo = b.to ? new Date(b.to) : new Date();
      return bTo - aTo;
    });
    
    container.innerHTML = residential.map(addr => {
      const isCurrent = !addr.to || addr.status === 'Current';
      const dateRange = formatAddressDateRange(addr.from, addr.to);
      
      return `
        <div class="address-item ${isCurrent ? 'is-current' : ''}" onclick="editAddress('${addr.id}')">
          <div class="address-main">${escapeHtml(addr.calculatedName) || 'No address'}</div>
          ${dateRange ? `<div class="address-dates">${dateRange}</div>` : ''}
          ${addr.status ? `<div class="address-status">${escapeHtml(addr.status)}</div>` : ''}
        </div>
      `;
    }).join('');
  };
  
  // ============================================================
  // Format Helpers
  // ============================================================
  
  function formatAddressDateRange(from, to) {
    const fromStr = from ? formatDateDisplay(from) : '';
    const toStr = to ? formatDateDisplay(to) : 'Present';
    if (fromStr) {
      return `${fromStr} - ${toStr}`;
    }
    return '';
  }
  
  // ============================================================
  // Open Address Modal
  // ============================================================
  
  window.openAddressModal = function(isPostal) {
    const modal = document.getElementById('addressFormModal');
    const title = document.getElementById('addressFormTitle');
    const deleteBtn = document.getElementById('addressDeleteBtn');
    const statusRow = document.getElementById('addressStatusRow');
    
    // Reset form
    document.getElementById('addressFormId').value = '';
    document.getElementById('addressFormIsPostal').value = isPostal ? 'true' : 'false';
    document.querySelectorAll('input[name="addressFormat"]').forEach(r => r.checked = r.value === 'Standard');
    
    // Clear all fields
    ['addressFloor', 'addressBuilding', 'addressUnit', 'addressStreetNo', 'addressStreetName', 
     'addressStreetType', 'addressCity', 'addressState', 'addressPostcode', 'addressCountry',
     'addressLabel', 'addressStatus', 'addressFrom', 'addressTo', 'addressNonStandardLine',
     'addressPOBoxNo'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.value = '';
    });
    
    title.textContent = isPostal ? 'Add Postal Address' : 'Add Address';
    deleteBtn.style.display = 'none';
    
    // Hide status/from/to for postal
    statusRow.style.display = isPostal ? 'none' : 'flex';
    
    state.editingAddressId = null;
    updateAddressFormatFields();
    
    modal.style.display = 'flex';
    setTimeout(() => modal.classList.add('showing'), 10);
  };
  
  // ============================================================
  // Postal Address Modal
  // ============================================================
  
  window.openPostalAddressModal = function() {
    const existingPostal = state.currentContactAddresses.find(a => a.isPostal);
    const hasResidential = state.currentContactAddresses.filter(a => !a.isPostal).length > 0;
    
    if (!existingPostal && hasResidential) {
      // Show copy choice modal
      const copyModal = document.getElementById('postalCopyModal');
      const residentialList = document.getElementById('postalCopyResidentialList');
      
      if (residentialList) {
        const residential = state.currentContactAddresses.filter(a => !a.isPostal);
        residentialList.innerHTML = residential.map(addr => `
          <div class="postal-copy-item" onclick="copyAddressAsPostal('${addr.id}')">
            ${escapeHtml(addr.calculatedName) || 'No address'}
          </div>
        `).join('');
      }
      
      copyModal.style.display = 'flex';
      setTimeout(() => copyModal.classList.add('showing'), 10);
    } else {
      // Open blank postal form
      openAddressModal(true);
    }
  };
  
  window.closePostalCopyModal = function() {
    const modal = document.getElementById('postalCopyModal');
    modal.classList.remove('showing');
    setTimeout(() => modal.style.display = 'none', 250);
  };
  
  window.openPostalAddressNew = function() {
    closePostalCopyModal();
    openAddressModal(true);
  };
  
  window.copyAddressAsPostal = function(addressId) {
    closePostalCopyModal();
    
    const sourceAddr = state.currentContactAddresses.find(a => a.id === addressId);
    if (!sourceAddr) {
      openAddressModal(true);
      return;
    }
    
    // Open modal and populate with source data
    openAddressModal(true);
    
    // Populate fields from source
    document.querySelectorAll('input[name="addressFormat"]').forEach(r => {
      r.checked = r.value === (sourceAddr.format || 'Standard');
    });
    
    document.getElementById('addressFloor').value = sourceAddr.floor || '';
    document.getElementById('addressBuilding').value = sourceAddr.building || '';
    document.getElementById('addressUnit').value = sourceAddr.unit || '';
    document.getElementById('addressStreetNo').value = sourceAddr.streetNo || '';
    document.getElementById('addressStreetName').value = sourceAddr.streetName || '';
    document.getElementById('addressStreetType').value = sourceAddr.streetType || '';
    document.getElementById('addressCity').value = sourceAddr.city || '';
    document.getElementById('addressState').value = sourceAddr.state || '';
    document.getElementById('addressPostcode').value = sourceAddr.postcode || '';
    document.getElementById('addressCountry').value = sourceAddr.country || 'Australia';
    document.getElementById('addressLabel').value = sourceAddr.label || '';
    
    if (sourceAddr.format === 'Non-Standard') {
      document.getElementById('addressNonStandardLine').value = sourceAddr.nonStandardLine || '';
    }
    if (sourceAddr.format === 'PO Box') {
      document.getElementById('addressPOBoxNo').value = sourceAddr.poBoxNo || '';
    }
    
    updateAddressFormatFields();
  };
  
  // ============================================================
  // Edit Address
  // ============================================================
  
  window.editAddress = function(addressId) {
    const addr = state.currentContactAddresses.find(a => a.id === addressId);
    if (!addr) return;
    
    const modal = document.getElementById('addressFormModal');
    const title = document.getElementById('addressFormTitle');
    const deleteBtn = document.getElementById('addressDeleteBtn');
    const statusRow = document.getElementById('addressStatusRow');
    
    state.editingAddressId = addressId;
    
    document.getElementById('addressFormId').value = addressId;
    document.getElementById('addressFormIsPostal').value = addr.isPostal ? 'true' : 'false';
    
    document.querySelectorAll('input[name="addressFormat"]').forEach(r => {
      r.checked = r.value === (addr.format || 'Standard');
    });
    
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
    
    if (addr.format === 'Non-Standard') {
      document.getElementById('addressNonStandardLine').value = addr.nonStandardLine || '';
    }
    if (addr.format === 'PO Box') {
      document.getElementById('addressPOBoxNo').value = addr.poBoxNo || '';
    }
    
    title.textContent = addr.isPostal ? 'Edit Postal Address' : 'Edit Address';
    deleteBtn.style.display = 'inline-block';
    
    // Hide status/from/to for postal
    statusRow.style.display = addr.isPostal ? 'none' : 'flex';
    
    updateAddressFormatFields();
    
    modal.style.display = 'flex';
    setTimeout(() => modal.classList.add('showing'), 10);
  };
  
  // ============================================================
  // Close Address Form
  // ============================================================
  
  window.closeAddressForm = function() {
    const modal = document.getElementById('addressFormModal');
    modal.classList.remove('showing');
    setTimeout(() => modal.style.display = 'none', 250);
    state.editingAddressId = null;
  };
  
  // ============================================================
  // Update Format Fields Visibility
  // ============================================================
  
  window.updateAddressFormatFields = function() {
    const format = document.querySelector('input[name="addressFormat"]:checked')?.value || 'Standard';
    
    document.getElementById('addressNonStandardFields').style.display = format === 'Non-Standard' ? 'block' : 'none';
    document.getElementById('addressPOBoxFields').style.display = format === 'PO Box' ? 'block' : 'none';
    document.getElementById('addressStreetFields').style.display = format === 'PO Box' ? 'none' : 'block';
  };
  
  // ============================================================
  // Save Address
  // ============================================================
  
  window.saveAddress = function() {
    const recordId = state.currentContactRecord?.id;
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
      status: isPostal ? null : (document.getElementById('addressStatus').value || null),
      from: isPostal ? null : parseDateInput(document.getElementById('addressFrom').value),
      to: isPostal ? null : parseDateInput(document.getElementById('addressTo').value),
      isPostal: isPostal
    };
    
    // For postal addresses, delete existing postal before creating new (replace behavior)
    const existingPostal = state.currentContactAddresses.find(a => a.isPostal);
    if (isPostal && !state.editingAddressId && existingPostal) {
      google.script.run
        .withSuccessHandler(function(deleteResult) {
          if (deleteResult.success) {
            createNewAddress(recordId, fields);
          } else {
            alert('Error replacing postal address: ' + (deleteResult.error || 'Unknown error'));
          }
        })
        .withFailureHandler(function(err) {
          console.error('Error deleting old postal:', err);
          alert('Error replacing postal address: ' + (err.message || err));
        })
        .deleteAddress(existingPostal.id);
      return;
    }
    
    if (state.editingAddressId) {
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
        .updateAddress(state.editingAddressId, fields);
    } else {
      createNewAddress(recordId, fields);
    }
  };
  
  function createNewAddress(recordId, fields) {
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
  
  // ============================================================
  // Delete Address
  // ============================================================
  
  window.deleteAddress = function() {
    if (!state.editingAddressId) {
      console.error('deleteAddress: No editingAddressId');
      return;
    }
    
    const recordId = state.currentContactRecord?.id;
    const addressId = state.editingAddressId;
    const isPostal = document.getElementById('addressFormIsPostal').value === 'true';
    const confirmMessage = isPostal
      ? 'Are you sure you want to remove the postal address?'
      : 'Are you sure you want to delete this address?';
    
    showConfirmModal(confirmMessage, function() {
      console.log('Deleting address:', addressId);
      google.script.run
        .withSuccessHandler(function(result) {
          console.log('Delete result:', result);
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
  
  // ============================================================
  // Toggle Address Expand
  // ============================================================
  
  window.toggleAddressExpand = function(event) {
    const item = event.currentTarget.closest('.address-item');
    if (item) {
      item.classList.toggle('expanded');
    }
  };
  
})();
