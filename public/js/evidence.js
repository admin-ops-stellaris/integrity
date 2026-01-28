/**
 * evidence.js - Evidence & Data Collection Module for Integrity CRM
 * 
 * Handles evidence management for loan applications including:
 * - Evidence modal with category grouping and progress tracking
 * - Evidence items CRUD operations
 * - Client-facing view with copy to clipboard
 * - Evidence email generation (Initial/Subsequent/Appointment)
 * - Evidence templates management
 * 
 * Dependencies:
 * - window.IntegrityState (shared-state.js)
 * - window.showAlert, window.showConfirmModal (modal-utils.js)
 * - window.openModal, window.closeModal (modal-utils.js)
 * - Quill.js for rich text editing
 * - google.script.run API bridge
 * 
 * Load Order: After email.js, before app.js
 */

(function() {
  'use strict';

  // ==================== STATE ====================
  let currentEvidenceOpportunityId = null;
  let currentEvidenceOpportunityName = '';
  let currentEvidenceOpportunityType = '';
  let currentEvidenceLender = '';
  let currentEvidenceItems = [];
  
  // Quill editor instances
  let editEvidenceDescQuill = null;
  let newEvidenceDescQuill = null;
  let evidenceEmailQuill = null;
  let evTplDescQuill = null;
  
  // Email state
  let pendingEvidenceEmailItemIds = [];
  let currentEvidenceEmailType = null;
  
  // Delete confirmation state
  let pendingDeleteEvidenceItemId = null;
  
  // Templates state
  let allEvidenceTemplates = [];
  let editingEvidenceTemplateId = null;

  // ==================== EVIDENCE MODAL ====================
  
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
    
    let itemsToRender = [...currentEvidenceItems];
    if (outstandingFirst) {
      const getPriority = (status) => status === 'Received' ? 2 : status === 'N/A' ? 3 : 1;
      itemsToRender.sort((a, b) => getPriority(a.status) - getPriority(b.status));
    }
    
    const categoryOrder = ['Identification', 'Income', 'Assets', 'Liabilities', 'Refinance', 'Purchase & Property', 'Construction', 'Expenses', 'Other'];
    const grouped = {};
    
    itemsToRender.forEach(item => {
      const cat = item.category || 'Other';
      if (!grouped[cat]) grouped[cat] = [];
      grouped[cat].push(item);
    });
    
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
    
    let html = '';
    categoryOrder.forEach(cat => {
      if (!grouped[cat]) return;
      
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
    
    const hasRealContent = (html) => {
      if (!html) return false;
      const stripped = html.replace(/<[^>]*>/g, '').replace(/&nbsp;/g, ' ').replace(/\s+/g, ' ').trim();
      return stripped.length > 0;
    };
    
    let notesHtml = '';
    if (hasRealContent(item.notes)) {
      const notesMeta = item.modifiedByName && item.modifiedOn 
        ? `<span class="evidence-notes-meta">${item.modifiedByName} - ${formatPerthDate(item.modifiedOn)}</span>`
        : '';
      notesHtml = `<div class="evidence-item-notes"><strong>Internal Notes:</strong> ${item.notes}${notesMeta}</div>`;
    }
    
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

  // ==================== EDIT EVIDENCE ITEM ====================

  window.editEvidenceItem = function(itemId) {
    const item = currentEvidenceItems.find(i => i.id === itemId);
    if (!item) return;
    
    document.getElementById('editEvidenceItemId').value = itemId;
    document.getElementById('editEvidenceCategory').value = item.category || 'Other';
    document.getElementById('editEvidenceName').value = item.name || '';
    document.getElementById('editEvidenceNotes').value = item.notes || '';
    
    const modal = document.getElementById('editEvidenceItemModal');
    modal.classList.add('visible');
    setTimeout(() => modal.classList.add('showing'), 10);
    
    if (!editEvidenceDescQuill) {
      editEvidenceDescQuill = new Quill('#editEvidenceDescEditor', {
        theme: 'snow',
        modules: {
          toolbar: '#editEvidenceDescToolbar'
        },
        placeholder: 'Description...'
      });
      
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

  // ==================== DELETE EVIDENCE ITEM ====================
  
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

  // ==================== POPULATE FROM TEMPLATES ====================

  window.populateEvidenceFromTemplates = function() {
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

  // ==================== EMAIL MENU ====================

  window.toggleEvidenceEmailMenu = function() {
    const menu = document.getElementById('evidenceEmailMenu');
    menu.classList.toggle('show');
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

  // ==================== CLIENT VIEW ====================

  function buildClientEvidenceMarkup() {
    const outstanding = currentEvidenceItems.filter(i => i.status === 'Outstanding');
    const received = currentEvidenceItems.filter(i => i.status === 'Received');
    
    const total = outstanding.length + received.length;
    const pct = total > 0 ? Math.round((received.length / total) * 100) : 0;
    
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
    
    const leafIcon = `<img src="https://img1.wsimg.com/isteam/ip/2c5f94ee-4964-4e9b-9b9c-a55121f8611b/favicon/31eb51a1-8979-4194-bfa2-e4b30ee1178d/2437d5de-854d-40b2-86b2-fd879f3469f0.png" width="18" height="18" style="width:18px; height:18px; flex-shrink:0;">`;
    
    html += `<div style="margin-bottom:15px;">`;
    html += `<div style="display:flex; align-items:center; gap:10px; max-width:50%;">`;
    html += leafIcon;
    html += `<div style="flex:1; height:10px; background:#E0E0E0; border-radius:5px; overflow:hidden;">`;
    html += `<div style="width:${pct}%; height:100%; background:#7B8B64; border-radius:5px;"></div>`;
    html += `</div>`;
    html += `<span style="font-weight:bold; color:#2C2622; white-space:nowrap;">${pct}% (${received.length}/${total})</span>`;
    html += `</div></div>`;
    
    if (outstanding.length > 0) {
      html += `<div style="margin-bottom:15px;">`;
      html += `<h4 style="margin:0 0 8px 0; font-size:13px; font-weight:bold; color:#2C2622;">Outstanding</h4>`;
      html += `<ul style="margin:0; padding-left:0; list-style:none;">`;
      outstanding.forEach(item => { html += renderItem(item); });
      html += `</ul></div>`;
    }
    
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
    const html = buildClientEvidenceMarkup();
    const plainText = html.replace(/<[^>]*>/g, '').replace(/\s+/g, ' ').trim();
    
    const blob = new Blob([html], { type: 'text/html' });
    const clipboardItem = new ClipboardItem({
      'text/html': blob,
      'text/plain': new Blob([plainText], { type: 'text/plain' })
    });
    
    navigator.clipboard.write([clipboardItem]).then(() => {
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
    openEvidenceClientView();
  };

  // ==================== EVIDENCE EMAIL GENERATION ====================
  
  window.generateEvidenceEmail = function(type) {
    document.getElementById('evidenceEmailMenu').classList.remove('show');
    
    const outstanding = currentEvidenceItems.filter(i => i.status === 'Outstanding');
    if (outstanding.length === 0 && type !== 'appointment') {
      showAlert('info', 'No Items', 'No outstanding items to request!');
      return;
    }
    
    currentEvidenceEmailType = type;
    pendingEvidenceEmailItemIds = outstanding.map(i => i.id);
    
    const state = window.IntegrityState || {};
    const currentContactRecord = state.currentContactRecord;
    const contactName = currentContactRecord?.fields?.PreferredName || currentContactRecord?.fields?.FirstName || 'there';
    const contactEmail = currentContactRecord?.fields?.EmailAddress1 || '';
    
    const received = currentEvidenceItems.filter(i => i.status === 'Received');
    const total = outstanding.length + received.length;
    const pct = total > 0 ? Math.round((received.length / total) * 100) : 0;
    
    const cleanDesc = (desc) => {
      if (!desc) return '';
      return desc.replace(/<p>/gi, '').replace(/<\/p>/gi, ' ').replace(/<br\s*\/?>/gi, ' ').replace(/<div>/gi, '').replace(/<\/div>/gi, ' ').replace(/\s+/g, ' ').trim();
    };
    
    let evidenceListHtml = '';
    
    evidenceListHtml += `<p style="margin:8px 0 16px 0;"><strong style="color:#7B8B64;">Progress: ${pct}% (${received.length}/${total})</strong></p>`;
    
    if (outstanding.length > 0) {
      evidenceListHtml += `<p style="margin:16px 0 8px 0;"><strong style="color:#2C2622;">Outstanding</strong></p>`;
      outstanding.forEach(item => {
        const desc = cleanDesc(item.description);
        let line = `○ <strong>${item.name || 'Item'}</strong>`;
        if (desc) line += ` – ${desc}`;
        evidenceListHtml += `<p style="margin:6px 0 6px 12px; color:#2C2622;">${line}</p>`;
      });
      evidenceListHtml += `<p style="margin:0;"></p>`;
    }
    
    if (received.length > 0) {
      evidenceListHtml += `<p style="margin:16px 0 8px 0;"><strong style="color:#7B8B64;">Received</strong></p>`;
      received.forEach(item => {
        const desc = cleanDesc(item.description);
        let line = `✓ <strong>${item.name || 'Item'}</strong>`;
        if (desc) line += ` – ${desc}`;
        evidenceListHtml += `<p style="margin:6px 0 6px 12px; color:#7B8B64;">${line}</p>`;
      });
      evidenceListHtml += `<p style="margin:0;"></p>`;
    }
    
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
      
      let apptDetails = '';
      if (window.currentOpportunityAppointments && window.currentOpportunityAppointments.length > 0) {
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
    
    window.pendingEvidenceEmailBodyBase = body;
    
    document.getElementById('evidenceEmailModalTitle').textContent = title;
    document.getElementById('evidenceEmailTo').value = contactEmail;
    document.getElementById('evidenceEmailSubject').value = subject;
    document.getElementById('evidenceEmailItemCount').textContent = outstanding.length;
    
    const modal = document.getElementById('evidenceEmailModal');
    modal.classList.add('visible');
    setTimeout(() => modal.classList.add('showing'), 10);
    
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
    showCustomConfirm('Reset email to the original template? Your edits will be lost.', function() {
      initEmailEditorIframe(window.pendingEvidenceEmailBodyBase || '');
    });
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
        if (content && content.trim() && content.trim() !== '<br>') {
          return content;
        }
      }
    } catch (e) {
      console.warn('Could not read iframe content:', e);
    }
    
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
    
    google.script.run
      .withSuccessHandler(function(result) {
        if (result && result.success) {
          if (pendingEvidenceEmailItemIds.length > 0) {
            google.script.run
              .withSuccessHandler(function() {
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

  // ==================== ADD EVIDENCE ITEM ====================

  window.openAddEvidenceItemModal = function() {
    document.getElementById('newEvidenceName').value = '';
    document.getElementById('newEvidenceCategory').value = 'Other';
    
    const modal = document.getElementById('addEvidenceItemModal');
    modal.classList.add('visible');
    setTimeout(() => modal.classList.add('showing'), 10);
    
    if (!newEvidenceDescQuill) {
      newEvidenceDescQuill = new Quill('#newEvidenceDescEditor', {
        theme: 'snow',
        modules: {
          toolbar: '#newEvidenceDescToolbar'
        },
        placeholder: 'Describe what you need - supports bold, links, bullet points...'
      });
      
      const toolbar = newEvidenceDescQuill.getModule('toolbar');
      toolbar.addHandler('link', function(value) {
        if (value) {
          let href = prompt('Enter the link URL:');
          if (href) {
            if (!/^https?:\/\//i.test(href) && !/^mailto:/i.test(href)) {
              href = 'https://' + href;
            }
            const range = newEvidenceDescQuill.getSelection();
            if (range && range.length > 0) {
              newEvidenceDescQuill.format('link', href);
            } else {
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

  // ==================== EVIDENCE TEMPLATES MANAGEMENT ====================

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

  // ==================== MODULE INITIALIZATION ====================
  
  console.log('Evidence module loaded');
  
})();
