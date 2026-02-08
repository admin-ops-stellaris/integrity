(function() {
  'use strict';

  let allContacts = [];
  let cachedExportData = [];
  let campaigns = [];
  let parsedImportRows = [];
  let isLoading = false;
  let isImporting = false;

  window.openMarketingModal = function() {
    openModal('marketingModal');
    loadMarketingData();
    loadCampaigns();
  };

  window.closeMarketingModal = function() {
    closeModal('marketingModal');
    resetImportForm();
  };

  function loadMarketingData() {
    if (isLoading) return;
    isLoading = true;

    var statsEl = document.getElementById('marketingStats');
    var downloadBtn = document.getElementById('marketingDownloadBtn');
    statsEl.innerHTML = '<span style="color:#999;">Loading contacts...</span>';
    downloadBtn.disabled = true;

    google.script.run
      .withSuccessHandler(function(contacts) {
        allContacts = contacts || [];
        cachedExportData = calculateExportData(allContacts);
        isLoading = false;
        renderStats();
        downloadBtn.disabled = false;
      })
      .withFailureHandler(function(err) {
        console.error('Failed to load marketing data:', err);
        isLoading = false;
        statsEl.innerHTML = '<span style="color:#c00;">Failed to load contacts.</span>';
      })
      .getAllContactsForExport();
  }

  function loadCampaigns() {
    var selectEl = document.getElementById('campaignSelect');
    selectEl.disabled = true;
    selectEl.innerHTML = '<option value="">Loading campaigns...</option>';

    google.script.run
      .withSuccessHandler(function(data) {
        campaigns = data || [];
        selectEl.innerHTML = '<option value="">-- Select a campaign --</option>';
        campaigns.forEach(function(c) {
          var opt = document.createElement('option');
          opt.value = c.id;
          opt.textContent = c.name;
          selectEl.appendChild(opt);
        });
        selectEl.disabled = false;
        document.getElementById('importFileInput').disabled = false;
      })
      .withFailureHandler(function(err) {
        console.error('Failed to load campaigns:', err);
        selectEl.innerHTML = '<option value="">Failed to load campaigns</option>';
      })
      .getCampaigns();
  }

  function calculateExportData(contacts) {
    var marketable = contacts.filter(function(c) {
      return !c.unsubscribed && c.email && c.email.trim() !== '';
    });

    var grouped = {};
    marketable.forEach(function(contact) {
      var emailKey = contact.email.trim().toLowerCase();
      if (!grouped[emailKey]) {
        grouped[emailKey] = { names: [], ids: [], email: contact.email.trim() };
      }
      var fullName = [contact.firstName, contact.lastName].filter(Boolean).join(' ');
      if (fullName) grouped[emailKey].names.push(fullName);
      grouped[emailKey].ids.push(contact.id);
    });

    var rows = [];
    Object.keys(grouped).forEach(function(emailKey) {
      var group = grouped[emailKey];
      var combinedName;
      if (group.names.length === 0) {
        combinedName = '';
      } else if (group.names.length === 1) {
        combinedName = group.names[0];
      } else {
        var firstNames = group.names.map(function(n) { return n.split(' ')[0]; });
        var lastName = group.names[0].split(' ').slice(1).join(' ');
        combinedName = firstNames.join(' & ') + (lastName ? ' ' + lastName : '');
      }
      rows.push({
        name: combinedName,
        email: group.email,
        integrityId: group.ids.join(';')
      });
    });

    rows.sort(function(a, b) { return a.name.localeCompare(b.name); });
    return rows;
  }

  function formatNumber(n) {
    return n.toLocaleString();
  }

  function renderStats() {
    var total = allContacts.length;
    var unsubscribed = allContacts.filter(function(c) { return c.unsubscribed; }).length;
    var marketable = total - unsubscribed;
    var exportCount = cachedExportData.length;

    var statsEl = document.getElementById('marketingStats');
    statsEl.innerHTML =
      '<div class="marketing-stat-row">' +
        '<span class="marketing-stat-label">Total Contacts</span>' +
        '<span class="marketing-stat-value">' + formatNumber(total) + '</span>' +
      '</div>' +
      '<div class="marketing-stat-row">' +
        '<span class="marketing-stat-label">Unsubscribed</span>' +
        '<span class="marketing-stat-value marketing-stat-unsub">' + formatNumber(unsubscribed) + '</span>' +
      '</div>' +
      '<div class="marketing-stat-row marketing-stat-highlight">' +
        '<span class="marketing-stat-label">Marketable Contacts</span>' +
        '<span class="marketing-stat-value marketing-stat-green">' + formatNumber(marketable) + '</span>' +
      '</div>' +
      '<div class="marketing-stat-row marketing-stat-export">' +
        '<span class="marketing-stat-label">Ready to Send (Clean & Deduplicated)</span>' +
        '<span id="marketing-export-count" class="marketing-stat-value marketing-stat-export-value">' + formatNumber(exportCount) + '</span>' +
      '</div>';
  }

  window.downloadMarketingCsv = function() {
    var rows = cachedExportData;

    var csvLines = ['Name,Email,Integrity_ID'];
    rows.forEach(function(row) {
      csvLines.push(
        escapeCsvField(row.name) + ',' +
        escapeCsvField(row.email) + ',' +
        escapeCsvField(row.integrityId)
      );
    });

    var csvContent = csvLines.join('\n');
    var blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    var url = URL.createObjectURL(blob);

    var today = new Date();
    var dateStr = [
      today.getFullYear(),
      String(today.getMonth() + 1).padStart(2, '0'),
      String(today.getDate()).padStart(2, '0')
    ].join('');

    var link = document.createElement('a');
    link.href = url;
    link.download = 'integrity_export_' + dateStr + '.csv';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);

    var btn = document.getElementById('marketingDownloadBtn');
    var origText = btn.textContent;
    btn.textContent = 'Downloaded!';
    btn.disabled = true;
    setTimeout(function() {
      btn.textContent = origText;
      btn.disabled = false;
    }, 2000);
  };

  document.addEventListener('DOMContentLoaded', function() {
    var fileInput = document.getElementById('importFileInput');
    if (fileInput) {
      fileInput.addEventListener('change', handleFileSelect);
    }
  });

  function handleFileSelect(e) {
    var file = e.target.files[0];
    var infoEl = document.getElementById('importFileInfo');
    var importBtn = document.getElementById('marketingImportBtn');

    if (!file) {
      infoEl.style.display = 'none';
      parsedImportRows = [];
      importBtn.disabled = true;
      return;
    }

    var reader = new FileReader();
    reader.onload = function(evt) {
      var text = evt.target.result;
      parsedImportRows = parseCsv(text);

      if (parsedImportRows.length === 0) {
        infoEl.style.display = 'block';
        infoEl.innerHTML = '<span style="color:#c00;">No valid rows found. Check your CSV format.</span>';
        importBtn.disabled = true;
        return;
      }

      infoEl.style.display = 'block';
      infoEl.innerHTML = '<span style="color:var(--color-cedar);">' + formatNumber(parsedImportRows.length) + ' rows ready to import</span>';
      updateImportButton();
    };
    reader.readAsText(file);
  }

  function parseCsv(text) {
    var records = parseFullCsv(text);
    if (records.length < 2) return [];

    var headers = records[0].map(function(h) { return h.trim().toLowerCase(); });

    var emailIdx = headers.indexOf('email');
    var statusIdx = headers.indexOf('status');
    var idIdx = headers.indexOf('integrity_id');

    if (emailIdx === -1 || statusIdx === -1 || idIdx === -1) {
      console.error('CSV missing required columns. Found:', headers);
      return [];
    }

    var rows = [];
    for (var i = 1; i < records.length; i++) {
      var cols = records[i];
      if (cols.length <= Math.max(emailIdx, statusIdx, idIdx)) continue;
      var email = (cols[emailIdx] || '').trim();
      var status = (cols[statusIdx] || '').trim();
      var integrityId = (cols[idIdx] || '').trim();
      if (!integrityId) continue;
      rows.push({ email: email, status: status, integrityId: integrityId });
    }
    return rows;
  }

  function parseFullCsv(text) {
    var records = [];
    var current = '';
    var fields = [];
    var inQuotes = false;

    for (var i = 0; i < text.length; i++) {
      var ch = text[i];

      if (inQuotes) {
        if (ch === '"') {
          if (i + 1 < text.length && text[i + 1] === '"') {
            current += '"';
            i++;
          } else {
            inQuotes = false;
          }
        } else {
          current += ch;
        }
      } else {
        if (ch === '"') {
          inQuotes = true;
        } else if (ch === ',') {
          fields.push(current);
          current = '';
        } else if (ch === '\n' || ch === '\r') {
          if (ch === '\r' && i + 1 < text.length && text[i + 1] === '\n') {
            i++;
          }
          fields.push(current);
          current = '';
          if (fields.some(function(f) { return f.trim() !== ''; })) {
            records.push(fields);
          }
          fields = [];
        } else {
          current += ch;
        }
      }
    }

    fields.push(current);
    if (fields.some(function(f) { return f.trim() !== ''; })) {
      records.push(fields);
    }

    return records;
  }

  function updateImportButton() {
    var btn = document.getElementById('marketingImportBtn');
    var selectEl = document.getElementById('campaignSelect');
    btn.disabled = !(selectEl.value && parsedImportRows.length > 0);
  }

  document.addEventListener('DOMContentLoaded', function() {
    var selectEl = document.getElementById('campaignSelect');
    if (selectEl) {
      selectEl.addEventListener('change', updateImportButton);
    }
  });

  window.importCampaignResults = function() {
    if (isImporting) return;

    var selectEl = document.getElementById('campaignSelect');
    var campaignId = selectEl.value;
    if (!campaignId) {
      showImportStatus('Please select a campaign.', 'error');
      return;
    }
    if (parsedImportRows.length === 0) {
      showImportStatus('No rows to import. Please select a CSV file.', 'error');
      return;
    }

    var campaignName = selectEl.options[selectEl.selectedIndex].text;
    if (!confirm('Import ' + formatNumber(parsedImportRows.length) + ' rows into campaign "' + campaignName + '"?\n\nThis will update contact statuses for bounced/unsubscribed entries and create log records.')) {
      return;
    }

    isImporting = true;
    var btn = document.getElementById('marketingImportBtn');
    btn.disabled = true;
    btn.textContent = 'Importing...';
    showImportStatus('Processing ' + formatNumber(parsedImportRows.length) + ' rows... This may take a moment.', 'info');

    google.script.run
      .withSuccessHandler(function(response) {
        isImporting = false;
        btn.textContent = 'Import Results';

        if (response && response.success) {
          var r = response.results;
          var msg = 'Import complete: ' + formatNumber(r.processed) + ' records processed, ' +
            formatNumber(r.logged) + ' logs created';
          if (r.bounced > 0) msg += ', ' + r.bounced + ' bounced';
          if (r.unsubscribed > 0) msg += ', ' + r.unsubscribed + ' unsubscribed';
          if (r.skipped > 0) msg += ', ' + r.skipped + ' skipped (unknown status)';
          if (r.errors && r.errors.length > 0) {
            msg += '. ' + r.errors.length + ' error(s) occurred.';
            console.warn('Import errors:', r.errors);
          }
          showImportStatus(msg, 'success');
        } else {
          showImportStatus('Import failed: ' + (response.error || 'Unknown error'), 'error');
        }
        resetImportForm();
      })
      .withFailureHandler(function(err) {
        isImporting = false;
        btn.textContent = 'Import Results';
        btn.disabled = false;
        showImportStatus('Import failed: ' + (err.message || err), 'error');
      })
      .importCampaignResults({ campaignId: campaignId, rows: parsedImportRows });
  };

  function showImportStatus(message, type) {
    var el = document.getElementById('importResultsStatus');
    el.style.display = 'block';
    el.className = 'marketing-import-status marketing-import-status-' + type;
    el.textContent = message;
  }

  function resetImportForm() {
    parsedImportRows = [];
    var fileInput = document.getElementById('importFileInput');
    if (fileInput) fileInput.value = '';
    var infoEl = document.getElementById('importFileInfo');
    if (infoEl) infoEl.style.display = 'none';
    var btn = document.getElementById('marketingImportBtn');
    if (btn) {
      btn.disabled = true;
      btn.textContent = 'Import Results';
    }
  }

  function escapeCsvField(value) {
    if (!value) return '';
    if (value.indexOf(',') !== -1 || value.indexOf('"') !== -1 || value.indexOf('\n') !== -1) {
      return '"' + value.replace(/"/g, '""') + '"';
    }
    return value;
  }

})();
