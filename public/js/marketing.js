(function() {
  'use strict';

  var allContacts = [];
  var cachedExportData = [];
  var campaigns = [];
  var parsedImportRows = [];
  var isLoading = false;
  var isImporting = false;

  var campaignStatsCache = null;
  var currentDetailLogs = [];
  var currentDetailFilter = 'all';
  var currentDetailCampaignName = '';
  var currentSortKey = 'dateSent';
  var currentSortDir = 'desc';

  window.openMarketingModal = function() {
    openModal('marketingModal');
    switchMarketingTab('campaigns');
  };

  window.closeMarketingModal = function() {
    closeModal('marketingModal');
    resetImportForm();
    showCampaignDashboard();
  };

  window.switchMarketingTab = function(tabName) {
    var tabs = document.querySelectorAll('.marketing-tab');
    var panels = document.querySelectorAll('.marketing-tab-panel');
    tabs.forEach(function(t) { t.classList.toggle('active', t.dataset.tab === tabName); });
    panels.forEach(function(p) { p.classList.remove('active'); });

    if (tabName === 'campaigns') {
      document.getElementById('marketingTabCampaigns').classList.add('active');
      loadCampaignStats();
    } else if (tabName === 'datatools') {
      document.getElementById('marketingTabDatatools').classList.add('active');
      loadMarketingData();
      loadCampaigns();
    }
  };

  function showCreateError(btn, message) {
    var container = btn.parentElement;
    var errorEl = document.getElementById('camp-create-error');
    if (!errorEl) {
      errorEl = document.createElement('div');
      errorEl.id = 'camp-create-error';
      errorEl.style.color = '#dc3545';
      errorEl.style.fontSize = '12px';
      errorEl.style.marginTop = '4px';
      container.appendChild(errorEl);
    }
    errorEl.textContent = message;
    setTimeout(function() { errorEl.textContent = ''; }, 4000);
  }

  function checkDuplicateAndCreate(name, subject, btn, nameInput, subjectInput, existingCampaigns) {
    var duplicate = (existingCampaigns || []).find(function(c) {
      return c && c.name && c.name.toLowerCase() === name.toLowerCase();
    });
    if (duplicate) {
      btn.disabled = false;
      btn.textContent = '+ Create';
      showCreateError(btn, 'A campaign with this name already exists.');
      nameInput.focus();
      nameInput.select();
      return;
    }

    google.script.run
      .withSuccessHandler(function(campaign) {
        nameInput.value = '';
        subjectInput.value = '';
        btn.disabled = false;
        btn.textContent = '+ Create';
        campaignStatsCache = null;
        campaigns = [];
        loadCampaignStats();
      })
      .withFailureHandler(function(err) {
        btn.disabled = false;
        btn.textContent = '+ Create';
        showCreateError(btn, 'Error: ' + (err.message || err));
      })
      .createCampaign(name, subject);
  }

  window.createNewCampaign = function() {
    var nameInput = document.getElementById('newCampaignNameInput');
    var subjectInput = document.getElementById('newCampaignSubjectInput');
    var name = (nameInput.value || '').trim();
    var subject = (subjectInput.value || '').trim();
    if (!name) {
      nameInput.focus();
      return;
    }

    var btn = document.querySelector('.campaign-new-btn');
    btn.disabled = true;
    btn.textContent = 'Checking...';

    if (campaignStatsCache && campaignStatsCache.length > 0) {
      checkDuplicateAndCreate(name, subject, btn, nameInput, subjectInput, campaignStatsCache);
    } else {
      google.script.run
        .withSuccessHandler(function(stats) {
          campaignStatsCache = stats || [];
          btn.textContent = 'Creating...';
          checkDuplicateAndCreate(name, subject, btn, nameInput, subjectInput, campaignStatsCache);
        })
        .withFailureHandler(function(err) {
          btn.disabled = false;
          btn.textContent = '+ Create';
          showCreateError(btn, 'Could not verify campaign name. Please try again.');
        })
        .getCampaignStats();
    }
  };

  function loadCampaignStats() {
    if (campaignStatsCache) {
      renderCampaignTable(campaignStatsCache);
      return;
    }
    document.getElementById('campaignStatsLoading').style.display = 'block';
    document.getElementById('campaignTable').style.display = 'none';

    google.script.run
      .withSuccessHandler(function(stats) {
        campaignStatsCache = stats || [];
        currentSortKey = 'dateSent';
        currentSortDir = 'desc';
        sortCampaigns(campaignStatsCache);
        renderCampaignTable(campaignStatsCache);
      })
      .withFailureHandler(function(err) {
        console.error('Failed to load campaign stats:', err);
        document.getElementById('campaignStatsLoading').textContent = 'Failed to load campaigns.';
      })
      .getCampaignStats();
  }

  function sortCampaigns(arr) {
    arr.sort(function(a, b) {
      var valA, valB;
      if (currentSortKey === 'dateSent') {
        valA = a.dateSent || '';
        valB = b.dateSent || '';
      } else if (currentSortKey === 'name') {
        valA = (a.name || '').toLowerCase();
        valB = (b.name || '').toLowerCase();
      } else if (currentSortKey === 'subject') {
        valA = (a.subject || '').toLowerCase();
        valB = (b.subject || '').toLowerCase();
      } else if (currentSortKey === 'totalSent') {
        valA = a.totalSent || 0;
        valB = b.totalSent || 0;
      } else if (currentSortKey === 'uniqueOpens') {
        valA = a.uniqueOpens || 0;
        valB = b.uniqueOpens || 0;
      } else if (currentSortKey === 'uniqueClicks') {
        valA = a.uniqueClicks || 0;
        valB = b.uniqueClicks || 0;
      } else if (currentSortKey === 'unsubscribed') {
        valA = a.unsubscribed || 0;
        valB = b.unsubscribed || 0;
      } else {
        valA = a[currentSortKey] || '';
        valB = b[currentSortKey] || '';
      }

      var cmp;
      if (typeof valA === 'number' && typeof valB === 'number') {
        cmp = valA - valB;
      } else {
        cmp = String(valA).localeCompare(String(valB));
      }
      return currentSortDir === 'desc' ? -cmp : cmp;
    });
  }

  window.sortCampaignsByHeader = function(key) {
    if (currentSortKey === key) {
      currentSortDir = currentSortDir === 'asc' ? 'desc' : 'asc';
    } else {
      currentSortKey = key;
      currentSortDir = (key === 'dateSent') ? 'desc' : 'asc';
    }
    if (campaignStatsCache) {
      sortCampaigns(campaignStatsCache);
      renderCampaignTable(campaignStatsCache);
    }
  };

  function sortArrow(key) {
    if (currentSortKey !== key) return '';
    return currentSortDir === 'asc' ? ' \u25B2' : ' \u25BC';
  }

  function renderCampaignTable(stats) {
    document.getElementById('campaignStatsLoading').style.display = 'none';
    var table = document.getElementById('campaignTable');
    var tbody = document.getElementById('campaignTableBody');

    if (!stats || stats.length === 0) {
      table.style.display = 'none';
      var loadingEl = document.getElementById('campaignStatsLoading');
      loadingEl.style.display = 'block';
      loadingEl.textContent = 'No campaigns found.';
      return;
    }

    var thead = table.querySelector('thead tr');
    thead.innerHTML =
      '<th class="sortable-header" onclick="sortCampaignsByHeader(\'dateSent\')">Date' + sortArrow('dateSent') + '</th>' +
      '<th class="sortable-header" onclick="sortCampaignsByHeader(\'name\')">Name' + sortArrow('name') + '</th>' +
      '<th class="sortable-header" onclick="sortCampaignsByHeader(\'subject\')">Subject' + sortArrow('subject') + '</th>' +
      '<th class="sortable-header" onclick="sortCampaignsByHeader(\'totalSent\')" title="Delivered / Total Sent">Sent' + sortArrow('totalSent') + '</th>' +
      '<th class="sortable-header" onclick="sortCampaignsByHeader(\'uniqueOpens\')" title="Unique Opens / Delivered">Opens' + sortArrow('uniqueOpens') + '</th>' +
      '<th class="sortable-header" onclick="sortCampaignsByHeader(\'uniqueClicks\')" title="Unique Clicks / Unique Opens">Clicks' + sortArrow('uniqueClicks') + '</th>' +
      '<th class="sortable-header" onclick="sortCampaignsByHeader(\'unsubscribed\')" title="Unsubscribes / Delivered">Unsubs' + sortArrow('unsubscribed') + '</th>';

    var html = '';
    for (var i = 0; i < stats.length; i++) {
      var c = stats[i];
      var dateDisplay = formatCampaignDate(c.dateSent);
      var deliveryRate = c.totalSent > 0 ? (c.delivered / c.totalSent * 100) : 0;
      var openRate = c.delivered > 0 ? (c.uniqueOpens / c.delivered * 100) : 0;
      var engagementRate = c.uniqueOpens > 0 ? (c.uniqueClicks / c.uniqueOpens * 100) : 0;
      var unsubRate = c.delivered > 0 ? (c.unsubscribed / c.delivered * 100) : 0;

      html +=
        '<tr class="campaign-row" onclick="openCampaignDetail(\'' + escapeAttr(c.id) + '\', \'' + escapeAttr(c.name) + '\')">' +
          '<td>' + escapeHtml(dateDisplay) + '</td>' +
          '<td class="campaign-name">' + escapeHtml(c.name) + '</td>' +
          '<td class="campaign-subject">' + escapeHtml(c.subject) + '</td>' +
          '<td>' + rateSpan(deliveryRate, 95, 85) + '</td>' +
          '<td>' + rateSpan(openRate, 30, 15) + '</td>' +
          '<td>' + rateSpan(engagementRate, 20, 5) + '</td>' +
          '<td>' + rateSpan(unsubRate, -1, 1, true) + '</td>' +
        '</tr>';
    }
    tbody.innerHTML = html;
    table.style.display = 'table';
  }

  function rateSpan(value, goodThresh, okThresh, invert) {
    var pct = value.toFixed(1) + '%';
    var cls = 'campaign-rate ';
    if (invert) {
      cls += value <= 0.5 ? 'campaign-rate-good' : (value <= okThresh ? 'campaign-rate-ok' : 'campaign-rate-bad');
    } else {
      cls += value >= goodThresh ? 'campaign-rate-good' : (value >= okThresh ? 'campaign-rate-ok' : 'campaign-rate-bad');
    }
    return '<span class="' + cls + '">' + pct + '</span>';
  }

  window.openCampaignDetail = function(campaignId, campaignName) {
    currentDetailCampaignName = campaignName;
    document.getElementById('campaignDashboard').style.display = 'none';
    document.getElementById('campaignDetail').style.display = 'block';
    document.getElementById('campaignDetailTitle').textContent = campaignName;
    document.getElementById('campaignDetailLoading').style.display = 'block';
    document.getElementById('campaignDetailTable').style.display = 'none';
    document.getElementById('campaignDetailEmpty').style.display = 'none';

    currentDetailFilter = 'all';
    var btns = document.querySelectorAll('.campaign-filter-btn');
    btns.forEach(function(b) { b.classList.toggle('active', b.dataset.filter === 'all'); });

    google.script.run
      .withSuccessHandler(function(logs) {
        currentDetailLogs = logs || [];
        renderDetailTable();
      })
      .withFailureHandler(function(err) {
        console.error('Failed to load campaign logs:', err);
        document.getElementById('campaignDetailLoading').textContent = 'Failed to load recipients.';
      })
      .getCampaignLogs(campaignId, campaignName);
  };

  window.showCampaignDashboard = function() {
    document.getElementById('campaignDetail').style.display = 'none';
    document.getElementById('campaignDashboard').style.display = 'block';
    currentDetailLogs = [];
  };

  window.filterCampaignLogs = function(filter) {
    currentDetailFilter = filter;
    var btns = document.querySelectorAll('.campaign-filter-btn');
    btns.forEach(function(b) { b.classList.toggle('active', b.dataset.filter === filter); });
    renderDetailTable();
  };

  function renderDetailTable() {
    document.getElementById('campaignDetailLoading').style.display = 'none';
    var table = document.getElementById('campaignDetailTable');
    var tbody = document.getElementById('campaignDetailBody');
    var emptyEl = document.getElementById('campaignDetailEmpty');

    var filtered = currentDetailLogs;
    if (currentDetailFilter !== 'all') {
      filtered = currentDetailLogs.filter(function(l) {
        return l.event.toLowerCase() === currentDetailFilter;
      });
    }

    if (filtered.length === 0) {
      table.style.display = 'none';
      emptyEl.style.display = 'block';
      return;
    }

    emptyEl.style.display = 'none';
    var html = '';
    for (var i = 0; i < filtered.length; i++) {
      var log = filtered[i];
      var nameCell;
      if (log.contactName) {
        nameCell = '<a class="contact-link" onclick="navigateToContactFromCampaign(\'' + escapeAttr(log.contactId) + '\')">' + escapeHtml(log.contactName) + '</a>';
      } else if (log.email) {
        nameCell = '<span class="campaign-detail-email-name">' + escapeHtml(log.email) + '</span>';
      } else {
        nameCell = '<span class="campaign-detail-unknown">Unknown</span>';
      }
      var statusLower = (log.event || '').toLowerCase();
      var badgeCls = 'campaign-status-badge campaign-status-' + statusLower;
      var timeDisplay = formatLogTime(log.timestamp);

      var actionCell = '';
      if (statusLower === 'unsubscribed' && log.email) {
        var greeting = log.contactName ? log.contactName.split(' ')[0] : '';
        var gmailSubject = encodeURIComponent('Quick check re: your unsubscribe');
        var gmailBody = encodeURIComponent('Hi' + (greeting ? ' ' + greeting : '') + ',\n\nJust taking as much care as we can with regard to your unsubscribe just now - which is totally fine by the way!\n\nIf you would prefer to keep hearing from us but via another email address, please reply and let me know.\n\nOtherwise, if you do nothing, you\'ll remain unsubscribed.\n\nAll the very best either way!\n\nCheers,');
        var gmailUrl = 'https://mail.google.com/mail/?view=cm&to=' + encodeURIComponent(log.email) + '&su=' + gmailSubject + '&body=' + gmailBody;
        actionCell = '<a class="campaign-followup-btn" href="' + gmailUrl + '" target="_blank" title="Send follow-up email">&#9993;</a>';
      }

      html +=
        '<tr>' +
          '<td>' + nameCell + '</td>' +
          '<td>' + escapeHtml(log.email) + '</td>' +
          '<td><span class="' + badgeCls + '">' + escapeHtml(log.event) + '</span></td>' +
          '<td>' + escapeHtml(timeDisplay) + '</td>' +
          '<td>' + actionCell + '</td>' +
        '</tr>';
    }
    tbody.innerHTML = html;
    table.style.display = 'table';
  }

  window.navigateToContactFromCampaign = function(contactId) {
    if (!contactId) return;
    closeMarketingModal();
    if (window.loadContactById) {
      loadContactById(contactId, true);
    }
  };

  function loadMarketingData() {
    if (isLoading) return;
    if (cachedExportData.length > 0) {
      renderStats();
      return;
    }
    isLoading = true;

    var statsEl = document.getElementById('marketingStats');
    var downloadBtn = document.getElementById('marketingDownloadBtn');
    statsEl.innerHTML = '<span style="color:#999;">Loading contacts...</span>';
    downloadBtn.disabled = true;

    google.script.run
      .withSuccessHandler(function(contacts) {
        allContacts = (contacts || []).map(function(c) {
          return {
            id: c.id,
            calculatedName: c.calculatedName || '',
            email: c.email || '',
            unsubscribed: c.unsubscribed || false,
            inactive: (c.status || 'Active') === 'Inactive',
            deceased: c.deceased || false
          };
        });
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
    if (campaigns.length > 0) {
      return;
    }
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
      return !c.unsubscribed && !c.inactive && !c.deceased && c.email && c.email.trim() !== '';
    });

    var grouped = {};
    marketable.forEach(function(contact) {
      var emailKey = contact.email.trim().toLowerCase();
      if (!grouped[emailKey]) {
        grouped[emailKey] = { names: [], ids: [], email: contact.email.trim() };
      }
      var name = contact.calculatedName || '';
      if (name) grouped[emailKey].names.push(name);
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
        combinedName = group.names.join(' & ');
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
    var inactiveCount = 0;
    var deceasedCount = 0;
    var unsubCount = 0;
    var totalExcludedCount = 0;

    allContacts.forEach(function(c) {
      var isInactive = c.inactive;
      var isDeceased = c.deceased;
      var isUnsub = c.unsubscribed;
      if (isInactive) inactiveCount++;
      if (isDeceased) deceasedCount++;
      if (isUnsub) unsubCount++;
      if (isInactive || isDeceased || isUnsub) totalExcludedCount++;
    });

    var marketable = total - totalExcludedCount;
    var exportCount = cachedExportData.length;
    var deduped = marketable - exportCount;

    var excludedDetail = inactiveCount + ' Inactive, ' + deceasedCount + ' Deceased, ' + unsubCount + ' Unsubscribed';

    var statsEl = document.getElementById('marketingStats');
    statsEl.innerHTML =
      '<div class="marketing-stat-row">' +
        '<span class="marketing-stat-label">Total Contacts</span>' +
        '<span class="marketing-stat-value">' + formatNumber(total) + '</span>' +
      '</div>' +
      '<div class="marketing-stat-row marketing-stat-deduction">' +
        '<span class="marketing-stat-label">Less: Excluded (' + excludedDetail + ')</span>' +
        '<span class="marketing-stat-value">\u2212 ' + formatNumber(totalExcludedCount) + '</span>' +
      '</div>' +
      '<div class="marketing-stat-row marketing-stat-highlight marketing-stat-divider">' +
        '<span class="marketing-stat-label">= Marketable Contacts</span>' +
        '<span class="marketing-stat-value marketing-stat-green">' + formatNumber(marketable) + '</span>' +
      '</div>' +
      '<div class="marketing-stat-row marketing-stat-deduction">' +
        '<span class="marketing-stat-label">Less: Deduplication (only one row per email address)</span>' +
        '<span class="marketing-stat-value">\u2212 ' + formatNumber(deduped > 0 ? deduped : 0) + '</span>' +
      '</div>' +
      '<div class="marketing-stat-row marketing-stat-export">' +
        '<span class="marketing-stat-label">= Ready to Send</span>' +
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

    var csvContent = '\uFEFF' + csvLines.join('\n');
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

    var timestampIdx = -1;
    for (var t = 0; t < headers.length; t++) {
      if (headers[t] === 'date' || headers[t] === 'timestamp') {
        timestampIdx = t;
        break;
      }
    }

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
      var timestamp = (timestampIdx !== -1 && cols[timestampIdx]) ? cols[timestampIdx].trim() : '';
      if (!integrityId) continue;
      rows.push({ email: email, status: status, integrityId: integrityId, timestamp: timestamp });
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
          campaignStatsCache = null;
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

  function formatCampaignDate(dateStr) {
    if (!dateStr) return '';
    try {
      var d = new Date(dateStr);
      if (isNaN(d.getTime())) return dateStr;
      var day = String(d.getDate()).padStart(2, '0');
      var month = String(d.getMonth() + 1).padStart(2, '0');
      var year = d.getFullYear();
      return day + '/' + month + '/' + year;
    } catch (e) {
      return dateStr;
    }
  }

  function formatLogTime(ts) {
    if (!ts) return '';
    try {
      var d = new Date(ts);
      if (isNaN(d.getTime())) return ts;
      var day = String(d.getDate()).padStart(2, '0');
      var month = String(d.getMonth() + 1).padStart(2, '0');
      var hours = d.getHours();
      var minutes = String(d.getMinutes()).padStart(2, '0');
      var ampm = hours >= 12 ? 'PM' : 'AM';
      var h12 = hours % 12 || 12;
      return day + '/' + month + ' ' + h12 + ':' + minutes + ' ' + ampm;
    } catch (e) {
      return ts;
    }
  }

  function escapeHtml(str) {
    if (!str) return '';
    return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function escapeAttr(str) {
    if (!str) return '';
    return str.replace(/'/g, "\\'").replace(/"/g, '&quot;');
  }

  function escapeCsvField(value) {
    if (!value) return '';
    if (value.indexOf(',') !== -1 || value.indexOf('"') !== -1 || value.indexOf('\n') !== -1) {
      return '"' + value.replace(/"/g, '""') + '"';
    }
    return value;
  }

  window.openFollowUpComposer = function(email, contactName) {
    var greeting = contactName ? contactName.split(' ')[0] : '';
    var defaultSubject = 'Quick check re: your unsubscribe';
    var defaultBody = 'Hi' + (greeting ? ' ' + greeting : '') + ',\n\nJust taking as much care as we can with regard to your unsubscribe just now - which is totally fine by the way!\n\nIf you would prefer to keep hearing from us but via another email address, please reply and let me know.\n\nOtherwise, if you do nothing, you\'ll remain unsubscribed.\n\nAll the very best either way!\n\nCheers,';

    var existing = document.getElementById('followUpComposerModal');
    if (existing) existing.remove();

    var modal = document.createElement('div');
    modal.id = 'followUpComposerModal';
    modal.className = 'modal visible showing';
    modal.style.zIndex = '100001';
    modal.innerHTML =
      '<div class="followup-composer-panel">' +
        '<div class="followup-composer-header">' +
          '<h3>Follow-Up Email</h3>' +
          '<button class="followup-close-btn" onclick="closeFollowUpComposer()">&times;</button>' +
        '</div>' +
        '<div class="followup-composer-body">' +
          '<label>To</label>' +
          '<input type="text" id="followUpTo" value="' + escapeAttr(email) + '" readonly />' +
          '<label>Subject</label>' +
          '<input type="text" id="followUpSubject" value="' + escapeAttr(defaultSubject) + '" />' +
          '<label>Message</label>' +
          '<textarea id="followUpBody" rows="10">' + escapeHtml(defaultBody) + '</textarea>' +
          '<div id="followUpStatus" style="margin-top:8px;font-size:12px;"></div>' +
          '<div class="followup-composer-actions">' +
            '<button class="btn-cancel" onclick="closeFollowUpComposer()">Cancel</button>' +
            '<button class="btn-confirm" id="followUpSendBtn" onclick="sendFollowUpEmail()">Send</button>' +
          '</div>' +
        '</div>' +
      '</div>';
    document.body.appendChild(modal);
  };

  window.closeFollowUpComposer = function() {
    var modal = document.getElementById('followUpComposerModal');
    if (modal) modal.remove();
  };

  window.sendFollowUpEmail = function() {
    var to = document.getElementById('followUpTo').value.trim();
    var subject = document.getElementById('followUpSubject').value.trim();
    var body = document.getElementById('followUpBody').value.trim();
    var statusEl = document.getElementById('followUpStatus');
    var btn = document.getElementById('followUpSendBtn');

    if (!to || !subject || !body) {
      statusEl.style.color = '#dc3545';
      statusEl.textContent = 'Please fill in all fields.';
      return;
    }

    btn.disabled = true;
    btn.textContent = 'Sending...';
    statusEl.style.color = '#666';
    statusEl.textContent = '';

    var htmlBody = '<div style="font-family:sans-serif;font-size:14px;line-height:1.6;">' +
      body.replace(/\n/g, '<br>') + '</div>';

    google.script.run
      .withSuccessHandler(function(result) {
        if (result && result.success) {
          statusEl.style.color = '#28a745';
          statusEl.textContent = 'Email sent successfully!';
          btn.textContent = 'Sent';
          setTimeout(function() { closeFollowUpComposer(); }, 1500);
        } else {
          statusEl.style.color = '#dc3545';
          statusEl.textContent = 'Failed: ' + ((result && result.error) || 'Unknown error');
          btn.disabled = false;
          btn.textContent = 'Send';
        }
      })
      .withFailureHandler(function(err) {
        statusEl.style.color = '#dc3545';
        statusEl.textContent = 'Error: ' + (err.message || err);
        btn.disabled = false;
        btn.textContent = 'Send';
      })
      .sendEmail(to, subject, htmlBody);
  };

})();
