(function() {
  'use strict';

  let allContacts = [];
  let cachedExportData = [];
  let isLoading = false;

  window.openMarketingModal = function() {
    openModal('marketingModal');
    loadMarketingData();
  };

  window.closeMarketingModal = function() {
    closeModal('marketingModal');
  };

  function loadMarketingData() {
    if (isLoading) return;
    isLoading = true;

    const statsEl = document.getElementById('marketingStats');
    const downloadBtn = document.getElementById('marketingDownloadBtn');
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

  function escapeCsvField(value) {
    if (!value) return '';
    if (value.indexOf(',') !== -1 || value.indexOf('"') !== -1 || value.indexOf('\n') !== -1) {
      return '"' + value.replace(/"/g, '""') + '"';
    }
    return value;
  }

})();
