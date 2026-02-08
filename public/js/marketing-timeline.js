(function() {
  'use strict';

  var EVENT_CONFIG = {
    'Opened':       { icon: '📬', color: '#16a34a' },
    'Clicked':      { icon: '🔗', color: '#2563eb' },
    'Delivered':    { icon: '✉️', color: '#6b7280' },
    'Sent':         { icon: '📤', color: '#6b7280' },
    'Bounced':      { icon: '⚠️', color: '#dc2626' },
    'Unsubscribed': { icon: '🚫', color: '#dc2626' },
    'Complained':   { icon: '🚩', color: '#dc2626' }
  };

  window.loadMarketingTimeline = function(contactId) {
    var container = document.getElementById('marketingTimeline');
    if (!container) return;

    container.innerHTML =
      '<div class="mt-loading">' +
        '<span class="mt-spinner"></span> Loading marketing history...' +
      '</div>';

    google.script.run
      .withSuccessHandler(function(logs) {
        renderTimeline(logs || []);
      })
      .withFailureHandler(function(err) {
        console.error('Failed to load marketing timeline:', err);
        container.innerHTML =
          '<p class="mt-empty">Failed to load marketing history.</p>';
      })
      .getMarketingLogsForContact(contactId);
  };

  window.clearMarketingTimeline = function() {
    var container = document.getElementById('marketingTimeline');
    if (container) {
      container.innerHTML =
        '<p class="mt-empty">No marketing history.</p>';
    }
  };

  function renderTimeline(logs) {
    var container = document.getElementById('marketingTimeline');
    if (!container) return;

    if (logs.length === 0) {
      container.innerHTML =
        '<p class="mt-empty">No marketing history found.</p>';
      return;
    }

    var html = '<ul class="mt-list">';
    for (var i = 0; i < logs.length; i++) {
      var log = logs[i];
      var config = EVENT_CONFIG[log.event] || { icon: '📋', color: '#6b7280' };
      var timeDisplay = formatLogTimestamp(log.timestamp);
      var campaign = log.campaignName || 'Unknown Campaign';

      html +=
        '<li class="mt-item">' +
          '<div class="mt-icon" style="color:' + config.color + ';">' + config.icon + '</div>' +
          '<div class="mt-content">' +
            '<div class="mt-event" style="color:' + config.color + ';">' + escapeHtml(log.event) + '</div>' +
            '<div class="mt-campaign">' + escapeHtml(campaign) + '</div>' +
            '<div class="mt-time">' + escapeHtml(timeDisplay) + '</div>' +
          '</div>' +
        '</li>';
    }
    html += '</ul>';
    container.innerHTML = html;
  }

  function formatLogTimestamp(ts) {
    if (!ts) return '';
    try {
      var d = new Date(ts);
      if (isNaN(d.getTime())) return ts;
      var day = String(d.getDate()).padStart(2, '0');
      var month = String(d.getMonth() + 1).padStart(2, '0');
      var year = d.getFullYear();
      var hours = d.getHours();
      var minutes = String(d.getMinutes()).padStart(2, '0');
      var ampm = hours >= 12 ? 'PM' : 'AM';
      var h12 = hours % 12 || 12;
      return day + '/' + month + '/' + year + ' ' + h12 + ':' + minutes + ' ' + ampm;
    } catch (e) {
      return ts;
    }
  }

  function escapeHtml(str) {
    if (!str) return '';
    return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

})();
