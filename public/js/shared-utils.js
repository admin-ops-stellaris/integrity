/**
 * Shared Utilities Module
 * Pure utility functions with NO dependencies on other app code
 * 
 * IMPORTANT: This must be loaded SECOND, after shared-state.js
 * These are "leaf" functions - they don't call other app functions
 */
(function() {
  'use strict';
  
  // ============================================================
  // HTML Escape/Unescape Functions
  // ============================================================
  
  window.escapeHtml = function(str) {
    if (!str) return '';
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
  };
  
  window.escapeHtmlForAttr = function(str) {
    if (!str) return '';
    return str.replace(/'/g, "&#39;").replace(/"/g, '&quot;');
  };
  
  window.unescapeHtml = function(str) {
    if (!str) return '';
    return str.replace(/&#39;/g, "'").replace(/&quot;/g, '"');
  };
  
  // ============================================================
  // Date/Time Formatting - Ordinal suffix helper
  // ============================================================
  
  window.getOrdinalSuffix = function(day) {
    if (day > 3 && day < 21) return 'th';
    switch (day % 10) {
      case 1: return 'st';
      case 2: return 'nd';
      case 3: return 'rd';
      default: return 'th';
    }
  };
  
  // ============================================================
  // Appointment Time Formatting (Perth timezone)
  // ============================================================
  
  window.formatAppointmentTime = function(dateStr) {
    if (!dateStr) return '[appointment time]';
    
    try {
      if (dateStr.includes('T') || dateStr.includes('Z')) {
        const date = new Date(dateStr);
        if (isNaN(date.getTime())) return dateStr;
        
        const options = {
          weekday: 'long',
          day: 'numeric',
          month: 'long',
          hour: 'numeric',
          minute: '2-digit',
          hour12: true,
          timeZone: 'Australia/Perth'
        };
        
        const formatted = date.toLocaleString('en-AU', options);
        
        const dayMatch = formatted.match(/(\d+)/);
        if (dayMatch) {
          const day = parseInt(dayMatch[1]);
          const suffix = window.getOrdinalSuffix(day);
          return formatted.replace(/(\d+)/, `${day}${suffix}`);
        }
        return formatted;
      }
      
      return dateStr;
    } catch (e) {
      console.error('Error formatting date:', e);
      return dateStr;
    }
  };
  
  // ============================================================
  // Calculate Days Until Appointment
  // ============================================================
  
  window.calculateDaysUntil = function(appointmentTimeStr) {
    if (!appointmentTimeStr) return '?';
    
    try {
      let apptDate;
      
      if (appointmentTimeStr.includes('T') || appointmentTimeStr.includes('Z')) {
        apptDate = new Date(appointmentTimeStr);
        if (isNaN(apptDate.getTime())) return '?';
      } else {
        const months = {
          'january': 0, 'jan': 0,
          'february': 1, 'feb': 1,
          'march': 2, 'mar': 2,
          'april': 3, 'apr': 3,
          'may': 4,
          'june': 5, 'jun': 5,
          'july': 6, 'jul': 6,
          'august': 7, 'aug': 7,
          'september': 8, 'sep': 8, 'sept': 8,
          'october': 9, 'oct': 9,
          'november': 10, 'nov': 10,
          'december': 11, 'dec': 11
        };
        
        const cleanStr = appointmentTimeStr.toLowerCase().replace(/,/g, '');
        
        const dayMatch = cleanStr.match(/(\d{1,2})(st|nd|rd|th)?/);
        if (!dayMatch) return '?';
        const day = parseInt(dayMatch[1]);
        
        let monthIdx = -1;
        for (const [name, idx] of Object.entries(months)) {
          if (cleanStr.includes(name)) {
            monthIdx = idx;
            break;
          }
        }
        if (monthIdx === -1) return '?';
        
        const now = new Date();
        let year = now.getFullYear();
        apptDate = new Date(year, monthIdx, day);
        
        if (apptDate < now) {
          apptDate = new Date(year + 1, monthIdx, day);
        }
      }
      
      const now = new Date();
      const diffMs = apptDate - now;
      const diffDays = Math.ceil(diffMs / (1000 * 60 * 60 * 24));
      
      return diffDays;
    } catch (e) {
      console.error('Error calculating days until:', e);
      return '?';
    }
  };
  
  // ============================================================
  // Datetime Input/Display Formatting
  // ============================================================
  
  window.formatDatetimeForInput = function(dateStr) {
    if (!dateStr) return '';
    try {
      const d = new Date(dateStr);
      if (isNaN(d.getTime())) return '';
      return d.toISOString().slice(0, 16);
    } catch (e) {
      return '';
    }
  };
  
  window.formatDatetimeForDisplay = function(dateStr) {
    if (!dateStr) return 'Time not set';
    try {
      const d = new Date(dateStr);
      if (isNaN(d.getTime())) return dateStr;
      return d.toLocaleString('en-AU', { 
        weekday: 'short', 
        day: '2-digit', 
        month: '2-digit', 
        year: '2-digit',
        hour: 'numeric',
        minute: '2-digit',
        hour12: true 
      });
    } catch (e) {
      return dateStr;
    }
  };
  
  // ============================================================
  // Date Display Formatting (DD/MM/YYYY)
  // ============================================================
  
  window.formatDateDisplay = function(isoDate) {
    if (!isoDate) return '';
    const parts = isoDate.split('-');
    if (parts.length !== 3) return isoDate;
    return `${parts[2]}/${parts[1]}/${parts[0]}`;
  };
  
  window.parseDateInput = function(value) {
    if (!value) return null;
    const match = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (match) {
      return `${match[3]}-${match[2].padStart(2, '0')}-${match[1].padStart(2, '0')}`;
    }
    return value;
  };
  
  // ============================================================
  // Smart Date Parsing - Flexible Input Formats
  // ============================================================
  
  /**
   * Parse flexible date input and return structured result
   * Handles: DDMMYY, DDMMYYYY, DD/MM/YY, DD/MM/YYYY, DD.MM.YY, DD.MM.YYYY
   * @param {string} inputStr - Raw user input
   * @returns {object|null} { iso: 'YYYY-MM-DD', display: 'DD/MM/YYYY' } or null if invalid
   */
  window.parseFlexibleDate = function(inputStr) {
    if (!inputStr || typeof inputStr !== 'string') return null;
    
    const str = inputStr.trim();
    if (!str) return null;
    
    let day, month, year;
    
    // Pattern 1: No separators - DDMMYY or DDMMYYYY
    const noSepMatch = str.match(/^(\d{2})(\d{2})(\d{2,4})$/);
    if (noSepMatch) {
      day = parseInt(noSepMatch[1], 10);
      month = parseInt(noSepMatch[2], 10);
      year = noSepMatch[3];
    }
    
    // Pattern 2: With separators (/ or . or -) - D/M/YY or DD/MM/YYYY etc
    if (!day) {
      const sepMatch = str.match(/^(\d{1,2})[\/\.\-](\d{1,2})[\/\.\-](\d{2,4})$/);
      if (sepMatch) {
        day = parseInt(sepMatch[1], 10);
        month = parseInt(sepMatch[2], 10);
        year = sepMatch[3];
      }
    }
    
    // If still no match, return null
    if (!day) return null;
    
    // Convert 2-digit year to 4-digit (00-49 = 2000s, 50-99 = 1900s)
    if (year.length === 2) {
      const yy = parseInt(year, 10);
      year = yy < 50 ? 2000 + yy : 1900 + yy;
    } else {
      year = parseInt(year, 10);
    }
    
    // Validate date components
    if (month < 1 || month > 12) return null;
    if (day < 1 || day > 31) return null;
    if (year < 1900 || year > 2100) return null;
    
    // Check if date is actually valid (e.g., Feb 30 is invalid)
    const testDate = new Date(year, month - 1, day);
    if (testDate.getFullYear() !== year || 
        testDate.getMonth() !== month - 1 || 
        testDate.getDate() !== day) {
      return null;
    }
    
    // Format output
    const dd = String(day).padStart(2, '0');
    const mm = String(month).padStart(2, '0');
    const yyyy = String(year);
    
    return {
      iso: `${yyyy}-${mm}-${dd}`,
      display: `${dd}/${mm}/${yyyy}`
    };
  };
  
  window.formatAddressDateRange = function(from, to) {
    const fromStr = from ? window.formatDateDisplay(from) : '?';
    const toStr = to ? window.formatDateDisplay(to) : 'Present';
    return `${fromStr} - ${toStr}`;
  };
  
})();
