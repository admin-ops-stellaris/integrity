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
      // Check if it's a floating ISO (our saved data) vs UTC Z (external)
      const isUTC = dateStr.endsWith('Z');
      
      if (isUTC) {
        // UTC string - convert to Perth time
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
      } else if (dateStr.includes('T')) {
        // Floating ISO string - parse directly without conversion
        const match = dateStr.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/);
        if (!match) return dateStr;
        
        const [, year, month, day, hour, minute] = match;
        const d = parseInt(day, 10);
        const h = parseInt(hour, 10);
        
        // Create date for weekday/month formatting
        const date = new Date(parseInt(year), parseInt(month) - 1, d);
        const weekday = date.toLocaleDateString('en-AU', { weekday: 'long' });
        const monthName = date.toLocaleDateString('en-AU', { month: 'long' });
        
        const suffix = window.getOrdinalSuffix(d);
        const displayHour = h % 12 || 12;
        const ampm = h < 12 ? 'am' : 'pm';
        
        return `${weekday} ${d}${suffix} ${monthName} ${displayHour}:${minute} ${ampm}`;
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
      
      // Check if it's an ISO datetime string
      if (appointmentTimeStr.includes('T')) {
        // Use parseFloatingDate for consistent handling (falls back to Date parsing for Z strings)
        apptDate = window.parseFloatingDate ? window.parseFloatingDate(appointmentTimeStr) : null;
        if (!apptDate) {
          apptDate = new Date(appointmentTimeStr);
        }
        if (isNaN(apptDate.getTime())) return '?';
      } else if (appointmentTimeStr.includes('Z')) {
        // UTC string without T (unlikely but handle it)
        apptDate = new Date(appointmentTimeStr);
        if (isNaN(apptDate.getTime())) return '?';
      } else {
        // Legacy text format parsing
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
  
  // DEPRECATED: Use formatDatetimeForInput in appointments.js or parseFloatingDate
  // Kept for backward compatibility but now parses floating ISO directly
  window.formatDatetimeForInput = function(dateStr) {
    if (!dateStr) return '';
    
    // For floating ISO strings, parse directly without Date conversion
    const match = dateStr.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/);
    if (match) {
      return `${match[1]}-${match[2]}-${match[3]}T${match[4]}:${match[5]}`;
    }
    
    // Fallback for other formats
    try {
      const d = new Date(dateStr);
      if (isNaN(d.getTime())) return '';
      const year = d.getFullYear();
      const month = String(d.getMonth() + 1).padStart(2, '0');
      const day = String(d.getDate()).padStart(2, '0');
      const hours = String(d.getHours()).padStart(2, '0');
      const mins = String(d.getMinutes()).padStart(2, '0');
      return `${year}-${month}-${day}T${hours}:${mins}`;
    } catch (e) {
      return '';
    }
  };
  
  // DEPRECATED: Use window.formatDateTimeForDisplay (Perth Standard) instead
  // Kept for backward compatibility, now delegates to Perth Standard
  window.formatDatetimeForDisplay = function(dateStr) {
    if (!dateStr) return 'Time not set';
    // Delegate to Perth Standard formatter
    const result = window.formatDateTimeForDisplay ? 
      window.formatDateTimeForDisplay(dateStr, { format: 'long' }) : null;
    if (result) return result;
    
    // Fallback if Perth Standard not loaded yet
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
  
  /**
   * Parse date input value to ISO format for API/storage
   * Checks for data-iso-date attribute first (set by smart date listener)
   * Falls back to parsing DD/MM/YYYY or returning value as-is
   * @param {string} value - Display value from input
   * @param {HTMLElement} [inputEl] - Optional input element to check dataset
   * @returns {string|null} ISO date (YYYY-MM-DD) or original value
   */
  window.parseDateInput = function(value, inputEl) {
    if (!value) return null;
    
    // Check for pre-parsed ISO from smart date listener
    if (inputEl && inputEl.dataset && inputEl.dataset.isoDate) {
      return inputEl.dataset.isoDate;
    }
    
    // Try flexible parsing first (handles all formats)
    const parsed = window.parseFlexibleDate(value);
    if (parsed) {
      return parsed.iso;
    }
    
    // Fallback: legacy DD/MM/YYYY pattern
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
  
  // ============================================================
  // Smart Time Parsing - Flexible Input Formats
  // ============================================================
  
  /**
   * Parse flexible time input and return structured result
   * Handles: 1300, 130, 1:30, 13, 1p, 1pm, 1:30pm, 13:00
   * @param {string} inputStr - Raw user input
   * @returns {object|null} { value24: 'HH:MM', display: 'h:mm AM/PM' } or null if invalid
   */
  window.parseFlexibleTime = function(inputStr) {
    if (!inputStr || typeof inputStr !== 'string') return null;
    
    const str = inputStr.trim().toLowerCase();
    if (!str) return null;
    
    let hours = null;
    let minutes = 0;
    let isPM = false;
    let isAM = false;
    
    // Check for AM/PM suffix
    if (str.includes('p')) isPM = true;
    if (str.includes('a')) isAM = true;
    
    // Remove am/pm suffixes for parsing
    const numPart = str.replace(/[ap]m?/gi, '').trim();
    
    // Pattern 1: HH:MM or H:MM format (e.g., 13:00, 1:30)
    const colonMatch = numPart.match(/^(\d{1,2}):(\d{2})$/);
    if (colonMatch) {
      hours = parseInt(colonMatch[1], 10);
      minutes = parseInt(colonMatch[2], 10);
    }
    
    // Pattern 2: 3-4 digit military time (e.g., 1300, 130, 930)
    if (hours === null) {
      const militaryMatch = numPart.match(/^(\d{3,4})$/);
      if (militaryMatch) {
        const num = militaryMatch[1];
        if (num.length === 4) {
          hours = parseInt(num.substring(0, 2), 10);
          minutes = parseInt(num.substring(2, 4), 10);
        } else {
          // 3 digits: first is hour, last two are minutes (e.g., 130 = 1:30)
          hours = parseInt(num.substring(0, 1), 10);
          minutes = parseInt(num.substring(1, 3), 10);
        }
      }
    }
    
    // Pattern 3: 1-2 digit hour only (e.g., 9, 13)
    if (hours === null) {
      const hourMatch = numPart.match(/^(\d{1,2})$/);
      if (hourMatch) {
        hours = parseInt(hourMatch[1], 10);
        minutes = 0;
      }
    }
    
    if (hours === null) return null;
    
    // Apply AM/PM modifier if hours <= 12
    if (isPM && hours < 12) hours += 12;
    if (isAM && hours === 12) hours = 0;
    
    // Validate
    if (hours < 0 || hours > 23) return null;
    if (minutes < 0 || minutes > 59) return null;
    
    // Format output
    const hh = String(hours).padStart(2, '0');
    const mm = String(minutes).padStart(2, '0');
    const value24 = `${hh}:${mm}`;
    
    // 12-hour display format
    let displayHour = hours % 12;
    if (displayHour === 0) displayHour = 12;
    const ampm = hours < 12 ? 'AM' : 'PM';
    const display = `${displayHour}:${mm} ${ampm}`;
    
    return { value24, display };
  };
  
  // ============================================================
  // Perth Standard - Global Date/Time Strategy
  // ============================================================
  
  /**
   * Construct a Floating ISO string for saving to Airtable
   * Returns YYYY-MM-DDTHH:mm:00 (No Z, no timezone offset)
   * This preserves the user's intended local time without browser conversion
   * @param {string} dateStr - ISO date (YYYY-MM-DD)
   * @param {string} timeStr - 24h time (HH:mm) - optional, defaults to 00:00
   * @returns {string} Floating ISO string
   */
  window.constructDateForSave = function(dateStr, timeStr) {
    if (!dateStr) return '';
    const time = timeStr || '00:00';
    return `${dateStr}T${time}:00`;
  };
  
  /**
   * Parse a floating ISO datetime string and return a comparable Date object
   * This interprets the string as local Perth time for comparison purposes
   * @param {string} isoString - Floating ISO string (YYYY-MM-DDTHH:mm:00)
   * @returns {Date|null} Date object or null if invalid
   */
  window.parseFloatingDate = function(isoString) {
    if (!isoString) return null;
    
    // For floating strings, parse components directly to avoid timezone shifts
    const match = isoString.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/);
    if (!match) {
      // Try native Date parsing as fallback
      const d = new Date(isoString);
      return isNaN(d.getTime()) ? null : d;
    }
    
    const [, year, month, day, hour, minute] = match;
    // Create date using local time interpretation
    return new Date(
      parseInt(year, 10),
      parseInt(month, 10) - 1,
      parseInt(day, 10),
      parseInt(hour, 10),
      parseInt(minute, 10)
    );
  };
  
  /**
   * Format any ISO datetime string for display in Perth timezone
   * - If string ends in Z (UTC from Airtable Created/Modified): convert to Perth time
   * - If string is Floating (our saved appointments): display as-is
   * @param {string} isoString - ISO datetime string
   * @param {object} options - Optional formatting options
   * @returns {string} Formatted datetime string for display
   */
  window.formatDateTimeForDisplay = function(isoString, options) {
    if (!isoString) return '';
    
    options = options || {};
    const format = options.format || 'short'; // 'short' or 'long'
    
    try {
      // Only Z suffix indicates true UTC (from Airtable Created/Modified fields)
      // Timezone offsets like +08:00 should be treated as explicit local time
      const isUTC = isoString.endsWith('Z');
      
      if (isUTC) {
        // UTC string from Airtable Created/Modified - convert to Perth time
        const date = new Date(isoString);
        if (isNaN(date.getTime())) return isoString;
        
        if (format === 'long') {
          return date.toLocaleString('en-AU', {
            timeZone: 'Australia/Perth',
            weekday: 'short',
            day: '2-digit',
            month: '2-digit',
            year: '2-digit',
            hour: 'numeric',
            minute: '2-digit',
            hour12: true
          });
        } else {
          return date.toLocaleString('en-AU', {
            timeZone: 'Australia/Perth',
            day: '2-digit',
            month: '2-digit',
            year: '2-digit',
            hour: 'numeric',
            minute: '2-digit',
            hour12: true
          });
        }
      } else {
        // Floating string (our saved data) - parse and display as-is
        // Format: YYYY-MM-DDTHH:mm:00
        const match = isoString.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/);
        if (!match) return isoString;
        
        const [, year, month, day, hour, minute] = match;
        const h = parseInt(hour, 10);
        const displayHour = h % 12 || 12;
        const ampm = h < 12 ? 'AM' : 'PM';
        
        if (format === 'long') {
          const date = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
          const weekday = date.toLocaleDateString('en-AU', { weekday: 'short' });
          return `${weekday} ${day}/${month}/${year.slice(2)} ${displayHour}:${minute} ${ampm}`;
        } else {
          return `${day}/${month}/${year.slice(2)} ${displayHour}:${minute} ${ampm}`;
        }
      }
    } catch (e) {
      console.error('Error formatting datetime:', e);
      return isoString;
    }
  };
  
  /**
   * PERTH STANDARD: Parse ISO datetime for form/editor inputs
   * Converts UTC Z strings to Perth time, keeps floating ISO as-is
   * 
   * @param {string} isoString - ISO datetime string (with Z or floating)
   * @returns {Object} { dateDisplay, timeDisplay, isoDate, time24 } or empty values
   */
  window.parseDateForEditor = function(isoString) {
    const empty = { dateDisplay: '', timeDisplay: '', isoDate: '', time24: '' };
    
    if (!isoString) return empty;
    
    try {
      let year, month, day, hour, minute;
      
      // Check if UTC Z string - need to convert to Perth time first
      if (isoString.endsWith('Z')) {
        const d = new Date(isoString);
        if (isNaN(d.getTime())) return empty;
        
        // Get Perth time components using toLocaleString
        const perthStr = d.toLocaleString('en-AU', { 
          timeZone: 'Australia/Perth',
          year: 'numeric', month: '2-digit', day: '2-digit',
          hour: '2-digit', minute: '2-digit', hour12: false
        });
        // Format: DD/MM/YYYY, HH:mm
        const dateTimeParts = perthStr.split(', ');
        if (dateTimeParts.length === 2) {
          const [datePart, timePart] = dateTimeParts;
          [day, month, year] = datePart.split('/');
          [hour, minute] = timePart.split(':');
        }
      } else {
        // Floating ISO - parse directly without conversion
        const match = isoString.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/);
        if (match) {
          [, year, month, day, hour, minute] = match;
        }
      }
      
      if (!year || !month || !day || hour === undefined || !minute) {
        return empty;
      }
      
      // Build return object
      const h = parseInt(hour, 10);
      const ampm = h >= 12 ? 'PM' : 'AM';
      const displayHour = h % 12 || 12;
      
      return {
        dateDisplay: `${day}/${month}/${year}`,
        timeDisplay: `${displayHour}:${String(minute).padStart(2, '0')} ${ampm}`,
        isoDate: `${year}-${month}-${day}`,
        time24: `${String(h).padStart(2, '0')}:${String(minute).padStart(2, '0')}`
      };
    } catch (e) {
      console.error('Error parsing date for editor:', e);
      return empty;
    }
  };
  
  // ============================================================
  // Breadcrumb Navigation
  // ============================================================
  
  window.renderBreadcrumbs = function(pathArray) {
    if (!pathArray || pathArray.length === 0) return '';
    
    return pathArray.map((item, index) => {
      const isLast = index === pathArray.length - 1;
      const separator = index > 0 ? '<span class="breadcrumb-sep">›</span>' : '';
      
      if (isLast) {
        return `${separator}<span class="breadcrumb-current">${window.escapeHtml(item.label)}</span>`;
      } else {
        const onclick = item.action ? `onclick="${item.action}"` : '';
        return `${separator}<a class="breadcrumb-link" ${onclick}>${window.escapeHtml(item.label)}</a>`;
      }
    }).join('');
  };
  
  window.updateBreadcrumbs = function(pathArray) {
    const bar = document.getElementById('breadcrumb-bar');
    if (bar) {
      bar.innerHTML = window.renderBreadcrumbs(pathArray);
      bar.style.display = pathArray && pathArray.length > 0 ? 'block' : 'none';
    }
  };
  
  // ============================================================
  // Phone Number Formatting
  // ============================================================
  
  /**
   * Strip all non-digit characters from phone number for storage
   * @param {string} phone - Phone number in any format
   * @returns {string} Digits only (e.g., "0412345678")
   */
  window.stripPhoneForStorage = function(phone) {
    if (!phone) return '';
    return String(phone).replace(/\D/g, '');
  };
  
  /**
   * Format Australian mobile number for display (0412 345 678)
   * @param {string} phone - Phone number (with or without spaces)
   * @returns {string} Formatted phone or original if not standard format
   */
  window.formatPhoneForDisplay = function(phone) {
    if (!phone) return '';
    const digits = String(phone).replace(/\D/g, '');
    
    // Australian mobile: 10 digits starting with 04
    if (digits.length === 10 && digits.startsWith('04')) {
      return digits.slice(0, 4) + ' ' + digits.slice(4, 7) + ' ' + digits.slice(7);
    }
    
    // Australian landline: 10 digits starting with 0
    if (digits.length === 10 && digits.startsWith('0')) {
      return digits.slice(0, 2) + ' ' + digits.slice(2, 6) + ' ' + digits.slice(6);
    }
    
    // International +61: 11 digits starting with 61
    if (digits.length === 11 && digits.startsWith('61')) {
      return '+61 ' + digits.slice(2, 5) + ' ' + digits.slice(5, 8) + ' ' + digits.slice(8);
    }
    
    // Return original if format not recognized
    return phone;
  };
  
  /**
   * Get user preference for phone copy format (from cached user profile)
   * @returns {boolean} true = copy with spaces, false = copy without spaces
   */
  window.getPhoneCopyPreference = function() {
    if (window.currentUserProfile) {
      return window.currentUserProfile.phoneCopyWithSpaces === true;
    }
    return false;
  };
  
  /**
   * Set user preference for phone copy format (saves to Airtable)
   * @param {boolean} withSpaces - true = copy with spaces
   */
  window.setPhoneCopyPreference = async function(withSpaces) {
    try {
      const result = await window.apiBridge.call('updateUserPreference', 'phoneCopyWithSpaces', withSpaces);
      if (result && result.success && window.currentUserProfile) {
        window.currentUserProfile.phoneCopyWithSpaces = withSpaces;
      }
      return result;
    } catch (err) {
      console.error('Failed to save phone copy preference:', err);
      return { success: false, error: err.message };
    }
  };
  
  /**
   * Copy phone number to clipboard, respecting user preference
   * @param {string} phone - Phone number
   * @param {HTMLElement} element - Optional element for visual feedback
   */
  window.copyPhoneToClipboard = async function(phone, element) {
    if (!phone) return;
    
    const withSpaces = window.getPhoneCopyPreference();
    const textToCopy = withSpaces ? window.formatPhoneForDisplay(phone) : window.stripPhoneForStorage(phone);
    
    try {
      await navigator.clipboard.writeText(textToCopy);
      
      // Visual feedback
      if (element) {
        const original = element.innerHTML;
        element.innerHTML = '<span style="color: var(--color-cedar);">Copied!</span>';
        setTimeout(() => { element.innerHTML = original; }, 1500);
      }
    } catch (err) {
      // Fallback for older browsers
      const textarea = document.createElement('textarea');
      textarea.value = textToCopy;
      textarea.style.position = 'fixed';
      textarea.style.left = '-9999px';
      document.body.appendChild(textarea);
      textarea.select();
      document.execCommand('copy');
      document.body.removeChild(textarea);
      
      if (element) {
        const original = element.innerHTML;
        element.innerHTML = '<span style="color: var(--color-cedar);">Copied!</span>';
        setTimeout(() => { element.innerHTML = original; }, 1500);
      }
    }
  };
  
})();
