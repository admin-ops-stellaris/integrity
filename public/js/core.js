/**
 * Core Module
 * Initialization, dark mode, keyboard shortcuts, screensaver, scroll header
 */
(function() {
  'use strict';
  
  const state = window.IntegrityState;
  
  // ============================================================
  // Initialization
  // ============================================================
  
  window.onload = function() {
    loadContacts();
    checkUserIdentity();
    initKeyboardShortcuts();
    initDarkMode();
    initScreensaver();
    initInlineEditing();
    initAllNoteFields();
    initScrollHeader();
  };
  
  // Initialize contact status filter on DOM ready
  document.addEventListener('DOMContentLoaded', function() {
    const saved = localStorage.getItem('contactStatusFilter') || 'Active';
    state.contactStatusFilter = saved;
    updateStatusToggleUI(saved);
  });
  
  // ============================================================
  // Contact Status Filter
  // ============================================================
  
  window.setContactStatusFilter = function(status) {
    state.contactStatusFilter = status;
    localStorage.setItem('contactStatusFilter', status);
    updateStatusToggleUI(status);
    const query = document.getElementById('searchInput')?.value?.trim();
    if (query && query.length > 0) {
      handleSearch({ target: { value: query } });
    } else {
      loadContacts();
    }
  };
  
  function updateStatusToggleUI(status) {
    document.querySelectorAll('.status-toggle-btn').forEach(btn => {
      btn.classList.toggle('active', btn.dataset.status === status);
    });
  }
  
  // ============================================================
  // Scroll Header (Mobile/Tablet)
  // ============================================================
  
  function initScrollHeader() {
    let lastScrollY = 0;
    let ticking = false;
    const header = document.querySelector('.app-header');
    
    if (!header) return;
    
    function handleScroll(scrollTop) {
      if (!ticking) {
        window.requestAnimationFrame(function() {
          const isMobileOrTablet = window.innerWidth <= 1024;
          
          if (isMobileOrTablet) {
            if (scrollTop > lastScrollY && scrollTop > 50) {
              header.classList.add('header-hidden');
            } else {
              header.classList.remove('header-hidden');
            }
          } else {
            header.classList.remove('header-hidden');
          }
          
          lastScrollY = scrollTop;
          ticking = false;
        });
        ticking = true;
      }
    }
    
    const container = document.querySelector('.container');
    const columns = document.querySelectorAll('.column');
    
    if (container) {
      container.addEventListener('scroll', function() {
        handleScroll(this.scrollTop);
      });
    }
    
    columns.forEach(function(col) {
      col.addEventListener('scroll', function() {
        handleScroll(this.scrollTop);
      });
    });
  }
  
  // ============================================================
  // Keyboard Shortcuts
  // ============================================================
  
  function initKeyboardShortcuts() {
    document.addEventListener('keydown', function(e) {
      if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA' || e.target.isContentEditable) {
        if (e.key === 'Escape') {
          e.target.blur();
          hideSearchDropdown();
        }
        return;
      }
      
      if (e.key === '/' || e.key === 'k' && (e.metaKey || e.ctrlKey)) {
        e.preventDefault();
        const searchInput = document.getElementById('searchInput');
        if (searchInput) {
          searchInput.focus();
          showSearchDropdown();
        }
      } else if (e.key === 'Escape') {
        const oppPanel = document.getElementById('oppDetailPanel');
        if (oppPanel && oppPanel.classList.contains('open')) {
          closeOppPanel();
        } else {
          hideSearchDropdown();
        }
      } else if (e.key === 'n' || e.key === 'N') {
        e.preventDefault();
        resetForm();
      } else if (e.key === 'e' || e.key === 'E') {
        if (state.currentContactRecord && document.getElementById('editBtn').style.visibility !== 'hidden') {
          e.preventDefault();
          enableEditMode();
        }
      }
    });
  }
  
  window.showShortcutsHelp = function() {
    openModal('shortcutsModal');
  };
  
  window.closeShortcutsModal = function() {
    closeModal('shortcutsModal');
  };
  
  // ============================================================
  // Dark Mode
  // ============================================================
  
  function initDarkMode() {
    if (localStorage.getItem('darkMode') === 'true') {
      document.body.classList.add('dark-mode');
    }
  }
  
  window.toggleDarkMode = function() {
    document.body.classList.toggle('dark-mode');
    localStorage.setItem('darkMode', document.body.classList.contains('dark-mode'));
  };
  
  // ============================================================
  // Screensaver
  // ============================================================
  
  function initScreensaver() {
    let screensaverTimer;
    const IDLE_TIME = 5 * 60 * 1000;
    
    function resetTimer() {
      clearTimeout(screensaverTimer);
      hideScreensaver();
      screensaverTimer = setTimeout(showScreensaver, IDLE_TIME);
    }
    
    function showScreensaver() {
      document.getElementById('screensaverOverlay')?.classList.add('active');
    }
    
    function hideScreensaver() {
      document.getElementById('screensaverOverlay')?.classList.remove('active');
    }
    
    ['mousemove', 'mousedown', 'keydown', 'touchstart', 'scroll'].forEach(event => {
      document.addEventListener(event, resetTimer, { passive: true });
    });
    
    resetTimer();
  }
  
  // ============================================================
  // User Identity
  // ============================================================
  
  function checkUserIdentity() {
    google.script.run.withSuccessHandler(function(user) {
      document.getElementById('userEmail').innerText = user?.email || 'Unknown';
    }).getCurrentUser();
  }
  
  window.updateHeaderTitle = function(isEditing) {
    const titleEl = document.querySelector('.header-title');
    if (titleEl) {
      titleEl.textContent = isEditing ? 'INTEGRITY*' : 'INTEGRITY';
    }
  };
  
  // ============================================================
  // Profile View Toggle
  // ============================================================
  
  window.toggleProfileView = function(show) {
    if (show) {
      document.getElementById('emptyState').style.display = 'none';
      document.getElementById('profileContent').style.display = 'block';
    } else {
      document.getElementById('emptyState').style.display = 'flex';
      document.getElementById('profileContent').style.display = 'none';
    }
  };
  
  // ============================================================
  // Go Home
  // ============================================================
  
  window.goHome = function() {
    state.currentContactRecord = null;
    state.currentOppRecords = [];
    state.currentContactAddresses = [];
    toggleProfileView(false);
    closeOppPanel();
    
    document.getElementById('searchInput').value = '';
    loadContacts();
    
    const modals = document.querySelectorAll('.modal-overlay, .modal');
    modals.forEach(m => {
      m.style.display = 'none';
      m.classList.remove('showing', 'visible');
    });
  };
  
  // ============================================================
  // Won Celebration
  // ============================================================
  
  window.triggerWonCelebration = function() {
    const container = document.getElementById('celebrationContainer');
    if (!container) return;
    container.innerHTML = '';
    for (let i = 0; i < 50; i++) {
      const confetti = document.createElement('div');
      confetti.className = 'confetti';
      confetti.style.left = Math.random() * 100 + '%';
      confetti.style.animationDelay = Math.random() * 0.5 + 's';
      confetti.style.backgroundColor = ['#BB9934', '#7B8B64', '#19414C', '#FFD700', '#FFA500'][Math.floor(Math.random() * 5)];
      container.appendChild(confetti);
    }
    setTimeout(() => container.innerHTML = '', 3000);
  };
  
})();
