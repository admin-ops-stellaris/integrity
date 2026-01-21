/**
 * Core Module
 * Dark mode, screensaver, scroll header
 * These are standalone UI features extracted from app.js
 */
(function() {
  'use strict';
  
  const state = window.IntegrityState;
  
  // ============================================================
  // Dark Mode
  // ============================================================
  
  function initDarkMode() {
    const savedTheme = localStorage.getItem('integrity-theme');
    if (savedTheme === 'dark') document.body.classList.add('dark-mode');
  }
  
  function toggleDarkMode() {
    document.body.classList.toggle('dark-mode');
    const isDark = document.body.classList.contains('dark-mode');
    localStorage.setItem('integrity-theme', isDark ? 'dark' : 'light');
  }
  
  // ============================================================
  // Screensaver / Idle Timer
  // ============================================================
  
  const SCREENSAVER_DELAY = 120000; // 2 minutes
  
  function initScreensaver() {
    function resetScreensaverTimer() {
      if (document.body.classList.contains('screensaver-active')) {
        document.body.classList.remove('screensaver-active');
      }
      clearTimeout(state.screensaverTimer);
      state.screensaverTimer = setTimeout(() => {
        document.body.classList.add('screensaver-active');
      }, SCREENSAVER_DELAY);
    }
    
    ['mousemove', 'mousedown', 'keydown', 'touchstart', 'scroll'].forEach(event => {
      document.addEventListener(event, resetScreensaverTimer, { passive: true });
    });
    
    resetScreensaverTimer();
  }
  
  // ============================================================
  // Scroll-Hide Header (Mobile/Tablet)
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
    
    window.addEventListener('resize', function() {
      if (window.innerWidth > 1024) {
        header.classList.remove('header-hidden');
      }
    });
  }
  
  // ============================================================
  // Global Smart Date System - Event Delegation
  // ============================================================
  
  function initSmartDateListener() {
    document.addEventListener('change', function(e) {
      const target = e.target;
      if (!target || target.tagName !== 'INPUT') return;
      
      // Only target text inputs (type="date" uses browser picker, already returns YYYY-MM-DD)
      if (target.type !== 'text') return;
      
      // Check if this is a smart date field
      const isSmartDate = target.classList.contains('smart-date') ||
                          (target.id && /[Dd]ate/.test(target.id));
      
      if (!isSmartDate) return;
      
      const value = target.value;
      if (!value) {
        delete target.dataset.isoDate;
        return;
      }
      
      const parsed = window.parseFlexibleDate(value);
      if (parsed) {
        target.value = parsed.display;
        target.dataset.isoDate = parsed.iso;
      } else {
        delete target.dataset.isoDate;
      }
    }, true);
  }
  
  // ============================================================
  // Expose functions globally
  // ============================================================
  
  window.initDarkMode = initDarkMode;
  window.toggleDarkMode = toggleDarkMode;
  window.initScreensaver = initScreensaver;
  window.initScrollHeader = initScrollHeader;
  window.initSmartDateListener = initSmartDateListener;
  
})();
