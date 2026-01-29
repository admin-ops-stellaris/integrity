/**
 * Router Module
 * Handles URL-based navigation for contacts and opportunities
 * 
 * URL Patterns:
 *   /                          - Home (contact list)
 *   /contact/:contactId        - View contact
 *   /contact/:contactId/opportunity/:oppId - View contact with opportunity panel open
 */
const IntegrityRouter = (function() {
  
  let isNavigating = false;
  
  function parseRoute(pathname) {
    pathname = pathname || window.location.pathname;
    
    const contactMatch = pathname.match(/^\/contact\/([a-zA-Z0-9]+)(?:\/opportunity\/([a-zA-Z0-9]+))?$/);
    
    if (contactMatch) {
      return {
        type: 'contact',
        contactId: contactMatch[1],
        opportunityId: contactMatch[2] || null
      };
    }
    
    return { type: 'home' };
  }
  
  function buildUrl(contactId, opportunityId) {
    if (!contactId) return '/';
    if (opportunityId) {
      return '/contact/' + contactId + '/opportunity/' + opportunityId;
    }
    return '/contact/' + contactId;
  }
  
  function navigateTo(contactId, opportunityId, options) {
    options = options || {};
    const replaceState = options.replace || false;
    
    const url = buildUrl(contactId, opportunityId);
    const currentUrl = window.location.pathname;
    
    if (url === currentUrl) return;
    
    if (replaceState) {
      history.replaceState({ contactId: contactId, opportunityId: opportunityId }, '', url);
    } else {
      history.pushState({ contactId: contactId, opportunityId: opportunityId }, '', url);
    }
  }
  
  function navigateToHome(options) {
    navigateTo(null, null, options);
  }
  
  async function handleRoute(route) {
    if (isNavigating) return;
    isNavigating = true;
    
    route = route || parseRoute();
    
    try {
      if (route.type === 'home') {
        window.goHome && window.goHome();
        return;
      }
      
      if (route.type === 'contact' && route.contactId) {
        await loadContactFromUrl(route.contactId, route.opportunityId);
      }
    } catch (err) {
      console.error('Router: Error handling route', err);
      window.showAlert && window.showAlert('Failed to load: ' + err.message);
      navigateToHome({ replace: true });
    } finally {
      isNavigating = false;
    }
  }
  
  function loadContactFromUrl(contactId, opportunityId) {
    return new Promise(function(resolve, reject) {
      window.toggleProfileView && window.toggleProfileView(true);
      window.hideSearchDropdown && window.hideSearchDropdown();
      
      var formTitle = document.getElementById('formTitle');
      var formSubtitle = document.getElementById('formSubtitle');
      var emptyState = document.getElementById('emptyState');
      var profileContent = document.getElementById('profileContent');
      
      if (formTitle) formTitle.innerText = 'Loading...';
      if (formSubtitle) formSubtitle.innerHTML = '';
      if (emptyState) emptyState.style.display = 'none';
      if (profileContent) profileContent.style.display = 'flex';
      
      google.script.run
        .withSuccessHandler(function(record) {
          if (record && record.fields) {
            window.selectContactFromRouter(record, function() {
              if (opportunityId && window.openOpportunityPanel) {
                window.openOpportunityPanel(opportunityId);
              }
              resolve();
            });
          } else {
            reject(new Error('Contact not found'));
          }
        })
        .withFailureHandler(function(err) {
          reject(err);
        })
        .getContactById(contactId);
    });
  }
  
  function init() {
    window.addEventListener('popstate', function(e) {
      var state = e.state;
      if (state) {
        handleRoute({
          type: state.contactId ? 'contact' : 'home',
          contactId: state.contactId,
          opportunityId: state.opportunityId
        });
      } else {
        handleRoute(parseRoute());
      }
    });
    
    var route = parseRoute();
    handleRoute(route);
  }
  
  return {
    parseRoute: parseRoute,
    buildUrl: buildUrl,
    navigateTo: navigateTo,
    navigateToHome: navigateToHome,
    handleRoute: handleRoute,
    init: init
  };
  
})();

window.IntegrityRouter = IntegrityRouter;
