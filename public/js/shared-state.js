/**
 * Shared State Module
 * Central store for global state variables used across modules
 */
const IntegrityState = (function() {
  return {
    // Timeouts and intervals
    searchTimeout: null,
    spouseSearchTimeout: null,
    linkedSearchTimeout: null,
    loadingTimer: null,
    pollInterval: null,
    pollAttempts: 0,
    
    // Contact state
    contactStatusFilter: localStorage.getItem('contactStatusFilter') || 'Active',
    currentContactRecord: null,
    currentContactAddresses: [],
    contactHistory: [],
    
    // Opportunity state
    currentOppRecords: [],
    currentOppSortDirection: 'desc',
    currentPanelData: {},
    panelHistory: [],
    
    // Search state
    searchHighlightIndex: -1,
    currentSearchRecords: [],
    
    // Edit state
    pendingLinkedEdits: {},
    pendingRemovals: {},
    editingAddressId: null,
    
    // Evidence state
    evidenceOpportunityId: null,
    evidenceOpportunityName: '',
    evidenceOpportunityType: null,
    evidenceLender: null,
    evidenceItems: [],
    evidenceTemplates: [],
    evidenceCategories: [],
    evidenceEditingItemId: null,
    evidenceDeletingItemId: null,
    
    // Email templates
    emailTemplates: [],
    
    // Connections
    selectedConnectionTarget: null,
    deactivatingConnectionId: null,
    
    // Quill editors
    quillEditor: null,
    templateQuillEditor: null,
    evTplQuillEditor: null
  };
})();

window.IntegrityState = IntegrityState;
