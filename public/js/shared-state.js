/**
 * Shared State Module
 * Central store for global state variables used across modules
 * 
 * IMPORTANT: This must be loaded FIRST, before any other modules.
 * All modules access state via window.IntegrityState
 */
const IntegrityState = (function() {
  return {
    // ============================================================
    // Timeouts and intervals
    // ============================================================
    searchTimeout: null,
    spouseSearchTimeout: null,
    linkedSearchTimeout: null,
    loadingTimer: null,
    pollInterval: null,
    pollAttempts: 0,
    screensaverTimer: null,
    quickViewHoverTimeout: null,
    previewUpdateTimeout: null,
    noteSaveTimeout: null,
    
    // ============================================================
    // Contact state
    // ============================================================
    contactStatusFilter: localStorage.getItem('contactStatusFilter') || 'Active',
    currentContactRecord: null,
    currentContactAddresses: [],
    currentEmployment: [],
    contactHistory: [],
    contactInlineEditor: null,
    
    // ============================================================
    // Search state
    // ============================================================
    searchHighlightIndex: -1,
    currentSearchRecords: [],
    
    // ============================================================
    // Opportunity state
    // ============================================================
    currentOppRecords: [],
    currentOppSortDirection: 'desc',
    currentPanelData: {},
    panelHistory: [],
    currentOppToDelete: null,
    
    // ============================================================
    // Edit state
    // ============================================================
    pendingLinkedEdits: {},
    pendingRemovals: {},
    editingAddressId: null,
    pendingDeceasedAction: null,
    
    // ============================================================
    // Quick View state
    // ============================================================
    quickViewContactId: null,
    isQuickViewHovered: false,
    
    // ============================================================
    // Connections state
    // ============================================================
    connectionRoleTypes: [],
    allConnectionsData: [],
    connectionsExpanded: false,
    selectedConnectionTarget: null,
    deactivatingConnectionId: null,
    
    // ============================================================
    // Appointments state
    // ============================================================
    currentAppointmentOpportunityId: null,
    editingAppointmentId: null,
    
    // ============================================================
    // Notes state
    // ============================================================
    activeNotePopover: null,
    
    // ============================================================
    // Email state
    // ============================================================
    currentEmailContext: null,
    userSignature: '',
    emailSettingsLoaded: false,
    emailQuill: null,
    currentUserProfile: null,
    generatedSignatureHtml: '',
    airtableTemplates: [],
    templatesLoaded: false,
    currentEditingTemplate: null,
    templateEditorQuill: null,
    templateSubjectQuill: null,
    activeTemplateEditor: 'body',
    templatePreviewContext: null,
    isHighlighting: false,
    
    // ============================================================
    // Evidence state
    // ============================================================
    evidenceOpportunityId: null,
    evidenceOpportunityName: '',
    evidenceOpportunityType: null,
    evidenceLender: null,
    evidenceItems: [],
    evidenceTemplates: [],
    evidenceCategories: [],
    evidenceEditingItemId: null,
    evidenceDeletingItemId: null,
    
    // ============================================================
    // Taco data parsing
    // ============================================================
    parsedTacoFields: {}
  };
})();

window.IntegrityState = IntegrityState;
