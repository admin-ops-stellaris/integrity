# Integrity

## Overview
Integrity is a Customer Relationship Management (CRM) system for Stellaris Finance Broking, re-engineered from Google Apps Script to Node.js/Express. It centralizes contact management and integrates with Airtable as its primary backend. The system aims to significantly improve operational efficiency and streamline communication workflows for finance brokers, enhancing client interaction and opportunity management.

## User Preferences
- Professional, production-ready setup with GitHub and Fly.io deployment
- Security is important - Google Workspace OAuth for team access control
- Future plans: Mercury CRM integration, WYSIWYG email editor, Slack integration, requirements tracking

## System Architecture
Integrity employs a Node.js/Express backend serving a static frontend.

**UI/UX Decisions**:
- Modern UI with custom fonts (Geist, Libre Baskerville), dark mode toggle with persistence.
- Intuitive UX with keyboard shortcuts, status color-coding, opportunity status badges, and avatar initials.
- Redesigned header for a clean interface.
- **Contact Layout**: Full-width 3-column layout for detailed contact information.
- **Name Display**: Two-line format for clear presentation of contact names and metadata.

**Technical Implementations**:
- **Modular JavaScript Architecture**: Frontend refactored into lean initialization (~1,100 lines) and 16 independent IIFE modules, enhancing maintainability and scalability.
- **Authentication**: Google OAuth 2.0 secures access, restricted to a specified Google Workspace domain, with session management via encrypted cookies.
- **API Layer**: A unified API uses POST requests with JSON bodies for all CRUD and specific functionalities. An `api-bridge.js` layer ensures compatibility with standard fetch API requests.
- **Email Composition**: Rich text email composition is supported via Quill.js WYSIWYG editor, integrated with Gmail API.
- **InlineEditingManager**: Reusable module for click-to-edit form fields, supporting per-field state tracking, keyboard navigation, and bulk edit mode.
- **Accordion Sections Pattern**: Collapsible field groupings for organized content display.
- **Header Search with Dropdown**: Integrated search functionality with keyboard shortcuts.
- **Modal Systems**: Standardized modal system for confirmations and data input.
- **Note Fields System**: Configuration-driven popover system for attaching inline notes to form fields with auto-save and keyboard handling.
- **Data Parsing**: Includes a parser for Taco data, mapping external system data to Airtable fields for opportunity creation.

**Feature Specifications**:
- **Contact Management**: Comprehensive CRUD operations, including spouse linking/unlinking, and "Mark as Deceased" workflow.
- **Connections Management**: Tracks relationships between contacts with 12 role types, bidirectional querying, and note-taking capabilities.
- **Opportunity Management**: Full CRUD for opportunities, with user tracking, audit trails, and Taco data integration for streamlined creation.
- **Appointment Management**: Dedicated table for appointments linked to opportunities, supporting full CRUD operations and various appointment types.
- **Evidence & Data Collection**: Full-screen modal system for managing loan application evidence, including category-based organization, status tracking, progress bar, and email generation.
- **Address History**: Tracks residential and postal addresses with format-aware fields, date range tracking, and status management.
- **Email Integration**: Sends emails via Gmail API with HTML formatting, dynamic templates, and signature generation.
- **Settings Management**: Team-wide configurations stored in Airtable and accessible via a global settings modal.

**System Design Choices**:
- Designed for containerized deployment using Docker, optimized for platforms like Fly.io.
- Development environment in Replit allows authentication bypass for easier testing.
- All core business logic and data interactions are centralized through `services/airtable.js` and `services/gmail.js`.

## External Dependencies
- **Airtable**: Primary database for Contacts, Opportunities, Spouse History, Connections, Addresses, Users, Settings.
- **Google OAuth 2.0**: For user authentication and authorization.
- **Gmail API**: For sending emails directly from the application.
- **Quill.js**: WYSIWYG editor for rich text email composition.
- **openid-client**: Library for Google OAuth 2.0 flow.
- **cookie-session**: For encrypting and managing user sessions.
- **Fly.io**: Deployment platform.
- **Replit Gmail connector**: Manages OAuth tokens for Gmail integration in Replit.
- **Taco (external system)**: For importing and parsing data for opportunity creation.

## Refactoring Milestone - January 2026

### Shadow Strategy Refactor: Complete Success

The frontend codebase has been successfully refactored from a monolithic architecture to a clean, modular structure using the "Shadow Strategy" - extracting code incrementally while maintaining full functionality.

**Before:** `app.js` was 7,725 lines of tightly coupled code
**After:** `app.js` is ~1,100 lines (lean initialization), with 16 independent IIFE modules

### Module Structure (Source of Truth)

```
public/js/
├── shared-state.js      # Global state (currentContactRecord, panelHistory, timeouts)
├── shared-utils.js      # Pure utilities (escapeHtml, formatDate*, parseDateInput)
├── modal-utils.js       # Modal management (openModal, closeModal, showAlert, showConfirmModal)
├── contacts-search.js   # Contact search, display, keyboard nav, avatar helpers
├── core.js              # Dark mode, screensaver/idle timer, scroll-hide header
├── inline-editing.js    # InlineEditingManager, field mapping, edit mode
├── spouse.js            # Spouse section, history, connect/disconnect modal
├── connections.js       # 12 role types, bidirectional display, connection notes
├── notes.js             # Note popover system, NOTE_FIELDS config, auto-save
├── addresses.js         # Address history (residential + postal), CRUD
├── appointments.js      # Appointment CRUD, inline editing, datetime formatting
├── opportunities.js     # Quick Add, Taco parsing, panel navigation, URL updates
├── settings.js          # Team settings, signature generation, EMAIL_LINKS
├── quick-view.js        # Contact hover cards, positioning
├── email.js             # Quill WYSIWYG, template CRUD, conditional parsing
├── evidence.js          # Evidence modal, progress tracking, email generation
└── router.js            # URL routing, deep linking, browser history management
```

### Load Order (Critical)

```
shared-state.js → shared-utils.js → modal-utils.js → contacts-search.js → 
core.js → inline-editing.js → spouse.js → connections.js → notes.js → 
addresses.js → appointments.js → opportunities.js → settings.js → 
quick-view.js → email.js → evidence.js → router.js → app.js
```

### URL Routing (Deep Linking) - January 2026

**URL Patterns:**
- `/` - Home (contact list)
- `/contact/:contactId` - View specific contact
- `/contact/:contactId/opportunity/:oppId` - View contact with opportunity panel open

**Implementation:**
1. **Path-based Nested Routing**: Clean URLs using History API (pushState/popState)
2. **Server Catch-All**: `app.get("*")` serves index.html for all unknown routes
3. **Auth Trap**: Unauthenticated users redirected to login; original URL saved in session and restored after OAuth callback
4. **Daisy Chain Rehydration**: Deep links load contact first, then open opportunity panel
5. **Unidirectional Flow**: `selectContact()` and `loadPanelRecord()` update URL; `popstate` triggers reverse navigation

**Key Functions in router.js:**
- `parseRoute(pathname)` - Extract contactId/opportunityId from URL
- `navigateTo(contactId, oppId)` - Update URL via pushState
- `handleRoute(route)` - Load content for a given route
- `init()` - Set up popstate listener and handle initial deep link

### Global Smart Date System - January 2026

**Goal:** Users can type dates loosely (e.g., `210126`, `21/1/26`, `21.01.26`) and have them auto-format to `DD/MM/YYYY`.

**Implementation:**
1. **parseFlexibleDate()** in shared-utils.js handles all formats:
   - No separators: `DDMMYY`, `DDMMYYYY`
   - With separators: `DD/MM/YY`, `DD.MM.YYYY`, `D-M-YY`, etc.
   - Returns `{ iso: 'YYYY-MM-DD', display: 'DD/MM/YYYY' }` or null

2. **Event Delegation** in core.js:
   - Global `change` listener on document
   - Targets: `.smart-date` class OR text inputs with 'Date' in ID
   - Auto-formats value and stores ISO in `data-iso-date` attribute

3. **Integration** via enhanced `parseDateInput(value, inputEl)`:
   - Checks `dataset.isoDate` first (from smart listener)
   - Falls back to `parseFlexibleDate()` then legacy pattern
   - Backward compatible with existing code

**Usage:**
- Add `.smart-date` class to any text input for date formatting
- Or use ID containing 'Date' (e.g., `addressFromDate`)
- `type="date"` inputs use browser picker (already returns ISO)

### Global Smart Time System - January 2026

**Goal:** Users can type times loosely (e.g., `1300`, `130`, `1:30pm`) and have them auto-format to `h:mm AM/PM`.

**Implementation:**
1. **parseFlexibleTime()** in shared-utils.js handles all formats:
   - Military time: `1300`, `930`, `130` (3-4 digits)
   - Colon format: `13:00`, `1:30`
   - Hour only: `9`, `13`
   - AM/PM suffix: `1pm`, `1:30pm`, `9a`
   - Returns `{ value24: 'HH:MM', display: 'h:mm AM/PM' }` or null

2. **Event Delegation** in core.js:
   - Global `change` listener on document
   - Targets: `.smart-time` class only
   - Auto-formats value and stores 24h time in `data-time24` attribute

3. **Appointment Modal Refactor**:
   - Split single datetime field into separate Date + Time fields
   - Date field uses `.smart-date`, Time field uses `.smart-time`
   - On save: combines `data-iso-date` + `data-time24` into ISO string

**Usage:**
- Add `.smart-time` class to text inputs for time formatting
- Access 24h value via `element.dataset.time24`

### Performance Optimization: Lazy Loading Contacts

**Problem:** `getRecentContacts` was returning full deep records with heavy arrays (Opportunities, Connections, Address History, Spouse History, SEARCH_INDEX).

**Solution:** Implemented lazy loading pattern:
1. `getRecentContacts` and `searchContacts` now fetch only list fields:
   - `Calculated Name`, `FirstName`, `MiddleName`, `LastName`
   - `EmailAddress1`, `Mobile`, `Status`, `Deceased`, `Modified`
2. Returned records are marked with `_isPartial: true`
3. `selectContact()` detects partial records and fetches full data on click
4. Shows "Loading..." state while fetching full record

**Result:** Fast initial load, Smart Search list remains fully functional, heavy details only fetched when needed.

### Architecture Principles

- **IIFE Pattern**: Each module is an Immediately Invoked Function Expression exposing functions to `window`
- **State via IntegrityState**: All shared mutable state accessed through `window.IntegrityState`
- **API Bridge**: `api-bridge.js` converts `google.script.run` calls to REST API calls
- **No Build Step**: Plain ES5-compatible JavaScript, loaded via script tags in order

This modular architecture is now the **Source of Truth** for the Integrity CRM frontend.