# Integrity

## Overview
Integrity is a Stellaris Finance Broking contact management system (CRM). Originally built for Google Apps Script, it has been adapted to run in Node.js/Express with Airtable as the backend database.

## Project Structure
```
├── server.js              # Express server with Google OAuth and API endpoints
├── services/
│   ├── airtable.js        # Airtable API integration layer
│   └── gmail.js           # Gmail API integration for sending emails
├── data.js                # Legacy mock data store (not used in production)
├── public/                # Static frontend files
│   ├── index.html         # Main HTML page
│   ├── styles.css         # Stylesheet (with custom font definitions)
│   ├── app.js             # Frontend JavaScript (CRM logic)
│   ├── api-bridge.js      # API bridge layer (converts google.script.run to fetch)
│   └── fonts/             # Custom fonts (Geist, Libre Baskerville)
├── Dockerfile             # Docker configuration for Fly.io deployment
├── fly.toml               # Fly.io deployment configuration
├── package.json           # Node.js dependencies
└── README.md              # Project description
```

## Running the Application
- The app runs on port 5000
- Start command: `npm start`
- Server binds to 0.0.0.0:5000 for Replit compatibility

## Environment Variables
### Required for Production
- `GOOGLE_CLIENT_ID` - Google OAuth client ID
- `GOOGLE_CLIENT_SECRET` - Google OAuth client secret
- `ALLOWED_GOOGLE_DOMAIN` - Google Workspace domain for access control
- `SESSION_SECRET` - Secret for cookie session encryption
- `AIRTABLE_API_KEY` - Airtable Personal Access Token
- `AIRTABLE_BASE_ID` - Airtable Base ID

### Development Mode
- Authentication is automatically disabled in Replit (detected via REPL_ID environment variable)
- To force OAuth in Replit: set `AUTH_DISABLED=false`
- On Fly.io (production): OAuth is automatically enabled since REPL_ID is not present

## Authentication
- Google OAuth 2.0 with domain restriction (only allows users from ALLOWED_GOOGLE_DOMAIN)
- Uses openid-client library for OAuth flow
- Session stored in encrypted cookies (cookie-session)
- Dev mode bypass available with AUTH_DISABLED=true

## API Endpoints
All API endpoints use POST method with JSON body `{ args: [...] }`:
- `POST /api/getRecentContacts` - Get list of recent contacts from Airtable
- `POST /api/searchContacts` - Search contacts by name
- `POST /api/getContactById` - Get contact details by ID
- `POST /api/updateRecord` - Update a contact field
- `POST /api/createRecord` - Create a new contact
- `POST /api/setSpouseStatus` - Link/unlink spouse relationships
- `POST /api/getLinkedOpportunities` - Get opportunities linked to a contact
- `POST /api/deleteContact` - Delete a contact (only if not connected to spouse or opportunities)
- `POST /api/getEffectiveUserEmail` - Get current user email
- `POST /api/parseTacoData` - Parse Taco key:value format and map to Airtable fields
- `GET /api/health` - Health check endpoint

## Airtable Integration
- Uses official Airtable SDK
- Tables: Contacts, Opportunities, Spouse History, Spouse History Log, Users, Settings
- All CRUD operations go through services/airtable.js
- Graceful error handling with console logging
- Settings table stores team-wide configuration (email links, signature templates)

### Spouse Connection Workflow
- Spouse connections/disconnections are managed via Airtable automations
- Web app creates a record in "Spouse History" table with Contact 1, Contact 2, and action
- Airtable automations then: update Spouse fields on both contacts, create Spouse History Log entries
- "Spouse History Text" is a lookup field on Contacts that reflects the history automatically

### Taco Import Integration
- When creating a new Opportunity, users can paste data from Taco (external system)
- Data is parsed from key:value format (one per line) and mapped to Airtable fields
- TACO_FIELD_MAP in server.js defines the mapping between Taco fields and Airtable fields
- Parse preview shows matched and unmapped fields before opportunity creation
- Taco fields stored in Opportunities (in order): New or Existing Client, Lead Source, Last Thing We Did, How can we help, CM notes, Broker, Broker Assistant, Client Manager, Converted to Appt, Appointment Time, Type of Appointment, Appt Phone Number, How Appt Booked, How Appt Booked Other, Need Evidence in Advance, Need Appt Reminder, Appt Conf Email Sent, Appt Conf Text Sent

## Technical Notes
- The `api-bridge.js` (formerly gas-shim.js) provides Google Apps Script compatibility, converting `google.script.run` calls to fetch API requests
- The bridge captures handlers per-call to prevent race conditions with concurrent API calls
- Cache-control headers prevent caching issues in Replit's iframe preview
- Quill.js WYSIWYG editor is used for rich text email composition

## Deployment
- **Development**: Replit with AUTH_DISABLED=true
- **Production**: Fly.io with full OAuth enabled
- Port 5000 used consistently across all environments

## Gmail API Integration
- Uses Replit Gmail connector for OAuth token management
- Emails sent via Gmail API appear in user's Sent folder
- HTML emails with rich formatting supported
- services/gmail.js handles token refresh and email composition

## Recent Changes
- January 2, 2026: Signature template updated - exact Mercury/Gmail-compatible HTML with logo, disclaimers, Calendly link
- January 2, 2026: Signature generator modal enlarged - 800px width, 300px min-height preview for full visibility
- January 2, 2026: Quill link dialog centered - positioned at center of editor instead of off-screen left
- January 2, 2026: Gmail API integration - emails sent directly from app appear in user's Gmail Sent folder
- January 2, 2026: Quill WYSIWYG editor - full rich text editing with bold, italic, underline, links, lists
- January 2, 2026: Signature generator - pulls Name and Title from Users table, generates HTML signature, copy buttons for Gmail and Mercury
- January 2, 2026: Settings table - team-wide email link storage in Airtable, syncs across all team members
- January 2, 2026: Renamed gas-shim.js to api-bridge.js for clarity
- January 2, 2026: "Fields from Taco Enquiry tab" heading replaces old "Taco Fields"
- January 2, 2026: Calendly checkbox text now proper case and updates reactively
- January 2, 2026: User signature storage - signatures stored in Airtable Users table ("Email Signature" field), syncs across all devices, edited via email settings modal
- January 2, 2026: WYSIWYG email editor - actual clickable HTML links in preview (Office, here, Fact Find, myGov, video, instructions), editable contenteditable preview, converts to plain text with URLs when opening Gmail
- January 2, 2026: Calendly integration - "Need Appt Reminder" checkbox shows "Not required as Calendly will do it automatically" when How Appt Booked = Calendly
- January 2, 2026: Email composer enhancements - Cedar header color, auto-populates To field with Primary Applicant + Applicants emails, clickable links for Office (Google Maps), Our Team, Fact Find, myGov, help video, and instructions; settings modal (gear icon) to edit all links, saved to localStorage
- January 2, 2026: Taco section visual improvements - Taco Fields heading moved inside blue box, rounded box with sky blue background, removed edit pencil icons, auto-hide past appointment details with expandable notice
- January 2, 2026: Email composer - built-in appointment confirmation email composer with dynamic templates, variable substitution, conditional logic (appointment type, new/repeat client, prep handler), live preview, and Gmail integration
- January 2, 2026: Delete Opportunity feature - checks for connections (Primary Applicant, Applicants, Guarantors, Loan Applications, Tasks) before allowing deletion
- January 2, 2026: Fixed Taco field capitalization - corrected field names to match Airtable exactly (Last thing we did, How appt booked, Taco Client Manager)
- January 2, 2026: Taco import feature - paste Taco data in New Opportunity composer, parses key:value format, shows preview of matched/unmapped fields, stores data in Taco-prefixed Airtable fields
- December 31, 2025: Expanded Opportunity fields - added 25+ new editable fields including loan consultant/processor IDs, amounts, values, dates, referral info, and related IDs
- December 31, 2025: Long-text field support - textarea editing for multi-line text fields in Opportunity panel
- December 31, 2025: Date field support - date picker with DD/MM/YYYY display format for Submitted Date
- December 31, 2025: Opportunity user tracking - Creating/Last Site User Name/Email fields now tracked for Opportunities (same as Contacts)
- December 31, 2025: Bi-directional link tracking - when Applicants are linked/unlinked from Opportunities, both the Opportunity and affected Contacts are marked as modified
- December 31, 2025: Opportunity audit display - slide-in panel now shows Created/Modified info at top (same format as Contact audit section)
- December 31, 2025: Development mode email - changed from dev@example.com to admin.ops@stellaris.loans for proper user tracking in Replit preview
- December 31, 2025: Fixed contact update bug - form data now properly serialized so updates work correctly (was creating new contacts instead)
- December 31, 2025: Spouse connection refactored - now creates Spouse History records instead of direct Contact updates, letting Airtable automations manage the workflow
- December 31, 2025: Spouse history display - sorts chronologically by timestamp (including time), displays as "DD/MM/YYYY: connected as spouse to Name"
- December 31, 2025: Lead Source fields - added Lead Source Major and Lead Source Minor as read-only fields in opportunity details
- December 31, 2025: Spacing improvements - reduced gaps between client name, audit info, and form fields for tighter layout
- December 29, 2025: Edit pencil relocated - now at top-right of editable fields section instead of next to contact name
- December 29, 2025: Cancel button for edit mode - appears when editing existing contact, reverts to last saved state
- December 29, 2025: Dark mode button contrast - Update Contact button now has proper text visibility in dark mode
- December 29, 2025: Subtitle moved inline - "prefers X · in database for Y" now appears next to contact name, not below
- December 29, 2025: Directory time column padding - added right padding and larger hover area for easier tooltip access
- December 29, 2025: Dynamic spouse checkbox label - shows "Also add X as Applicant?" unchecked, "Adding X as Applicant" checked
- December 29, 2025: Custom star toggle image - using user-provided PNG with dark/light split design
- December 29, 2025: Keyboard shortcuts help - "?" button in header opens modal listing all shortcuts, also press "?" key
- December 29, 2025: Directory time column - compact "2m", "3h", "5d" column on right side of contacts, hover for full details
- December 29, 2025: Opportunity status badges on left - [WON] Primary Applicant format with aligned fixed-width badge slot
- December 29, 2025: Star icon for dark mode toggle - gold star with half-dark overlay (not leaf icon)
- December 29, 2025: New opportunity modal - replaces native prompt with in-app modal, default name uses today's date (DD/MM/YYYY), option to add spouse as Applicant if contact has spouse linked
- December 29, 2025: Logo swap in dark mode - shows reversed logo when dark mode active
- December 29, 2025: Added UX enhancements - keyboard shortcuts (/ for search, N for new, E for edit, Esc to close), status color coding for opportunities (Won=green/cedar, Lost=gray, Open=sky), confetti celebration when opportunity marked Won, quick-add opportunity button, avatar initials with colored circles in directory, dark mode toggle with localStorage persistence
- December 29, 2025: Added createOpportunity API endpoint and Airtable function
- December 28, 2025: Full user tracking - looks up user name from Users table by email, populates Creating/Last Site User Name + Email fields
- December 28, 2025: Header redesign - full salt background, midnight text for INTEGRITY and email, larger logo (48px height)
- December 28, 2025: Lightened text throughout - reduced font weights, smaller headings, contact names now 13px regular weight
- December 28, 2025: Directory contact format changed - name on first line, "Prefers X · in database for Y" on second line in italics
- December 28, 2025: Reorganized contact form - Mobile moved to share row with Email
- December 28, 2025: Added custom fonts (Geist for body, Libre Baskerville for headings)
- December 28, 2025: Fixed sorting to use "Modified On" field (proper datetime, not text formula)
- December 28, 2025: Added Opportunity edit support (updateOpportunity for linked records)
- December 28, 2025: Auto-detect Replit environment (AUTH_DISABLED automatic in Replit, OAuth enforced on Fly.io)
- December 28, 2025: Added Airtable integration replacing mock data with real database
- December 28, 2025: Restored Google OAuth with dev-mode bypass option
- December 28, 2025: Updated Dockerfile and fly.toml to use port 5000
- December 28, 2025: Fixed race condition in gas-shim.js causing Directory to not load contacts
- December 28, 2025: Fixed port configuration for Replit webview compatibility
- December 27, 2025: Initial project setup with Express server and static frontend

## User Preferences
- Professional, production-ready setup with GitHub and Fly.io deployment
- Security is important - Google Workspace OAuth for team access control
- Future plans: Mercury CRM integration, WYSIWYG email editor, Slack integration, requirements tracking
