# Integrity

## Overview
Integrity is a Stellaris Finance Broking contact management system (CRM). Originally built for Google Apps Script, it has been adapted to run in Node.js/Express with Airtable as the backend database.

## Project Structure
```
├── server.js              # Express server with Google OAuth and API endpoints
├── services/
│   └── airtable.js        # Airtable API integration layer
├── data.js                # Legacy mock data store (not used in production)
├── public/                # Static frontend files
│   ├── index.html         # Main HTML page
│   ├── styles.css         # Stylesheet (with custom font definitions)
│   ├── app.js             # Frontend JavaScript (CRM logic)
│   ├── gas-shim.js        # Google Apps Script compatibility layer
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
- `POST /api/getEffectiveUserEmail` - Get current user email
- `GET /api/health` - Health check endpoint

## Airtable Integration
- Uses official Airtable SDK
- Tables: Contacts, Opportunities
- All CRUD operations go through services/airtable.js
- Graceful error handling with console logging

## Technical Notes
- The `gas-shim.js` provides Google Apps Script compatibility, converting `google.script.run` calls to fetch API requests
- The shim captures handlers per-call to prevent race conditions with concurrent API calls
- Cache-control headers prevent caching issues in Replit's iframe preview

## Deployment
- **Development**: Replit with AUTH_DISABLED=true
- **Production**: Fly.io with full OAuth enabled
- Port 5000 used consistently across all environments

## Recent Changes
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
