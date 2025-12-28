# Integrity

## Overview
Integrity is a Stellaris Finance Broking contact management system (CRM). Originally built for Google Apps Script, it has been adapted to run in Node.js/Express with an in-memory data store.

## Project Structure
```
├── server.js          # Express server (main entry point)
├── data.js            # In-memory contact data store with mock contacts
├── public/            # Static frontend files
│   ├── index.html     # Main HTML page
│   ├── styles.css     # Stylesheet
│   ├── app.js         # Frontend JavaScript (CRM logic)
│   └── gas-shim.js    # Google Apps Script compatibility layer
├── Dockerfile         # Docker configuration for Fly.io deployment
├── fly.toml           # Fly.io deployment configuration
├── package.json       # Node.js dependencies
└── README.md          # Project description
```

## Running the Application
- The app runs on port 5000
- Start command: `npm start`
- Server binds to 0.0.0.0:5000 for Replit compatibility

## API Endpoints
All API endpoints use POST method with JSON body `{ args: [...] }`:
- `POST /api/getRecentContacts` - Get list of recent contacts
- `POST /api/searchContacts` - Search contacts by name
- `POST /api/getContactById` - Get contact details by ID
- `POST /api/updateRecord` - Update a contact field
- `POST /api/createRecord` - Create a new contact
- `POST /api/setSpouseStatus` - Link/unlink spouse relationships
- `POST /api/getLinkedOpportunities` - Get opportunities linked to a contact
- `POST /api/getEffectiveUserEmail` - Get current user email

## Technical Notes
- The `gas-shim.js` provides Google Apps Script compatibility, converting `google.script.run` calls to fetch API requests
- The shim captures handlers per-call to prevent race conditions with concurrent API calls
- Cache-busting query strings are used on script tags to ensure fresh JavaScript loads

## Recent Changes
- December 28, 2025: Updated Dockerfile and fly.toml to use port 5000 (was 3000) to prevent Replit auto-adding conflicting port mappings
- December 28, 2025: Fixed race condition in gas-shim.js causing Directory to not load contacts
- December 28, 2025: Added cache-busting to script tags for reliable updates
- December 28, 2025: Fixed port configuration for Replit webview compatibility
- December 27, 2025: Initial project setup with Express server and static frontend
