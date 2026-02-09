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
- Redesigned header for a clean interface, full-width 3-column contact layout, and two-line name display.

**Technical Implementations**:
- **Modular JavaScript Architecture**: Frontend refactored into lean initialization and independent IIFE modules for maintainability.
- **Authentication**: Google OAuth 2.0 secures access, restricted to a specified Google Workspace domain, with session management via encrypted cookies.
- **API Layer**: A unified API uses POST requests with JSON bodies for all CRUD and specific functionalities, with an `api-bridge.js` layer for compatibility.
- **Email Composition**: Rich text email composition via Quill.js WYSIWYG editor, integrated with Gmail API.
- **InlineEditingManager**: Reusable module for click-to-edit form fields with state tracking and navigation.
- **Accordion Sections Pattern**: Collapsible field groupings for organized content.
- **Header Search with Dropdown**: Integrated search functionality with keyboard shortcuts.
- **Modal Systems**: Standardized modal system for confirmations and data input.
- **Note Fields System**: Configuration-driven popover system for inline notes with auto-save.
- **Data Parsing**: Parser for Taco data, mapping external system data to Airtable fields for opportunity creation.
- **Global Smart Date/Time System**: Allows flexible date/time input (e.g., `210126`, `1:30pm`) with auto-formatting to `DD/MM/YYYY` and `h:mm AM/PM`, respectively.
- **Perth Standard Timezone Strategy**: All dates/times refer to Western Australia (UTC+8) using floating ISO strings to avoid timezone shifts.
- **Performance Optimization**: Lazy loading of contact data to ensure fast initial load times, fetching full details only when needed.

**Feature Specifications**:
- **Contact Management**: Comprehensive CRUD, spouse linking, and "Mark as Deceased" workflow.
- **Connections Management**: Tracks relationships with 12 role types, bidirectional querying, and note-taking.
- **Opportunity Management**: Full CRUD, user tracking, audit trails, and Taco data integration.
- **Appointment Management**: Dedicated table for appointments linked to opportunities, with full CRUD.
- **Evidence & Data Collection**: Full-screen modal for managing loan application evidence, with status tracking and email generation.
- **Address History**: Tracks residential and postal addresses with date range tracking.
- **Employment History**: Full CRUD for employment records with Primary/Secondary/Previous status, conflict resolution for Primary changes, conditional field visibility based on employment type, and JSON blob storage for Income entries and Employer Address.
- **Email Integration**: Sends emails via Gmail API with HTML formatting and dynamic templates.
- **Settings Management**: Team-wide configurations accessible via a global settings modal.

**System Design Choices**:
- Designed for containerized deployment using Docker, optimized for platforms like Fly.io.
- Development environment in Replit allows authentication bypass for easier testing.
- Core business logic and data interactions are centralized through `services/airtable.js` and `services/gmail.js`.
- Frontend uses a modular architecture with IIFE patterns, `window.IntegrityState` for shared state, and an API Bridge. No build step required.
- **URL Routing**: Path-based nested routing with History API for deep linking.

## External Dependencies
- **Airtable**: Primary database for Contacts, Opportunities, Spouse History, Connections, Addresses, Employment, Users, Settings.
- **Google OAuth 2.0**: For user authentication and authorization.
- **Gmail API**: For sending emails.
- **Quill.js**: WYSIWYG editor.
- **openid-client**: Library for Google OAuth 2.0 flow.
- **cookie-session**: For encrypting and managing user sessions.
- **Fly.io**: Deployment platform.
- **Replit Gmail connector**: Manages OAuth tokens for Gmail integration.
- **Taco (external system)**: For importing and parsing data for opportunity creation.

## Perth Standard Timezone Strategy

All dates/times in Integrity refer to Western Australia (UTC+8). This prevents timezone shifts when users in different locations view or edit data.

**Philosophy:**
- **Floating ISO strings** (e.g., `2026-01-21T14:30:00` without Z or offset) preserve user intent
- Only strings ending with `Z` are treated as UTC and converted to Perth time
- `Date()` constructor is avoided for floating strings to prevent browser timezone shifts

**Key Functions in `shared-utils.js`:**
| Function | Purpose |
|----------|---------|
| `constructDateForSave(dateStr, timeStr)` | Creates floating ISO: `YYYY-MM-DDTHH:mm:00` |
| `parseFloatingDate(isoString)` | Parses floating ISO to Date for comparisons |
| `formatDateTimeForDisplay(isoString, options)` | Display formatting; Z→Perth, floating→direct |
| `parseDateForEditor(isoString)` | **DATETIME FORMS** - Returns `{ dateDisplay, timeDisplay, isoDate, time24 }` |
| `formatDateDisplay(isoDate)` | **DATE-ONLY FORMS** - Converts `YYYY-MM-DD` → `DD/MM/YYYY` |
| `parseFlexibleDate(value)` | Smart date parsing - Returns `{ iso, display }` for saving |

**RULES:**
- **Datetime fields** (e.g., Appointment): Use `parseDateForEditor()` to load, `constructDateForSave()` to save
- **Date-only fields** (e.g., DOB, Employment dates): Use `formatDateDisplay()` to load, `parseFlexibleDate().iso` to save

## Global Layout System

Standardized flexbox utility classes for consistent spacing across the app (defined at top of `styles.css`).

**Stack Classes:**
| Class | Effect |
|-------|--------|
| `.layout-stack-y` | Vertical flex column |
| `.layout-stack-x` | Horizontal flex row, vertically centered |

**Gap Modifiers:**
| Class | Size | Use Case |
|-------|------|----------|
| `.gap-xs` | 0.5rem (8px) | Tight groupings |
| `.gap-sm` | 1rem (16px) | Related items |
| `.gap-md` | 1.5rem (24px) | Section elements |
| `.gap-lg` | 2.5rem (40px) | Major sections |
| `.gap-xl` | 4rem (64px) | Page-level separation |

**Usage:** Combine stack + gap classes, e.g., `class="layout-stack-y gap-md"`

## Phone Number Formatting System

Australian phone numbers are formatted for readability while storing raw digits in Airtable.

**Philosophy:**
- **Display with spaces** for readability: `0412 345 678`
- **Store without spaces** in Airtable: `0412345678`
- **Copy to clipboard** respects user preference (with or without spaces)
- **Input flexibility**: Users can type/paste with or without spaces

**Key Functions in `shared-utils.js`:**
| Function | Purpose |
|----------|---------|
| `stripPhoneForStorage(phone)` | Removes all non-digits for Airtable storage |
| `formatPhoneForDisplay(phone)` | Formats AU mobiles as `0412 345 678`, landlines as `08 9123 4567` |
| `getPhoneCopyPreference()` | Returns user's Airtable preference for copy format |
| `setPhoneCopyPreference(bool)` | Saves user preference to Airtable (true = with spaces) |
| `copyPhoneToClipboard(phone, el)` | Copies to clipboard using preference, shows "Copied!" feedback |

**User Preference:** Accessible via Settings cog → "Phone Number Copy Format" checkbox. Stored in Airtable Users table.

## User Preferences Storage

All user preferences are stored in the **Users table in Airtable** (not localStorage) to ensure they sync across devices and persist through cache clears.

**Current User Preference Fields:**
| Field Name | Type | Purpose |
|------------|------|---------|
| `Phone Copy With Spaces` | Checkbox | When checked, phone numbers copy with spaces |

**Adding New User Preferences:**
1. Add field to Users table in Airtable
2. Add mapping in `updateUserPreference()` fieldMap in `services/airtable.js`
3. Include field in `getUserSignature()` and `getUserProfileByEmail()` return objects
4. Access via `window.currentUserProfile.[preferenceName]` on frontend

## Recent Session Changes (January 2026)

- **Phone Number Formatting**: Display with spaces, store without, click-to-copy with user preference
- **Dossier Header Architecture**: Refactored contact header to left block (breadcrumb→name→badge) + right block (status badges + created/modified metadata)
- **3-Column CSS Grid Layout**: Profile columns use `grid-template-columns: 320px minmax(400px, 1fr) 400px` with Mobile Pulse responsive layout at 1024px (Activity > Opportunities > Relationships > Facts)
- **Activity Stream Placeholder**: Right column (400px) reserved for future Tasks/Slack/AI integration
- **Removed #contactMetaBar**: Functionality integrated into new dossier header
- **Global Layout System**: Added flexbox utility classes for consistent spacing
- **Breadcrumb Navigation**: Hierarchy-based eyebrow navigation (Contacts > Name > Opportunity), now inside dossier left block
- **Centralized timezone handling**: Created `parseDateForEditor()` for consistent UTC→Perth conversion
- **Marketing Badge Normalization**: Uses "Unsubscribed from Marketing" field consistently across app.js and contacts.js
- **Modal Search Dropdown Fix**: Changed to `position: static` with `max-height: 150px` and bottom margin to prevent covering Cancel buttons
- **Connections Empty State**: Added "Connections" label visible when no connections exist
- **Home Screen Cleanup**: Hide dossier-header on home screen (no "Contact" title when no contact selected)
- **Duplicate Warning System**: Manual warnings displayed in left column with EDIT/HIDE links, Add/Delete options in ACTIONS menu
- **Duplicate Detection**: Auto-checks for duplicates on new contact creation (mobile, email, name matching), shows modal with potential matches and "Create Anyway" option
- **User Preferences in Airtable**: Migrated user preferences (e.g., phone copy format) from localStorage to Airtable Users table for cross-device sync
- **Module count**: 21 JS modules in public/js/ (addresses, appointments, connections, contacts, contacts-search, core, email, employment, evidence, inline-editing, marketing, modal-utils, notes, opportunities, quick-view, router, settings, shared-state, shared-utils, spouse, ui-utils)
- **Employment History Module**: Added full Employment module with Primary conflict resolution modal, conditional section visibility (PAYG/Self Employed/Unemployed/Retired), JSON blob storage for Income list and Employer Address, sorted list display (Primary first, then Secondary by start date, then Previous by end date)
- **Marketing Module (Sidecar Architecture)**: Export section with stats (Total/Unsubscribed/Marketable/Ready to Send), smart CSV export with deduplication by email (shared emails combine names as "John & Jane")
- **Marketing Import Results**: Multi-base Airtable support (MARKETING_BASE_ID env var). Campaign selector dropdown, CSV file picker with robust parser (handles quoted newlines). Backend validates status values (opened/clicked/bounced/unsubscribed/sent/delivered/complained), updates Main CRM (Bounced/Unsubscribed statuses), creates logs in Marketing Base Logs table. Confirmation dialog before import with detailed results summary.
- **Marketing Timeline**: Right column displays per-contact marketing history from Sidecar Base. Shows event type with color-coded icons (green=Opened, blue=Clicked, red=Bounced/Unsubscribed), campaign name, and timestamp. Loads automatically when a contact is selected. Module: `marketing-timeline.js`.
- **Campaign Manager Dashboard (Sprint 4)**: Mega Modal (90vw x 90vh) with 3 tabs: Campaigns (dashboard table with Delivery/Open/Engagement/Unsub rates, clickable rows for detail view), Import Tool (existing import flow), Export (existing CSV export). Detail view shows recipient list with filter toggles (All/Opened/Clicked/Bounced/Unsubscribed), clickable contact names, status badges, and mailto follow-up button for unsubscribed recipients. Backend: `getCampaignStats()` aggregates all logs per campaign, `getCampaignLogs(campaignId)` returns individual recipient logs with Contact Name lookup.