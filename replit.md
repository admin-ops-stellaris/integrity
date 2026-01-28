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
- **Email Integration**: Sends emails via Gmail API with HTML formatting and dynamic templates.
- **Settings Management**: Team-wide configurations accessible via a global settings modal.

**System Design Choices**:
- Designed for containerized deployment using Docker, optimized for platforms like Fly.io.
- Development environment in Replit allows authentication bypass for easier testing.
- Core business logic and data interactions are centralized through `services/airtable.js` and `services/gmail.js`.
- Frontend uses a modular architecture with IIFE patterns, `window.IntegrityState` for shared state, and an API Bridge. No build step required.
- **URL Routing**: Path-based nested routing with History API for deep linking.

## External Dependencies
- **Airtable**: Primary database for Contacts, Opportunities, Spouse History, Connections, Addresses, Users, Settings.
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
| `parseDateForEditor(isoString)` | **USE FOR FORMS** - Returns `{ dateDisplay, timeDisplay, isoDate, time24 }` |

**RULE:** When loading Airtable data into any form, ALWAYS use `parseDateForEditor()` to ensure Perth timezone consistency.

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

## Recent Session Changes (January 2026)

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
- **Module count**: 17 IIFE modules (no change)