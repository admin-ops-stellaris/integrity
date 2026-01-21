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