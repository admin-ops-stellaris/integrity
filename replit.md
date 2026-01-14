# Integrity

## Overview
Integrity is a Customer Relationship Management (CRM) system designed for Stellaris Finance Broking. It facilitates contact management and integrates with Airtable as its primary backend database. The system, originally a Google Apps Script project, has been re-engineered to run on Node.js/Express, providing a robust and scalable solution for managing client interactions and opportunities. The project aims to enhance operational efficiency and streamline communication workflows for finance brokers.

## User Preferences
- Professional, production-ready setup with GitHub and Fly.io deployment
- Security is important - Google Workspace OAuth for team access control
- Future plans: Mercury CRM integration, WYSIWYG email editor, Slack integration, requirements tracking

## Brand Colors
- Trail: #2C2622
- Midnight: #19414C
- Star: #BB9934
- Cedar: #7B8B64
- Salt: #F2F0E9
- Sky: #D0DFE6

## System Architecture
Integrity is built on a Node.js/Express backend, serving a static frontend.
- **UI/UX Decisions**: The application features a modern UI with custom fonts (Geist, Libre Baskerville), dark mode toggle with persistence, and intuitive UX enhancements like keyboard shortcuts and status color-coding. Opportunity status badges, avatar initials with colored circles, and a redesigned header contribute to a clean and efficient interface.
- **Technical Implementations**:
    - **Authentication**: Google OAuth 2.0 is used for secure access, restricted to a specified Google Workspace domain. Session management is handled via encrypted cookies.
    - **API Layer**: A unified API uses POST requests with JSON bodies for all CRUD operations and specific functionalities like contact search, spouse management, and opportunity handling. An `api-bridge.js` layer ensures compatibility by converting `google.script.run` calls to standard fetch API requests.
    - **Email Composition**: Rich text email composition is supported via Quill.js WYSIWYG editor.
    - **InlineEditingManager**: Reusable IIFE module for click-to-edit form fields. Features include: per-field state tracking, composite key session tracking (field+sessionId) for async race condition handling, Tab key navigation between fields, select element support via parent click handlers, and bulk edit mode for new record creation. Located in `public/app.js`.
    - **Note Fields System**: Configuration-driven note popover system for attaching inline notes to form fields. Add new note fields by adding entries to `NOTE_FIELDS` array in `public/app.js`. Each entry specifies: `fieldId` (hidden field ID), `airtableField` (Airtable column name), `inputId` (visible input to attach icon to). Current note fields: email1Comment, email2Comment, email3Comment (for email addresses), and genderOther (for Gender - Other specification). System handles: icon rendering, popover UI, auto-save with debounce, keyboard shortcuts (Tab/Enter/Escape), text selection protection, and icon state (grey on hover, gold when populated). Works with both input and select elements.
    - **Data Parsing**: The system includes a parser for Taco data, mapping key-value pairs from an external system to Airtable fields during opportunity creation.
- **Feature Specifications**:
    - **Contact Management**: CRUD operations for contacts, including linking/unlinking spouse relationships and retrieving linked opportunities.
    - **Connections Management**: Relationship tracking between contacts with 12 role types (Parent/Child, Sibling, Friend, Household Rep/Member, Employer/Employee, Referral-based). Single non-reciprocal records with bidirectional querying. Connections are deactivated rather than deleted to preserve history. UI includes color-coded role badges, clickable contact names for navigation, note icons per connection (popover with auto-save), and add/remove functionality. Two-step modal flow: contact search with recently modified list, then relationship type selection. Connection notes are stored in Airtable "Note" field.
    - **Opportunity Management**: Creation, updating, and deletion of opportunities with extensive fields, user tracking, and audit trails. Integration with Taco data for streamlined opportunity creation.
    - **Appointment Management**: Dedicated Appointments table linked to Opportunities. Full CRUD operations with fields: Appointment Time, Type (Office/Phone/Video), How Booked (Calendly/Email/Phone/Podium/Other), Phone Number, Video Meet URL, checkboxes for evidence/reminder needs, appointment status, and notes. Appointments section displays in the Opportunity panel with Add/Edit/Delete functionality.
    - **Evidence & Data Collection**: Full-screen modal system for managing loan application evidence requirements. Features include: 4 Airtable tables (Evidence Categories, Evidence Templates, Lender Evidence Rules, Evidence Items), category-based organization (Identification, Income, Assets, Liabilities, etc.), status tracking (Outstanding/Received/N/A), progress bar with percentage completion, email generation for Initial/Subsequent requests with automatic "Requested On/By" tracking, custom item creation, clipboard copy, and lender-specific rules support. Evidence button appears in Opportunity panel next to Add Appointment.
    - **Email Integration**: Sending emails via Gmail API with support for HTML formatting, dynamic templates, and signature generation. Emails appear in the user's Sent folder.
    - **Settings Management**: Team-wide configurations, such as email links and signature templates, are stored in Airtable and accessible via a global settings modal.
- **System Design Choices**:
    - The application is designed for containerized deployment using Docker, optimized for platforms like Fly.io.
    - Development environment in Replit allows for authentication bypass for easier testing, while production environments enforce full OAuth.
    - All core business logic and data interactions are channeled through `services/airtable.js` and `services/gmail.js` for modularity.

## External Dependencies
- **Airtable**: Primary database for storing Contacts, Opportunities, Spouse History, Spouse History Log, Connections, Users, and Settings. Utilizes the official Airtable SDK.
- **Google OAuth 2.0**: For user authentication and authorization, restricting access to a specific Google Workspace domain.
- **Gmail API**: Integrated for sending emails directly from the application, including rich text and dynamic content.
- **Quill.js**: WYSIWYG editor used for rich text email composition.
- **openid-client**: Library for handling Google OAuth 2.0 flow.
- **cookie-session**: Used for encrypting and managing user sessions.
- **Fly.io**: Deployment platform for the production environment.
- **Replit Gmail connector**: Used for managing OAuth tokens for Gmail integration in the Replit environment.
- **Taco (external system)**: Data from Taco can be imported and parsed for opportunity creation.