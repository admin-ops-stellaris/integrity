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