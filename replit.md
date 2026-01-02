# Integrity

## Overview
Integrity is a Customer Relationship Management (CRM) system designed for Stellaris Finance Broking. It facilitates contact management and integrates with Airtable as its primary backend database. The system, originally a Google Apps Script project, has been re-engineered to run on Node.js/Express, providing a robust and scalable solution for managing client interactions and opportunities. The project aims to enhance operational efficiency and streamline communication workflows for finance brokers.

## User Preferences
- Professional, production-ready setup with GitHub and Fly.io deployment
- Security is important - Google Workspace OAuth for team access control
- Future plans: Mercury CRM integration, WYSIWYG email editor, Slack integration, requirements tracking

## System Architecture
Integrity is built on a Node.js/Express backend, serving a static frontend.
- **UI/UX Decisions**: The application features a modern UI with custom fonts (Geist, Libre Baskerville), dark mode toggle with persistence, and intuitive UX enhancements like keyboard shortcuts and status color-coding. Opportunity status badges, avatar initials with colored circles, and a redesigned header contribute to a clean and efficient interface.
- **Technical Implementations**:
    - **Authentication**: Google OAuth 2.0 is used for secure access, restricted to a specified Google Workspace domain. Session management is handled via encrypted cookies.
    - **API Layer**: A unified API uses POST requests with JSON bodies for all CRUD operations and specific functionalities like contact search, spouse management, and opportunity handling. An `api-bridge.js` layer ensures compatibility by converting `google.script.run` calls to standard fetch API requests.
    - **Email Composition**: Rich text email composition is supported via Quill.js WYSIWYG editor.
    - **Data Parsing**: The system includes a parser for Taco data, mapping key-value pairs from an external system to Airtable fields during opportunity creation.
- **Feature Specifications**:
    - **Contact Management**: CRUD operations for contacts, including linking/unlinking spouse relationships and retrieving linked opportunities.
    - **Opportunity Management**: Creation, updating, and deletion of opportunities with extensive fields, user tracking, and audit trails. Integration with Taco data for streamlined opportunity creation.
    - **Email Integration**: Sending emails via Gmail API with support for HTML formatting, dynamic templates, and signature generation. Emails appear in the user's Sent folder.
    - **Settings Management**: Team-wide configurations, such as email links and signature templates, are stored in Airtable and accessible via a global settings modal.
- **System Design Choices**:
    - The application is designed for containerized deployment using Docker, optimized for platforms like Fly.io.
    - Development environment in Replit allows for authentication bypass for easier testing, while production environments enforce full OAuth.
    - All core business logic and data interactions are channeled through `services/airtable.js` and `services/gmail.js` for modularity.

## External Dependencies
- **Airtable**: Primary database for storing Contacts, Opportunities, Spouse History, Spouse History Log, Users, and Settings. Utilizes the official Airtable SDK.
- **Google OAuth 2.0**: For user authentication and authorization, restricting access to a specific Google Workspace domain.
- **Gmail API**: Integrated for sending emails directly from the application, including rich text and dynamic content.
- **Quill.js**: WYSIWYG editor used for rich text email composition.
- **openid-client**: Library for handling Google OAuth 2.0 flow.
- **cookie-session**: Used for encrypting and managing user sessions.
- **Fly.io**: Deployment platform for the production environment.
- **Replit Gmail connector**: Used for managing OAuth tokens for Gmail integration in the Replit environment.
- **Taco (external system)**: Data from Taco can be imported and parsed for opportunity creation.