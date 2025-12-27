# Integrity

## Overview
Integrity is a Node.js web application with Express backend serving a static frontend.

## Project Structure
```
├── server.js          # Express server (main entry point)
├── public/            # Static frontend files
│   ├── index.html     # Main HTML page
│   ├── styles.css     # Stylesheet
│   └── app.js         # Frontend JavaScript
├── package.json       # Node.js dependencies
└── README.md          # Project description
```

## Running the Application
- The app runs on port 5000
- Start command: `npm start`
- Server binds to 0.0.0.0:5000 for Replit compatibility

## API Endpoints
- `GET /` - Serves the main HTML page
- `GET /api/health` - Health check endpoint returning JSON status

## Recent Changes
- December 27, 2025: Initial project setup with Express server and static frontend
