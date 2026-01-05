import { google } from 'googleapis';

// Replit connector settings cache (for development)
let connectionSettings = null;

// Check if running in Replit environment
const IS_REPLIT = !!process.env.REPL_ID;

async function getReplitAccessToken() {
  if (connectionSettings && connectionSettings.settings.expires_at && new Date(connectionSettings.settings.expires_at).getTime() > Date.now()) {
    return connectionSettings.settings.access_token;
  }
  
  const hostname = process.env.REPLIT_CONNECTORS_HOSTNAME;
  const xReplitToken = process.env.REPL_IDENTITY 
    ? 'repl ' + process.env.REPL_IDENTITY 
    : process.env.WEB_REPL_RENEWAL 
    ? 'depl ' + process.env.WEB_REPL_RENEWAL 
    : null;

  if (!xReplitToken) {
    throw new Error('X_REPLIT_TOKEN not found for repl/depl');
  }

  connectionSettings = await fetch(
    'https://' + hostname + '/api/v2/connection?include_secrets=true&connector_names=google-mail',
    {
      headers: {
        'Accept': 'application/json',
        'X_REPLIT_TOKEN': xReplitToken
      }
    }
  ).then(res => res.json()).then(data => data.items?.[0]);

  const accessToken = connectionSettings?.settings?.access_token || connectionSettings.settings?.oauth?.credentials?.access_token;

  if (!connectionSettings || !accessToken) {
    throw new Error('Gmail not connected');
  }
  return accessToken;
}

async function getProductionGmailClient(options) {
  const { tokens, clientId, clientSecret, onTokenRefresh } = options;
  
  if (!tokens?.access_token) {
    throw new Error('Gmail not authorized. Please log out and log back in to grant email permissions.');
  }
  
  const oauth2Client = new google.auth.OAuth2(clientId, clientSecret);
  
  oauth2Client.setCredentials({
    access_token: tokens.access_token,
    refresh_token: tokens.refresh_token,
  });
  
  // Check if token is expired or will expire soon (within 5 minutes)
  const now = Date.now();
  const expiresAt = tokens.expires_at || 0;
  
  if (expiresAt < now + 300000 && tokens.refresh_token) {
    console.log('Access token expired or expiring soon, refreshing...');
    try {
      const { credentials } = await oauth2Client.refreshAccessToken();
      
      // Update stored tokens
      const newTokens = {
        access_token: credentials.access_token,
        refresh_token: credentials.refresh_token || tokens.refresh_token,
        expires_at: credentials.expiry_date || Date.now() + 3600000,
      };
      
      oauth2Client.setCredentials(credentials);
      
      // Callback to update session
      if (onTokenRefresh) {
        onTokenRefresh(newTokens);
      }
      
      console.log('Token refreshed successfully');
    } catch (refreshError) {
      console.error('Failed to refresh token:', refreshError.message);
      throw new Error('Gmail authorization expired. Please log out and log back in.');
    }
  }
  
  return google.gmail({ version: 'v1', auth: oauth2Client });
}

async function getReplitGmailClient() {
  const accessToken = await getReplitAccessToken();

  const oauth2Client = new google.auth.OAuth2();
  oauth2Client.setCredentials({
    access_token: accessToken
  });

  return google.gmail({ version: 'v1', auth: oauth2Client });
}

export async function sendEmail(to, subject, htmlBody, options = {}) {
  try {
    let gmail;
    
    // Use production OAuth if tokens are provided, otherwise try Replit connector
    if (options.tokens) {
      gmail = await getProductionGmailClient(options);
    } else if (IS_REPLIT) {
      gmail = await getReplitGmailClient();
    } else {
      throw new Error('Gmail not configured. Please log out and log back in to authorize email sending.');
    }
    
    // Build the email message in MIME format
    const toAddresses = to.split(',').map(e => e.trim()).join(', ');
    
    const messageParts = [
      `To: ${toAddresses}`,
      `Subject: ${subject}`,
      'Content-Type: text/html; charset=utf-8',
      'MIME-Version: 1.0',
      '',
      htmlBody
    ];
    
    const message = messageParts.join('\r\n');
    
    // Encode to base64url format
    const encodedMessage = Buffer.from(message)
      .toString('base64')
      .replace(/\+/g, '-')
      .replace(/\//g, '_')
      .replace(/=+$/, '');
    
    // Send the email using Gmail API
    const response = await gmail.users.messages.send({
      userId: 'me',
      requestBody: {
        raw: encodedMessage
      }
    });
    
    console.log('Email sent successfully:', response.data.id);
    return { success: true, messageId: response.data.id };
  } catch (err) {
    console.error('Gmail send error:', err.message);
    return { success: false, error: err.message };
  }
}
