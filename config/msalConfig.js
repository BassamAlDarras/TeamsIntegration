const msal = require('@azure/msal-node');

const msalConfig = {
    auth: {
        clientId: process.env.CLIENT_ID,
        clientSecret: process.env.CLIENT_SECRET,
        authority: `https://login.microsoftonline.com/${process.env.TENANT_ID || 'common'}`
    },
    system: {
        loggerOptions: {
            loggerCallback(logLevel, message) {
                if (process.env.NODE_ENV === 'development') {
                    console.log(message);
                }
            },
            piiLoggingEnabled: false,
            logLevel: msal.LogLevel.Warning
        }
    }
};

const cca = new msal.ConfidentialClientApplication(msalConfig);

// Scopes required for Microsoft Graph API - Calendar and Teams access
const SCOPES = [
    'User.Read',
    'Calendars.ReadWrite',
    'OnlineMeetings.ReadWrite',
    'OnlineMeetingArtifact.Read.All'
];

const REDIRECT_URI = process.env.REDIRECT_URI || 'http://localhost:3000/auth/callback';

module.exports = {
    cca,
    SCOPES,
    REDIRECT_URI
};
