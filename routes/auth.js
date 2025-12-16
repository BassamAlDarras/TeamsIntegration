const express = require('express');
const router = express.Router();
const { cca, SCOPES, REDIRECT_URI } = require('../config/msalConfig');

// Initiate Microsoft OAuth login
router.get('/login', async (req, res) => {
    try {
        const authCodeUrlParameters = {
            scopes: SCOPES,
            redirectUri: REDIRECT_URI,
            prompt: 'select_account'
        };

        const authUrl = await cca.getAuthCodeUrl(authCodeUrlParameters);
        res.redirect(authUrl);
    } catch (error) {
        console.error('Error generating auth URL:', error);
        res.status(500).json({ error: 'Failed to initiate login' });
    }
});

// OAuth callback handler
router.get('/callback', async (req, res) => {
    const { code, error, error_description } = req.query;

    if (error) {
        console.error('Auth error:', error, error_description);
        return res.redirect(`/?error=${encodeURIComponent(error_description || error)}`);
    }

    if (!code) {
        return res.redirect('/?error=No authorization code received');
    }

    try {
        const tokenRequest = {
            code: code,
            scopes: SCOPES,
            redirectUri: REDIRECT_URI
        };

        const response = await cca.acquireTokenByCode(tokenRequest);
        
        // Store tokens in session
        req.session.accessToken = response.accessToken;
        req.session.idToken = response.idToken;
        req.session.account = response.account;
        req.session.user = {
            name: response.account.name,
            email: response.account.username,
            id: response.account.homeAccountId
        };

        console.log('User authenticated:', req.session.user.email);
        res.redirect('/?success=true');
    } catch (error) {
        console.error('Error acquiring token:', error);
        res.redirect(`/?error=${encodeURIComponent('Authentication failed')}`);
    }
});

// Logout endpoint
router.get('/logout', (req, res) => {
    const postLogoutRedirectUri = `http://localhost:${process.env.PORT || 3000}`;
    
    req.session.destroy((err) => {
        if (err) {
            console.error('Error destroying session:', err);
        }
        // Redirect to Microsoft logout
        const logoutUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/logout?post_logout_redirect_uri=${encodeURIComponent(postLogoutRedirectUri)}`;
        res.redirect(logoutUrl);
    });
});

// Get current user info
router.get('/user', (req, res) => {
    if (!req.session.accessToken) {
        return res.status(401).json({ error: 'Not authenticated' });
    }
    res.json({ user: req.session.user });
});

module.exports = router;
