require('dotenv').config();
const express = require('express');
const session = require('express-session');
const cookieParser = require('cookie-parser');
const jwt = require('jsonwebtoken');
const cors = require('cors');
const path = require('path');
const authRoutes = require('./routes/auth');
const calendarRoutes = require('./routes/calendar');

const app = express();
const PORT = process.env.PORT || 3000;
const JWT_SECRET = process.env.SESSION_SECRET || 'your-secret-key';
const isProduction = process.env.NODE_ENV === 'production' || process.env.VERCEL === '1';

// Middleware
app.use(cors({
    origin: true,
    credentials: true
}));
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(cookieParser());
app.use(express.static(path.join(__dirname, 'public')));

// Trust proxy for secure cookies behind Vercel's proxy
if (isProduction) {
    app.set('trust proxy', 1);
}

// Session configuration (for local development)
app.use(session({
    secret: JWT_SECRET,
    resave: false,
    saveUninitialized: false,
    cookie: {
        secure: isProduction,
        httpOnly: true,
        sameSite: 'lax',
        maxAge: 24 * 60 * 60 * 1000 // 24 hours
    }
}));

// JWT Cookie middleware - restore session from JWT cookie for serverless
app.use((req, res, next) => {
    const token = req.cookies?.auth_token;
    console.log('JWT middleware - token exists:', !!token, 'session accessToken:', !!req.session?.accessToken);
    if (token && !req.session?.accessToken) {
        try {
            const decoded = jwt.verify(token, JWT_SECRET);
            req.session.accessToken = decoded.accessToken;
            req.session.user = decoded.user;
            console.log('JWT middleware - restored session for user:', decoded.user?.email);
        } catch (err) {
            console.log('JWT middleware - invalid token:', err.message);
            // Invalid token, clear it
            res.clearCookie('auth_token');
        }
    }
    next();
});

// Helper to set auth cookie
app.setAuthCookie = (res, data) => {
    const token = jwt.sign(data, JWT_SECRET, { expiresIn: '24h' });
    console.log('Setting auth cookie, isProduction:', isProduction);
    res.cookie('auth_token', token, {
        httpOnly: true,
        secure: isProduction,
        sameSite: 'lax',
        maxAge: 24 * 60 * 60 * 1000 // 24 hours
    });
};

// Helper to clear auth cookie
app.clearAuthCookie = (res) => {
    res.clearCookie('auth_token');
};

// Routes
app.use('/auth', authRoutes);
app.use('/api/calendar', calendarRoutes);

// Serve the main page
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// API endpoint to check auth status
app.get('/api/status', (req, res) => {
    res.json({
        isAuthenticated: !!req.session?.accessToken,
        user: req.session?.user || null
    });
});

// Error handling middleware
app.use((err, req, res, next) => {
    console.error('Error:', err);
    res.status(500).json({ error: 'Internal server error', message: err.message });
});

app.listen(PORT, () => {
    console.log(`Server running on http://localhost:${PORT}`);
    console.log('Make sure to configure your .env file with Azure AD credentials');
});

// Export for Vercel serverless
module.exports = app;
