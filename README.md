# Teams Calendar Integration App

A web application that allows users to link their Microsoft Teams account and create appointments with automatic Teams meeting links.

## Features

- 🔐 Microsoft OAuth 2.0 authentication
- 📅 Create calendar appointments with Teams meeting links
- 👥 Add attendees to meetings
- 📋 View upcoming appointments
- 🗑️ Delete appointments
- 🎥 Automatic Teams meeting link generation

## Prerequisites

- Node.js (v16 or higher)
- Microsoft Azure AD App Registration

## Azure AD App Setup

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** > **App registrations** > **New registration**
3. Configure your app:
   - **Name**: Teams Calendar Integration (or your preferred name)
   - **Supported account types**: Accounts in any organizational directory and personal Microsoft accounts
   - **Redirect URI**: Web - `http://localhost:3000/auth/callback`

4. After registration, note down:
   - **Application (client) ID** → This is your `CLIENT_ID`
   - **Directory (tenant) ID** → This is your `TENANT_ID` (use `common` for multi-tenant)

5. Create a client secret:
   - Go to **Certificates & secrets** > **New client secret**
   - Copy the secret value → This is your `CLIENT_SECRET`

6. Configure API permissions:
   - Go to **API permissions** > **Add a permission** > **Microsoft Graph**
   - Add the following **Delegated permissions**:
     - `User.Read`
     - `Calendars.ReadWrite`
     - `OnlineMeetings.ReadWrite`
   - Click **Grant admin consent** (if you have admin privileges)

## Installation

1. Clone or download this project

2. Install dependencies:
   ```bash
   npm install
   ```

3. Create a `.env` file from the example:
   ```bash
   cp .env.example .env
   ```

4. Edit `.env` and add your Azure AD credentials:
   ```
   CLIENT_ID=your_client_id_here
   CLIENT_SECRET=your_client_secret_here
   TENANT_ID=common
   SESSION_SECRET=generate_a_random_string_here
   PORT=3000
   REDIRECT_URI=http://localhost:3000/auth/callback
   ```

5. Start the server:
   ```bash
   # Development mode (with auto-reload)
   npm run dev
   
   # Production mode
   npm start
   ```

6. Open your browser and go to `http://localhost:3000`

## Usage

1. Click **"Link Microsoft Teams Account"** to authenticate
2. Sign in with your Microsoft account
3. Grant the requested permissions
4. You'll be redirected back to the app
5. Click **"New Appointment"** to create a meeting
6. Fill in the meeting details:
   - Meeting title
   - Start and end date/time
   - Location (optional)
   - Attendees (comma-separated emails)
   - Description
   - Enable/disable Teams meeting link
7. Click **"Create Appointment"**
8. View your appointments in the list below

## API Endpoints

### Authentication
- `GET /auth/login` - Initiate Microsoft OAuth login
- `GET /auth/callback` - OAuth callback handler
- `GET /auth/logout` - Logout and clear session
- `GET /auth/user` - Get current user info

### Calendar
- `GET /api/calendar/events` - Get user's calendar events
- `POST /api/calendar/events` - Create a new calendar event
- `PUT /api/calendar/events/:eventId` - Update an event
- `DELETE /api/calendar/events/:eventId` - Delete an event
- `POST /api/calendar/meeting` - Create a Teams meeting (without calendar event)

### Status
- `GET /api/status` - Check authentication status

## Project Structure

```
teamsintegration/
├── config/
│   └── msalConfig.js     # MSAL configuration
├── routes/
│   ├── auth.js           # Authentication routes
│   └── calendar.js       # Calendar/Events API routes
├── public/
│   ├── index.html        # Main HTML page
│   └── js/
│       └── app.js        # Frontend JavaScript
├── server.js             # Express server
├── package.json          # Dependencies
├── .env.example          # Environment variables template
└── README.md             # This file
```

## Troubleshooting

### "AADSTS50011: The reply URL specified does not match"
- Ensure the redirect URI in Azure AD matches exactly: `http://localhost:3000/auth/callback`

### "AADSTS65001: The user or administrator has not consented"
- Make sure you've added the correct API permissions
- Try signing out and signing back in to re-consent

### "Failed to create appointment"
- Verify your app has `Calendars.ReadWrite` permission
- Check that admin consent has been granted (if required by your organization)

## Security Notes

- Never commit your `.env` file to version control
- Use a strong, random `SESSION_SECRET`
- In production, use HTTPS and set `secure: true` for cookies
- Rotate your `CLIENT_SECRET` periodically

## License

MIT
