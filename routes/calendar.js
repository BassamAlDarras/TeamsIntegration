const express = require('express');
const router = express.Router();
const axios = require('axios');

const GRAPH_API_BASE = 'https://graph.microsoft.com/v1.0';

// Middleware to check authentication
const requireAuth = (req, res, next) => {
    if (!req.session.accessToken) {
        return res.status(401).json({ error: 'Not authenticated. Please link your Teams account first.' });
    }
    next();
};

// Helper function for Graph API calls
async function callGraphApi(accessToken, endpoint, method = 'GET', data = null) {
    const config = {
        method,
        url: `${GRAPH_API_BASE}${endpoint}`,
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
        }
    };

    if (data) {
        config.data = data;
    }

    const response = await axios(config);
    return response.data;
}

// Get user's calendar events
router.get('/events', requireAuth, async (req, res) => {
    try {
        const { startDateTime, endDateTime } = req.query;
        
        let endpoint = '/me/calendar/events?$orderby=start/dateTime&$top=50';
        
        if (startDateTime && endDateTime) {
            endpoint = `/me/calendarView?startDateTime=${startDateTime}&endDateTime=${endDateTime}&$orderby=start/dateTime`;
        }

        const events = await callGraphApi(req.session.accessToken, endpoint);
        
        const formattedEvents = events.value.map(event => ({
            id: event.id,
            subject: event.subject,
            start: event.start,
            end: event.end,
            location: event.location?.displayName || '',
            isOnlineMeeting: event.isOnlineMeeting,
            onlineMeetingUrl: event.onlineMeeting?.joinUrl || null,
            attendees: event.attendees?.map(a => ({
                email: a.emailAddress.address,
                name: a.emailAddress.name,
                status: a.status?.response
            })) || [],
            body: event.body?.content || '',
            organizer: event.organizer?.emailAddress?.address || '',
            isCancelled: event.isCancelled || false,
            importance: event.importance || 'normal',
            sensitivity: event.sensitivity || 'normal',
            showAs: event.showAs || 'busy',
            categories: event.categories || [],
            webLink: event.webLink || ''
        }));

        res.json({ events: formattedEvents });
    } catch (error) {
        console.error('Error fetching events:', error.response?.data || error.message);
        res.status(500).json({ error: 'Failed to fetch calendar events' });
    }
});

// Sync calendar - Get all events from user's Microsoft calendar
router.get('/sync', requireAuth, async (req, res) => {
    try {
        const { days = 30 } = req.query;
        
        const startDateTime = new Date().toISOString();
        const endDateTime = new Date(Date.now() + parseInt(days) * 24 * 60 * 60 * 1000).toISOString();
        
        // Get calendar view for the specified period
        const endpoint = `/me/calendarView?startDateTime=${startDateTime}&endDateTime=${endDateTime}&$orderby=start/dateTime&$top=100&$select=id,subject,start,end,location,isOnlineMeeting,onlineMeeting,attendees,body,organizer,isCancelled,importance,showAs,categories,webLink,recurrence`;
        
        const events = await callGraphApi(req.session.accessToken, endpoint);
        
        const syncedEvents = events.value.map(event => ({
            id: event.id,
            subject: event.subject || '(No title)',
            start: event.start,
            end: event.end,
            location: event.location?.displayName || '',
            isOnlineMeeting: event.isOnlineMeeting || false,
            onlineMeetingUrl: event.onlineMeeting?.joinUrl || null,
            attendees: event.attendees?.map(a => ({
                email: a.emailAddress?.address || '',
                name: a.emailAddress?.name || '',
                status: a.status?.response || 'none'
            })) || [],
            bodyPreview: event.body?.content ? event.body.content.substring(0, 200).replace(/<[^>]*>/g, '') : '',
            organizer: event.organizer?.emailAddress?.address || '',
            isCancelled: event.isCancelled || false,
            importance: event.importance || 'normal',
            showAs: event.showAs || 'busy',
            categories: event.categories || [],
            webLink: event.webLink || '',
            isRecurring: !!event.recurrence,
            syncedAt: new Date().toISOString()
        }));

        res.json({ 
            success: true,
            syncedAt: new Date().toISOString(),
            period: {
                start: startDateTime,
                end: endDateTime,
                days: parseInt(days)
            },
            totalEvents: syncedEvents.length,
            events: syncedEvents 
        });
    } catch (error) {
        console.error('Error syncing calendar:', error.response?.data || error.message);
        res.status(500).json({ 
            error: 'Failed to sync calendar',
            details: error.response?.data?.error?.message || error.message
        });
    }
});

// Get all user's calendars (for multi-calendar sync)
router.get('/calendars', requireAuth, async (req, res) => {
    try {
        const calendars = await callGraphApi(req.session.accessToken, '/me/calendars');
        
        const formattedCalendars = calendars.value.map(cal => ({
            id: cal.id,
            name: cal.name,
            color: cal.color,
            isDefaultCalendar: cal.isDefaultCalendar,
            canEdit: cal.canEdit,
            owner: cal.owner?.address || ''
        }));

        res.json({ calendars: formattedCalendars });
    } catch (error) {
        console.error('Error fetching calendars:', error.response?.data || error.message);
        res.status(500).json({ error: 'Failed to fetch calendars' });
    }
});

// Sync events from a specific calendar
router.get('/calendars/:calendarId/sync', requireAuth, async (req, res) => {
    try {
        const { calendarId } = req.params;
        const { days = 30 } = req.query;
        
        const startDateTime = new Date().toISOString();
        const endDateTime = new Date(Date.now() + parseInt(days) * 24 * 60 * 60 * 1000).toISOString();
        
        const endpoint = `/me/calendars/${calendarId}/calendarView?startDateTime=${startDateTime}&endDateTime=${endDateTime}&$orderby=start/dateTime&$top=100`;
        
        const events = await callGraphApi(req.session.accessToken, endpoint);
        
        const syncedEvents = events.value.map(event => ({
            id: event.id,
            calendarId: calendarId,
            subject: event.subject || '(No title)',
            start: event.start,
            end: event.end,
            location: event.location?.displayName || '',
            isOnlineMeeting: event.isOnlineMeeting || false,
            onlineMeetingUrl: event.onlineMeeting?.joinUrl || null,
            organizer: event.organizer?.emailAddress?.address || '',
            syncedAt: new Date().toISOString()
        }));

        res.json({ 
            success: true,
            calendarId,
            totalEvents: syncedEvents.length,
            events: syncedEvents 
        });
    } catch (error) {
        console.error('Error syncing calendar:', error.response?.data || error.message);
        res.status(500).json({ error: 'Failed to sync calendar' });
    }
});

// Create a new calendar event with Teams meeting
router.post('/events', requireAuth, async (req, res) => {
    try {
        const { 
            subject, 
            startDateTime, 
            endDateTime, 
            attendees, 
            body, 
            location,
            isOnlineMeeting = true,
            timeZone = 'UTC'
        } = req.body;

        // Validate required fields
        if (!subject || !startDateTime || !endDateTime) {
            return res.status(400).json({ 
                error: 'Missing required fields: subject, startDateTime, and endDateTime are required' 
            });
        }

        let teamsJoinUrl = null;

        // STEP 1: If online meeting requested, create Teams meeting FIRST
        if (isOnlineMeeting) {
            console.log('Creating Teams meeting first...');
            try {
                const meetingData = {
                    subject,
                    startDateTime: new Date(startDateTime).toISOString(),
                    endDateTime: new Date(endDateTime).toISOString()
                };
                
                const meeting = await callGraphApi(req.session.accessToken, '/me/onlineMeetings', 'POST', meetingData);
                teamsJoinUrl = meeting.joinUrl;
                console.log('Teams meeting created successfully:', teamsJoinUrl);
            } catch (meetingError) {
                console.error('Failed to create Teams meeting:', meetingError.response?.data || meetingError.message);
                // Continue without Teams link - will create normal calendar event
            }
        }

        // STEP 2: Create calendar event
        const eventData = {
            subject,
            start: {
                dateTime: startDateTime,
                timeZone
            },
            end: {
                dateTime: endDateTime,
                timeZone
            },
            body: {
                contentType: 'HTML',
                content: teamsJoinUrl 
                    ? `${body || ''}<br><br><p><strong>Microsoft Teams Meeting</strong></p><p><a href="${teamsJoinUrl}">Click here to join the meeting</a></p>`
                    : (body || '')
            },
            location: {
                displayName: location || (teamsJoinUrl ? 'Microsoft Teams Meeting' : '')
            }
        };

        // If we have a Teams URL, set the online meeting properties
        if (teamsJoinUrl) {
            eventData.isOnlineMeeting = true;
            eventData.onlineMeetingProvider = 'teamsForBusiness';
        }

        // Add attendees if provided
        if (attendees && attendees.length > 0) {
            eventData.attendees = attendees.map(email => ({
                emailAddress: {
                    address: email.trim(),
                    name: email.trim().split('@')[0]
                },
                type: 'required'
            }));
        }

        // Log the payload for debugging
        console.log('Creating calendar event with payload:', JSON.stringify(eventData, null, 2));

        const newEvent = await callGraphApi(req.session.accessToken, '/me/calendar/events', 'POST', eventData);

        console.log('Event created successfully:', newEvent.id);
        console.log('Teams meeting URL:', teamsJoinUrl || newEvent.onlineMeeting?.joinUrl || 'None');

        res.status(201).json({
            success: true,
            event: {
                id: newEvent.id,
                subject: newEvent.subject,
                start: newEvent.start,
                end: newEvent.end,
                onlineMeetingUrl: teamsJoinUrl || newEvent.onlineMeeting?.joinUrl || null,
                webLink: newEvent.webLink
            }
        });
    } catch (error) {
        console.error('Error creating event:', error.response?.data || error.message);
        res.status(500).json({ 
            error: 'Failed to create appointment',
            details: error.response?.data?.error?.message || error.message
        });
    }
});

// Update an existing event
router.put('/events/:eventId', requireAuth, async (req, res) => {
    try {
        const { eventId } = req.params;
        const updateData = req.body;

        const eventPayload = {};

        if (updateData.subject) eventPayload.subject = updateData.subject;
        if (updateData.startDateTime) {
            eventPayload.start = {
                dateTime: updateData.startDateTime,
                timeZone: updateData.timeZone || 'UTC'
            };
        }
        if (updateData.endDateTime) {
            eventPayload.end = {
                dateTime: updateData.endDateTime,
                timeZone: updateData.timeZone || 'UTC'
            };
        }
        if (updateData.body !== undefined) {
            eventPayload.body = {
                contentType: 'HTML',
                content: updateData.body
            };
        }
        if (updateData.location !== undefined) {
            eventPayload.location = {
                displayName: updateData.location
            };
        }

        const updatedEvent = await callGraphApi(
            req.session.accessToken, 
            `/me/calendar/events/${eventId}`, 
            'PATCH', 
            eventPayload
        );

        res.json({
            success: true,
            event: {
                id: updatedEvent.id,
                subject: updatedEvent.subject,
                start: updatedEvent.start,
                end: updatedEvent.end
            }
        });
    } catch (error) {
        console.error('Error updating event:', error.response?.data || error.message);
        res.status(500).json({ error: 'Failed to update appointment' });
    }
});

// Delete an event
router.delete('/events/:eventId', requireAuth, async (req, res) => {
    try {
        const { eventId } = req.params;
        
        await callGraphApi(req.session.accessToken, `/me/calendar/events/${eventId}`, 'DELETE');
        
        res.json({ success: true, message: 'Appointment deleted successfully' });
    } catch (error) {
        console.error('Error deleting event:', error.response?.data || error.message);
        res.status(500).json({ error: 'Failed to delete appointment' });
    }
});

// Create Teams online meeting directly (without calendar event)
router.post('/meeting', requireAuth, async (req, res) => {
    try {
        const { subject, startDateTime, endDateTime } = req.body;

        if (!subject || !startDateTime || !endDateTime) {
            return res.status(400).json({ 
                error: 'Missing required fields: subject, startDateTime, and endDateTime are required' 
            });
        }

        const meetingData = {
            subject,
            startDateTime,
            endDateTime
        };

        const meeting = await callGraphApi(req.session.accessToken, '/me/onlineMeetings', 'POST', meetingData);

        res.status(201).json({
            success: true,
            meeting: {
                id: meeting.id,
                subject: meeting.subject,
                joinUrl: meeting.joinUrl,
                joinWebUrl: meeting.joinWebUrl,
                startDateTime: meeting.startDateTime,
                endDateTime: meeting.endDateTime
            }
        });
    } catch (error) {
        console.error('Error creating meeting:', error.response?.data || error.message);
        res.status(500).json({ 
            error: 'Failed to create Teams meeting',
            details: error.response?.data?.error?.message || error.message
        });
    }
});

module.exports = router;
