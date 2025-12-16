// Global state
let isAuthenticated = false;
let syncedEvents = [];
let lastSyncTime = null;
let currentCalendarDate = new Date();
let selectedDate = null;
let currentView = 'month'; // 'month', 'week', 'day'

// Check authentication status on page load
document.addEventListener('DOMContentLoaded', async () => {
    await checkAuthStatus();
    
    // Check for URL parameters (success/error from OAuth)
    const urlParams = new URLSearchParams(window.location.search);
    if (urlParams.get('success')) {
        showSuccess('Successfully linked your Microsoft Teams account!');
        // Clean URL
        window.history.replaceState({}, document.title, '/');
    }
    if (urlParams.get('error')) {
        showError(urlParams.get('error'));
        window.history.replaceState({}, document.title, '/');
    }

    // Set default datetime values
    setDefaultDateTimes();
    
    // Load cached sync data if available
    loadCachedSyncData();
});

async function checkAuthStatus() {
    try {
        const response = await fetch('/api/status');
        const data = await response.json();
        
        isAuthenticated = data.isAuthenticated;
        
        if (isAuthenticated) {
            document.body.classList.add('logged-in');
            document.getElementById('userInfo').classList.remove('d-none');
            document.getElementById('userInfo').classList.add('d-flex');
            document.getElementById('userName').textContent = data.user.name || data.user.email;
            loadEvents();
        } else {
            document.body.classList.remove('logged-in');
            document.getElementById('userInfo').classList.add('d-none');
            document.getElementById('userInfo').classList.remove('d-flex');
        }
    } catch (error) {
        console.error('Error checking auth status:', error);
    }
}

function setDefaultDateTimes() {
    const now = new Date();
    const startTime = new Date(now.getTime() + 60 * 60 * 1000); // 1 hour from now
    const endTime = new Date(startTime.getTime() + 60 * 60 * 1000); // 1 hour duration
    
    document.getElementById('startDateTime').value = formatDateTimeLocal(startTime);
    document.getElementById('endDateTime').value = formatDateTimeLocal(endTime);
}

function formatDateTimeLocal(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    return `${year}-${month}-${day}T${hours}:${minutes}`;
}

function toggleForm() {
    const form = document.getElementById('appointmentForm');
    const btn = document.getElementById('toggleFormBtn');
    
    if (form.style.display === 'none' || form.style.display === '') {
        form.style.display = 'block';
        btn.innerHTML = '<i class="bi bi-x-lg me-1"></i>Cancel';
        btn.classList.remove('btn-teams');
        btn.classList.add('btn-secondary');
        setDefaultDateTimes();
    } else {
        form.style.display = 'none';
        btn.innerHTML = '<i class="bi bi-plus-lg me-1"></i>New Appointment';
        btn.classList.add('btn-teams');
        btn.classList.remove('btn-secondary');
        document.getElementById('createAppointmentForm').reset();
        hideAlerts();
    }
}

async function createAppointment(event) {
    event.preventDefault();
    
    const submitBtn = document.getElementById('submitBtn');
    submitBtn.classList.add('loading');
    submitBtn.disabled = true;
    
    hideAlerts();
    
    const subject = document.getElementById('subject').value.trim();
    const startDateTime = document.getElementById('startDateTime').value;
    const endDateTime = document.getElementById('endDateTime').value;
    const location = document.getElementById('location').value.trim();
    const attendeesInput = document.getElementById('attendees').value.trim();
    const body = document.getElementById('body').value.trim();
    const isOnlineMeeting = document.getElementById('isOnlineMeeting').checked;
    
    // Parse attendees
    const attendees = attendeesInput 
        ? attendeesInput.split(',').map(email => email.trim()).filter(email => email)
        : [];
    
    // Validate dates
    if (new Date(endDateTime) <= new Date(startDateTime)) {
        showError('End time must be after start time');
        submitBtn.classList.remove('loading');
        submitBtn.disabled = false;
        return;
    }
    
    try {
        const response = await fetch('/api/calendar/events', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                subject,
                startDateTime,
                endDateTime,
                location,
                attendees,
                body,
                isOnlineMeeting,
                timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone
            })
        });
        
        const data = await response.json();
        
        if (response.ok) {
            showSuccess(`Appointment "${subject}" created successfully!`, data.event.onlineMeetingUrl);
            toggleForm();
            loadEvents();
        } else {
            showError(data.details || data.error || 'Failed to create appointment');
        }
    } catch (error) {
        console.error('Error creating appointment:', error);
        showError('Network error. Please try again.');
    } finally {
        submitBtn.classList.remove('loading');
        submitBtn.disabled = false;
    }
}

async function loadEvents() {
    const eventsList = document.getElementById('eventsList');
    const eventsLoading = document.getElementById('eventsLoading');
    const noEvents = document.getElementById('noEvents');
    
    eventsLoading.classList.remove('d-none');
    eventsList.innerHTML = '';
    noEvents.classList.add('d-none');
    
    try {
        // Get events for the next 30 days
        const startDateTime = new Date().toISOString();
        const endDateTime = new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toISOString();
        
        const response = await fetch(`/api/calendar/events?startDateTime=${startDateTime}&endDateTime=${endDateTime}`);
        const data = await response.json();
        
        eventsLoading.classList.add('d-none');
        
        if (data.events && data.events.length > 0) {
            eventsList.innerHTML = data.events.map(event => createEventCard(event)).join('');
        } else {
            noEvents.classList.remove('d-none');
        }
    } catch (error) {
        console.error('Error loading events:', error);
        eventsLoading.classList.add('d-none');
        eventsList.innerHTML = `
            <div class="alert alert-warning">
                <i class="bi bi-exclamation-triangle me-2"></i>
                Failed to load appointments. Please try again.
            </div>
        `;
    }
}

function createEventCard(event) {
    const startDate = new Date(event.start.dateTime + 'Z');
    const endDate = new Date(event.end.dateTime + 'Z');
    
    const dateOptions = { weekday: 'short', month: 'short', day: 'numeric' };
    const timeOptions = { hour: '2-digit', minute: '2-digit' };
    
    const formattedDate = startDate.toLocaleDateString('en-US', dateOptions);
    const formattedStartTime = startDate.toLocaleTimeString('en-US', timeOptions);
    const formattedEndTime = endDate.toLocaleTimeString('en-US', timeOptions);
    
    const attendeesList = event.attendees && event.attendees.length > 0
        ? `<div class="mt-2">
            <small class="text-muted">
                <i class="bi bi-people me-1"></i>
                ${event.attendees.map(a => a.name || a.email).join(', ')}
            </small>
           </div>`
        : '';
    
    const meetingLink = event.onlineMeetingUrl
        ? `<a href="${event.onlineMeetingUrl}" target="_blank" class="btn btn-sm btn-outline-primary mt-2">
            <i class="bi bi-camera-video me-1"></i>Join Teams Meeting
           </a>`
        : '';
    
    const locationBadge = event.location
        ? `<span class="badge bg-secondary me-2"><i class="bi bi-geo-alt me-1"></i>${event.location}</span>`
        : '';
    
    const teamsBadge = event.isOnlineMeeting
        ? `<span class="badge teams-badge"><i class="bi bi-camera-video me-1"></i>Teams</span>`
        : '';
    
    return `
        <div class="card event-card mb-3">
            <div class="card-body">
                <div class="d-flex justify-content-between align-items-start">
                    <div class="flex-grow-1">
                        <h5 class="card-title mb-1">${escapeHtml(event.subject)}</h5>
                        <p class="card-text mb-2">
                            <i class="bi bi-clock me-1"></i>
                            ${formattedDate} · ${formattedStartTime} - ${formattedEndTime}
                        </p>
                        <div>
                            ${locationBadge}
                            ${teamsBadge}
                        </div>
                        ${attendeesList}
                        ${meetingLink}
                    </div>
                    <div class="btn-group">
                        <button class="btn btn-sm btn-outline-danger" onclick="deleteEvent('${event.id}')" title="Delete">
                            <i class="bi bi-trash"></i>
                        </button>
                    </div>
                </div>
            </div>
        </div>
    `;
}

async function deleteEvent(eventId) {
    if (!confirm('Are you sure you want to delete this appointment?')) {
        return;
    }
    
    try {
        const response = await fetch(`/api/calendar/events/${eventId}`, {
            method: 'DELETE'
        });
        
        if (response.ok) {
            showSuccess('Appointment deleted successfully');
            loadEvents();
        } else {
            const data = await response.json();
            showError(data.error || 'Failed to delete appointment');
        }
    } catch (error) {
        console.error('Error deleting event:', error);
        showError('Network error. Please try again.');
    }
}

function showSuccess(message, meetingUrl = null) {
    const alert = document.getElementById('successAlert');
    const messageEl = document.getElementById('successMessage');
    const linkContainer = document.getElementById('meetingLinkContainer');
    const meetingLink = document.getElementById('meetingLink');
    
    messageEl.textContent = message;
    
    if (meetingUrl) {
        linkContainer.classList.remove('d-none');
        meetingLink.href = meetingUrl;
        meetingLink.textContent = meetingUrl;
    } else {
        linkContainer.classList.add('d-none');
    }
    
    alert.classList.remove('d-none');
    
    // Auto-hide after 10 seconds
    setTimeout(() => {
        alert.classList.add('d-none');
    }, 10000);
}

function showError(message) {
    const alert = document.getElementById('errorAlert');
    document.getElementById('errorMessage').textContent = message;
    alert.classList.remove('d-none');
    
    // Auto-hide after 5 seconds
    setTimeout(() => {
        alert.classList.add('d-none');
    }, 5000);
}

function hideAlerts() {
    document.getElementById('successAlert').classList.add('d-none');
    document.getElementById('errorAlert').classList.add('d-none');
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// ============== CALENDAR SYNC FUNCTIONS ==============

function loadCachedSyncData() {
    const cached = localStorage.getItem('syncedCalendarEvents');
    const cachedTime = localStorage.getItem('lastSyncTime');
    
    if (cached && cachedTime) {
        try {
            syncedEvents = JSON.parse(cached);
            lastSyncTime = cachedTime;
            updateSyncUI();
        } catch (e) {
            console.error('Error loading cached sync data:', e);
        }
    }
}

async function syncCalendar(days = 30) {
    const syncBtn = document.getElementById('syncBtn');
    const syncStatus = document.getElementById('syncStatus');
    
    // Show loading state
    syncBtn.disabled = true;
    syncBtn.innerHTML = '<span class="spinner-border spinner-border-sm me-1"></span>Syncing...';
    syncStatus.classList.remove('d-none');
    
    try {
        const response = await fetch(`/api/calendar/sync?days=${days}`);
        const data = await response.json();
        
        if (response.ok && data.success) {
            syncedEvents = data.events;
            lastSyncTime = data.syncedAt;
            
            // Cache the data
            localStorage.setItem('syncedCalendarEvents', JSON.stringify(syncedEvents));
            localStorage.setItem('lastSyncTime', lastSyncTime);
            
            updateSyncUI();
            showSuccess(`Successfully synced ${data.totalEvents} events from your Microsoft calendar!`);
            
            // Switch to synced tab
            const syncedTab = document.getElementById('synced-tab');
            if (syncedTab) {
                const tab = new bootstrap.Tab(syncedTab);
                tab.show();
            }
        } else {
            showError(data.error || 'Failed to sync calendar');
        }
    } catch (error) {
        console.error('Error syncing calendar:', error);
        showError('Network error while syncing calendar');
    } finally {
        syncBtn.disabled = false;
        syncBtn.innerHTML = '<i class="bi bi-arrow-repeat me-1"></i>Sync Calendar';
        syncStatus.classList.add('d-none');
    }
}

function updateSyncUI() {
    const syncInfoCard = document.getElementById('syncInfoCard');
    const lastSyncTimeEl = document.getElementById('lastSyncTime');
    const syncedEventCount = document.getElementById('syncedEventCount');
    const syncBadge = document.getElementById('syncBadge');
    const noSyncedEvents = document.getElementById('noSyncedEvents');
    const calendarContainer = document.getElementById('calendarContainer');
    
    if (syncedEvents.length > 0) {
        syncInfoCard.classList.remove('d-none');
        lastSyncTimeEl.textContent = formatDateTime(new Date(lastSyncTime));
        syncedEventCount.textContent = syncedEvents.length;
        
        syncBadge.textContent = syncedEvents.length;
        syncBadge.classList.remove('d-none');
        
        noSyncedEvents.classList.add('d-none');
        calendarContainer.classList.remove('d-none');
        
        renderCurrentView();
        renderTeamsMeetings();
    } else {
        syncInfoCard.classList.add('d-none');
        syncBadge.classList.add('d-none');
        noSyncedEvents.classList.remove('d-none');
        calendarContainer.classList.add('d-none');
    }
}

function renderTeamsMeetings() {
    const teamsMeetingsList = document.getElementById('teamsMeetingsList');
    const noTeamsMeetings = document.getElementById('noTeamsMeetings');
    
    const teamsMeetings = syncedEvents.filter(event => event.isOnlineMeeting && event.onlineMeetingUrl);
    
    if (teamsMeetings.length > 0) {
        noTeamsMeetings.classList.add('d-none');
        teamsMeetingsList.innerHTML = teamsMeetings.map(event => createSyncedEventCard(event)).join('');
    } else {
        noTeamsMeetings.classList.remove('d-none');
        teamsMeetingsList.innerHTML = '';
    }
}

function formatDateTime(date) {
    return date.toLocaleString('en-US', {
        month: 'short',
        day: 'numeric',
        year: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
    });
}

function createSyncedEventCard(event) {
    const startDate = new Date(event.start.dateTime + 'Z');
    const endDate = new Date(event.end.dateTime + 'Z');
    
    const dateOptions = { weekday: 'short', month: 'short', day: 'numeric' };
    const timeOptions = { hour: '2-digit', minute: '2-digit' };
    
    const formattedDate = startDate.toLocaleDateString('en-US', dateOptions);
    const formattedStartTime = startDate.toLocaleTimeString('en-US', timeOptions);
    const formattedEndTime = endDate.toLocaleTimeString('en-US', timeOptions);
    
    const locationBadge = event.location
        ? `<span class="badge bg-secondary me-2"><i class="bi bi-geo-alt me-1"></i>${escapeHtml(event.location)}</span>`
        : '';
    
    const teamsBadge = event.isOnlineMeeting
        ? `<span class="badge teams-badge"><i class="bi bi-camera-video me-1"></i>Teams</span>`
        : '';
    
    const meetingLink = event.onlineMeetingUrl
        ? `<a href="${event.onlineMeetingUrl}" target="_blank" class="btn btn-sm btn-outline-primary mt-2">
            <i class="bi bi-camera-video me-1"></i>Join Teams Meeting
           </a>`
        : '';
    
    const organizerInfo = event.organizer 
        ? `<small class="text-muted d-block"><i class="bi bi-person me-1"></i>Organizer: ${escapeHtml(event.organizer)}</small>`
        : '';
    
    const importanceBadge = event.importance === 'high'
        ? `<span class="badge bg-danger me-2"><i class="bi bi-exclamation-circle me-1"></i>Important</span>`
        : '';
    
    const recurringBadge = event.isRecurring
        ? `<span class="badge bg-info me-2"><i class="bi bi-arrow-repeat me-1"></i>Recurring</span>`
        : '';

    return `
        <div class="card event-card mb-3 ${event.isCancelled ? 'opacity-50' : ''}" style="cursor: pointer;" onclick="showEventDetails('${event.id}')">
            <div class="card-body">
                <div class="d-flex justify-content-between align-items-start">
                    <div class="flex-grow-1">
                        <h5 class="card-title mb-1">
                            ${event.isCancelled ? '<s>' : ''}${escapeHtml(event.subject)}${event.isCancelled ? '</s>' : ''}
                            ${event.isCancelled ? '<span class="badge bg-danger ms-2">Cancelled</span>' : ''}
                        </h5>
                        <p class="card-text mb-2">
                            <i class="bi bi-clock me-1"></i>
                            ${formattedDate} · ${formattedStartTime} - ${formattedEndTime}
                        </p>
                        <div class="mb-2">
                            ${importanceBadge}
                            ${recurringBadge}
                            ${locationBadge}
                            ${teamsBadge}
                        </div>
                        ${organizerInfo}
                        ${event.bodyPreview ? `<small class="text-muted d-block mt-2">${escapeHtml(event.bodyPreview)}...</small>` : ''}
                        ${meetingLink}
                    </div>
                    <div>
                        ${event.webLink ? `<a href="${event.webLink}" target="_blank" class="btn btn-sm btn-outline-secondary" title="Open in Outlook" onclick="event.stopPropagation()">
                            <i class="bi bi-box-arrow-up-right"></i>
                        </a>` : ''}
                    </div>
                </div>
            </div>
        </div>
    `;
}

// ============== EVENT DETAILS MODAL ==============

function showEventDetails(eventId) {
    const event = syncedEvents.find(e => e.id === eventId);
    if (!event) return;
    
    const startDate = new Date(event.start.dateTime + 'Z');
    const endDate = new Date(event.end.dateTime + 'Z');
    
    // Set title
    document.getElementById('modalEventTitle').innerHTML = event.isCancelled 
        ? `<s>${escapeHtml(event.subject)}</s> <span class="badge bg-danger">Cancelled</span>`
        : escapeHtml(event.subject);
    
    // Set date and time
    document.getElementById('modalEventDate').textContent = startDate.toLocaleDateString('en-US', {
        weekday: 'long',
        month: 'long',
        day: 'numeric',
        year: 'numeric'
    });
    document.getElementById('modalEventTime').textContent = `${startDate.toLocaleTimeString('en-US', {
        hour: '2-digit',
        minute: '2-digit'
    })} - ${endDate.toLocaleTimeString('en-US', {
        hour: '2-digit',
        minute: '2-digit'
    })}`;
    
    // Set location
    const locationSection = document.getElementById('modalLocationSection');
    if (event.location) {
        document.getElementById('modalEventLocation').textContent = event.location;
        locationSection.classList.remove('d-none');
    } else {
        locationSection.classList.add('d-none');
    }
    
    // Set organizer
    const organizerSection = document.getElementById('modalOrganizerSection');
    if (event.organizer) {
        document.getElementById('modalEventOrganizer').textContent = event.organizer;
        organizerSection.classList.remove('d-none');
    } else {
        organizerSection.classList.add('d-none');
    }
    
    // Set attendees
    const attendeesSection = document.getElementById('modalAttendeesSection');
    if (event.attendees && event.attendees.length > 0) {
        const attendeesList = event.attendees.map(a => {
            const statusIcon = getAttendeeStatusIcon(a.status);
            return `<div class="small">${statusIcon} ${escapeHtml(a.name || a.email)}</div>`;
        }).join('');
        document.getElementById('modalEventAttendees').innerHTML = attendeesList;
        attendeesSection.classList.remove('d-none');
    } else {
        attendeesSection.classList.add('d-none');
    }
    
    // Set description
    const descriptionSection = document.getElementById('modalDescriptionSection');
    if (event.bodyPreview) {
        document.getElementById('modalEventDescription').textContent = event.bodyPreview;
        descriptionSection.classList.remove('d-none');
    } else {
        descriptionSection.classList.add('d-none');
    }
    
    // Set badges
    let badges = '';
    if (event.importance === 'high') {
        badges += '<span class="badge bg-danger me-1"><i class="bi bi-exclamation-circle me-1"></i>Important</span>';
    }
    if (event.isRecurring) {
        badges += '<span class="badge bg-info me-1"><i class="bi bi-arrow-repeat me-1"></i>Recurring</span>';
    }
    if (event.isOnlineMeeting) {
        badges += '<span class="badge teams-badge me-1"><i class="bi bi-camera-video me-1"></i>Teams Meeting</span>';
    }
    document.getElementById('modalEventBadges').innerHTML = badges;
    
    // Set Teams link
    const teamsSection = document.getElementById('modalTeamsSection');
    if (event.onlineMeetingUrl) {
        document.getElementById('modalTeamsLink').href = event.onlineMeetingUrl;
        teamsSection.classList.remove('d-none');
    } else {
        teamsSection.classList.add('d-none');
    }
    
    // Set Outlook link
    const outlookLink = document.getElementById('modalOutlookLink');
    if (event.webLink) {
        outlookLink.href = event.webLink;
        outlookLink.classList.remove('d-none');
    } else {
        outlookLink.classList.add('d-none');
    }
    
    // Show modal
    const modal = new bootstrap.Modal(document.getElementById('eventDetailModal'));
    modal.show();
}

function getAttendeeStatusIcon(status) {
    switch (status) {
        case 'accepted':
            return '<i class="bi bi-check-circle-fill text-success"></i>';
        case 'declined':
            return '<i class="bi bi-x-circle-fill text-danger"></i>';
        case 'tentativelyAccepted':
        case 'tentative':
            return '<i class="bi bi-question-circle-fill text-warning"></i>';
        default:
            return '<i class="bi bi-circle text-muted"></i>';
    }
}

// ============== CALENDAR VIEW FUNCTIONS ==============

function switchView(view) {
    currentView = view;
    
    // Update button states
    document.getElementById('monthViewBtn').classList.remove('active');
    document.getElementById('weekViewBtn').classList.remove('active');
    document.getElementById('dayViewBtn').classList.remove('active');
    document.getElementById(`${view}ViewBtn`).classList.add('active');
    
    // Hide all views
    document.getElementById('monthView').classList.add('d-none');
    document.getElementById('weekView').classList.add('d-none');
    document.getElementById('dayView').classList.add('d-none');
    
    // Show selected view
    document.getElementById(`${view}View`).classList.remove('d-none');
    
    // Hide selected day events panel in non-month views
    document.getElementById('selectedDayEvents').classList.add('d-none');
    
    renderCurrentView();
}

function renderCurrentView() {
    switch (currentView) {
        case 'month':
            renderCalendar();
            break;
        case 'week':
            renderWeekView();
            break;
        case 'day':
            renderDayView();
            break;
    }
}

function navigateCalendar(delta) {
    switch (currentView) {
        case 'month':
            currentCalendarDate.setMonth(currentCalendarDate.getMonth() + delta);
            break;
        case 'week':
            currentCalendarDate.setDate(currentCalendarDate.getDate() + (delta * 7));
            break;
        case 'day':
            currentCalendarDate.setDate(currentCalendarDate.getDate() + delta);
            break;
    }
    document.getElementById('selectedDayEvents').classList.add('d-none');
    selectedDate = null;
    renderCurrentView();
}

function goToToday() {
    currentCalendarDate = new Date();
    selectedDate = null;
    document.getElementById('selectedDayEvents').classList.add('d-none');
    renderCurrentView();
}

function renderCalendar() {
    const calendarBody = document.getElementById('calendarBody');
    const currentMonthYear = document.getElementById('currentMonthYear');
    
    const year = currentCalendarDate.getFullYear();
    const month = currentCalendarDate.getMonth();
    
    // Update header
    currentMonthYear.textContent = new Date(year, month).toLocaleDateString('en-US', { 
        month: 'long', 
        year: 'numeric' 
    });
    
    // Get first day of month and total days
    const firstDay = new Date(year, month, 1).getDay();
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    const daysInPrevMonth = new Date(year, month, 0).getDate();
    
    // Get today for highlighting
    const today = new Date();
    const isCurrentMonth = today.getFullYear() === year && today.getMonth() === month;
    
    // Group events by date
    const eventsByDate = {};
    syncedEvents.forEach(event => {
        const eventDate = new Date(event.start.dateTime + 'Z');
        const dateKey = `${eventDate.getFullYear()}-${eventDate.getMonth()}-${eventDate.getDate()}`;
        if (!eventsByDate[dateKey]) {
            eventsByDate[dateKey] = [];
        }
        eventsByDate[dateKey].push(event);
    });
    
    let html = '';
    let dayCount = 1;
    let nextMonthDay = 1;
    
    // Create 6 rows for the calendar
    for (let week = 0; week < 6; week++) {
        html += '<tr>';
        
        for (let dayOfWeek = 0; dayOfWeek < 7; dayOfWeek++) {
            const cellIndex = week * 7 + dayOfWeek;
            
            if (cellIndex < firstDay) {
                // Previous month days
                const prevDay = daysInPrevMonth - firstDay + cellIndex + 1;
                html += `<td class="other-month">
                    <div class="day-number">${prevDay}</div>
                </td>`;
            } else if (dayCount <= daysInMonth) {
                // Current month days
                const dateKey = `${year}-${month}-${dayCount}`;
                const dayEvents = eventsByDate[dateKey] || [];
                const isToday = isCurrentMonth && today.getDate() === dayCount;
                const isSelected = selectedDate && 
                    selectedDate.getFullYear() === year && 
                    selectedDate.getMonth() === month && 
                    selectedDate.getDate() === dayCount;
                
                let classes = '';
                if (isToday) classes += ' today';
                if (isSelected) classes += ' selected';
                
                html += `<td class="${classes}" onclick="selectDate(${year}, ${month}, ${dayCount})">
                    <div class="day-number">${dayCount}</div>
                    <div class="day-events">
                        ${renderDayEvents(dayEvents, 3)}
                    </div>
                </td>`;
                dayCount++;
            } else {
                // Next month days
                html += `<td class="other-month">
                    <div class="day-number">${nextMonthDay}</div>
                </td>`;
                nextMonthDay++;
            }
        }
        
        html += '</tr>';
        
        // Stop if we've rendered all days and at least 4 weeks
        if (dayCount > daysInMonth && week >= 3) break;
    }
    
    calendarBody.innerHTML = html;
}

function renderDayEvents(events, maxShow = 3) {
    if (events.length === 0) return '';
    
    let html = '';
    const showEvents = events.slice(0, maxShow);
    
    showEvents.forEach(event => {
        const eventTime = new Date(event.start.dateTime + 'Z').toLocaleTimeString('en-US', {
            hour: 'numeric',
            minute: '2-digit'
        });
        const isTeams = event.isOnlineMeeting ? 'teams-meeting' : '';
        html += `<div class="day-event ${isTeams}" 
            title="${escapeHtml(event.subject)} - ${eventTime}"
            onclick="event.stopPropagation(); showEventDetails('${event.id}')">
            ${escapeHtml(event.subject)}
        </div>`;
    });
    
    if (events.length > maxShow) {
        html += `<div class="more-events">+${events.length - maxShow} more</div>`;
    }
    
    return html;
}

function changeMonth(delta) {
    currentCalendarDate.setMonth(currentCalendarDate.getMonth() + delta);
    renderCalendar();
    // Hide selected day events when changing month
    document.getElementById('selectedDayEvents').classList.add('d-none');
    selectedDate = null;
}

// ============== WEEK VIEW ==============

function renderWeekView() {
    const currentMonthYear = document.getElementById('currentMonthYear');
    const weekHeader = document.getElementById('weekHeader');
    const weekBody = document.getElementById('weekBody');
    
    // Get the start of the week (Sunday)
    const startOfWeek = new Date(currentCalendarDate);
    startOfWeek.setDate(startOfWeek.getDate() - startOfWeek.getDay());
    
    const endOfWeek = new Date(startOfWeek);
    endOfWeek.setDate(endOfWeek.getDate() + 6);
    
    // Update header title
    const startMonth = startOfWeek.toLocaleDateString('en-US', { month: 'short' });
    const endMonth = endOfWeek.toLocaleDateString('en-US', { month: 'short' });
    const year = startOfWeek.getFullYear();
    
    if (startMonth === endMonth) {
        currentMonthYear.textContent = `${startMonth} ${startOfWeek.getDate()} - ${endOfWeek.getDate()}, ${year}`;
    } else {
        currentMonthYear.textContent = `${startMonth} ${startOfWeek.getDate()} - ${endMonth} ${endOfWeek.getDate()}, ${year}`;
    }
    
    const today = new Date();
    
    // Build header
    let headerHtml = '<tr><th class="time-label"></th>';
    for (let i = 0; i < 7; i++) {
        const day = new Date(startOfWeek);
        day.setDate(day.getDate() + i);
        const isToday = day.toDateString() === today.toDateString();
        const dayName = day.toLocaleDateString('en-US', { weekday: 'short' });
        const dayNum = day.getDate();
        headerHtml += `<th class="${isToday ? 'today' : ''}">${dayName}<br>${dayNum}</th>`;
    }
    headerHtml += '</tr>';
    weekHeader.innerHTML = headerHtml;
    
    // Group events by day and hour
    const eventsByDayHour = {};
    syncedEvents.forEach(event => {
        const eventDate = new Date(event.start.dateTime + 'Z');
        const dayKey = eventDate.toDateString();
        const hour = eventDate.getHours();
        const key = `${dayKey}-${hour}`;
        if (!eventsByDayHour[key]) {
            eventsByDayHour[key] = [];
        }
        eventsByDayHour[key].push(event);
    });
    
    // Build body - show hours from 7 AM to 9 PM
    let bodyHtml = '';
    for (let hour = 7; hour <= 21; hour++) {
        bodyHtml += '<tr>';
        bodyHtml += `<td class="time-label">${formatHour(hour)}</td>`;
        
        for (let dayOffset = 0; dayOffset < 7; dayOffset++) {
            const day = new Date(startOfWeek);
            day.setDate(day.getDate() + dayOffset);
            const key = `${day.toDateString()}-${hour}`;
            const events = eventsByDayHour[key] || [];
            
            bodyHtml += '<td>';
            events.forEach(event => {
                const eventTime = new Date(event.start.dateTime + 'Z').toLocaleTimeString('en-US', {
                    hour: 'numeric',
                    minute: '2-digit'
                });
                bodyHtml += `<div class="week-event" onclick="showEventDetails('${event.id}')" title="${escapeHtml(event.subject)}">
                    ${escapeHtml(event.subject)}
                </div>`;
            });
            bodyHtml += '</td>';
        }
        bodyHtml += '</tr>';
    }
    weekBody.innerHTML = bodyHtml;
}

function formatHour(hour) {
    if (hour === 0) return '12 AM';
    if (hour === 12) return '12 PM';
    if (hour < 12) return `${hour} AM`;
    return `${hour - 12} PM`;
}

// ============== DAY VIEW ==============

function renderDayView() {
    const currentMonthYear = document.getElementById('currentMonthYear');
    const dayViewContent = document.getElementById('dayViewContent');
    
    const today = new Date();
    const isToday = currentCalendarDate.toDateString() === today.toDateString();
    
    // Update header
    currentMonthYear.textContent = currentCalendarDate.toLocaleDateString('en-US', {
        weekday: 'long',
        month: 'long',
        day: 'numeric',
        year: 'numeric'
    });
    
    // Get events for this day
    const dayEvents = syncedEvents.filter(event => {
        const eventDate = new Date(event.start.dateTime + 'Z');
        return eventDate.toDateString() === currentCalendarDate.toDateString();
    }).sort((a, b) => new Date(a.start.dateTime) - new Date(b.start.dateTime));
    
    // Group events by hour
    const eventsByHour = {};
    dayEvents.forEach(event => {
        const eventDate = new Date(event.start.dateTime + 'Z');
        const hour = eventDate.getHours();
        if (!eventsByHour[hour]) {
            eventsByHour[hour] = [];
        }
        eventsByHour[hour].push(event);
    });
    
    // Build day view
    let html = `
        <div class="day-view-header">
            <h3>${currentCalendarDate.toLocaleDateString('en-US', { weekday: 'long' })}</h3>
            <p>${currentCalendarDate.toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' })}
            ${isToday ? ' <span class="badge bg-light text-dark">Today</span>' : ''}</p>
            <p class="mt-2"><strong>${dayEvents.length}</strong> event${dayEvents.length !== 1 ? 's' : ''}</p>
        </div>
    `;
    
    if (dayEvents.length === 0) {
        html += `
            <div class="no-events-day">
                <i class="bi bi-calendar-x display-4"></i>
                <p class="mt-2">No events scheduled for this day</p>
            </div>
        `;
    } else {
        // Show hours from 6 AM to 10 PM
        for (let hour = 6; hour <= 22; hour++) {
            const events = eventsByHour[hour] || [];
            const hasEvents = events.length > 0;
            
            html += `
                <div class="hour-slot ${hasEvents ? 'has-events' : ''}">
                    <div class="hour-label">${formatHour(hour)}</div>
                    <div class="hour-content">
            `;
            
            events.forEach(event => {
                const startTime = new Date(event.start.dateTime + 'Z');
                const endTime = new Date(event.end.dateTime + 'Z');
                const timeStr = `${startTime.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' })} - ${endTime.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' })}`;
                const isTeams = event.isOnlineMeeting ? 'teams-meeting' : '';
                
                html += `
                    <div class="hour-event ${isTeams}" onclick="showEventDetails('${event.id}')">
                        <div class="event-time">
                            <i class="bi bi-clock me-1"></i>${timeStr}
                            ${event.isOnlineMeeting ? '<i class="bi bi-camera-video ms-2"></i>' : ''}
                        </div>
                        <div class="event-title">${escapeHtml(event.subject)}</div>
                        ${event.location ? `<div class="event-location small opacity-75"><i class="bi bi-geo-alt me-1"></i>${escapeHtml(event.location)}</div>` : ''}
                    </div>
                `;
            });
            
            html += `
                    </div>
                </div>
            `;
        }
    }
    
    dayViewContent.innerHTML = html;
}

function selectDate(year, month, day) {
    selectedDate = new Date(year, month, day);
    
    // Double-click behavior: if already selected, switch to day view
    if (currentCalendarDate.toDateString() === selectedDate.toDateString() && currentView === 'month') {
        currentCalendarDate = new Date(year, month, day);
        switchView('day');
        return;
    }
    
    currentCalendarDate = new Date(year, month, day);
    renderCalendar(); // Re-render to show selection
    showSelectedDayEvents(year, month, day);
}

function showSelectedDayEvents(year, month, day) {
    const selectedDayEvents = document.getElementById('selectedDayEvents');
    const selectedDayTitle = document.getElementById('selectedDayTitle');
    const selectedDayEventsList = document.getElementById('selectedDayEventsList');
    
    const dateKey = `${year}-${month}-${day}`;
    const dateObj = new Date(year, month, day);
    
    // Filter events for this day
    const dayEvents = syncedEvents.filter(event => {
        const eventDate = new Date(event.start.dateTime + 'Z');
        return eventDate.getFullYear() === year && 
               eventDate.getMonth() === month && 
               eventDate.getDate() === day;
    });
    
    selectedDayTitle.innerHTML = `<i class="bi bi-calendar-event me-2"></i>${dateObj.toLocaleDateString('en-US', { 
        weekday: 'long', 
        month: 'long', 
        day: 'numeric',
        year: 'numeric'
    })} (${dayEvents.length} event${dayEvents.length !== 1 ? 's' : ''})`;
    
    if (dayEvents.length > 0) {
        selectedDayEventsList.innerHTML = dayEvents.map(event => createSyncedEventCard(event)).join('');
    } else {
        selectedDayEventsList.innerHTML = `
            <div class="text-center text-muted py-3">
                <i class="bi bi-calendar-x"></i> No events on this day
            </div>
        `;
    }
    
    selectedDayEvents.classList.remove('d-none');
}
