export interface OwaRequestOptions {
  action: string;
  body: Record<string, unknown>;
  token: string;
}

export interface OwaError {
  code: string;
  message: string;
}

export interface OwaResponse<T = unknown> {
  ok: boolean;
  status: number;
  data?: T;
  error?: OwaError;
}

const USER_AGENT =
  'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36';

export async function owaRequest<T = unknown>(
  options: OwaRequestOptions
): Promise<OwaResponse<T>> {
  const { action, body, token } = options;
  const url = `https://outlook.office.com/owa/service.svc?action=${action}&app=Mail&n=0`;

  try {
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json; charset=utf-8',
        Authorization: `Bearer ${token}`,
        'User-Agent': USER_AGENT,
        Accept: 'application/json',
      },
      body: JSON.stringify(body),
    });

    if (!response.ok) {
      return {
        ok: false,
        status: response.status,
        error: {
          code: `HTTP_${response.status}`,
          message: response.statusText,
        },
      };
    }

    const data = (await response.json()) as T;
    return { ok: true, status: response.status, data };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

export interface UserConfiguration {
  SessionSettings?: {
    UserDisplayName?: string;
    UserEmailAddress?: string;
  };
}

export async function getUserConfiguration(
  token: string
): Promise<OwaResponse<UserConfiguration>> {
  return owaRequest<UserConfiguration>({
    action: 'GetUserConfiguration',
    body: {
      __type: 'GetUserConfigurationRequest:#Exchange',
      Header: {
        __type: 'JsonRequestHeaders:#Exchange',
        RequestServerVersion: 'Exchange2016',
      },
      Body: {
        __type: 'GetUserConfigurationRequest:#Exchange',
        UserConfigurationName: {
          __type: 'UserConfigurationNameType:#Exchange',
          Name: 'OWA.SessionData',
        },
        UserConfigurationProperties: 'All',
      },
    },
    token,
  });
}

export interface OwaUserInfo {
  displayName: string;
  email: string;
}

export async function getOwaUserInfo(
  token: string
): Promise<OwaResponse<OwaUserInfo>> {
  // Use Outlook REST API to get user info
  const url = 'https://outlook.office.com/api/v2.0/me';

  try {
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        'User-Agent': USER_AGENT,
        Accept: 'application/json',
      },
    });

    if (!response.ok) {
      return {
        ok: false,
        status: response.status,
        error: {
          code: `HTTP_${response.status}`,
          message: response.statusText,
        },
      };
    }

    const data = await response.json() as { DisplayName?: string; EmailAddress?: string };

    return {
      ok: true,
      status: response.status,
      data: {
        displayName: data.DisplayName || 'Unknown',
        email: data.EmailAddress || 'Unknown',
      },
    };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

export interface CalendarAttendee {
  Type: 'Required' | 'Optional' | 'Resource';
  Status: {
    Response: 'None' | 'Organizer' | 'TentativelyAccepted' | 'Accepted' | 'Declined' | 'NotResponded';
    Time: string;
  };
  EmailAddress: {
    Name: string;
    Address: string;
  };
}

export interface CalendarEvent {
  Id: string;
  Subject: string;
  Start: { DateTime: string; TimeZone: string };
  End: { DateTime: string; TimeZone: string };
  Location?: { DisplayName?: string };
  Organizer?: { EmailAddress?: { Name?: string; Address?: string } };
  Attendees?: CalendarAttendee[];
  IsAllDay?: boolean;
  IsCancelled?: boolean;
  IsOrganizer?: boolean;
  BodyPreview?: string;
  Categories?: string[];
  ShowAs?: string;
  Importance?: string;
  IsOnlineMeeting?: boolean;
  OnlineMeetingUrl?: string;
  WebLink?: string;
}

export interface CalendarViewResponse {
  value: CalendarEvent[];
}

export async function getCalendarEvents(
  token: string,
  startDateTime: string,
  endDateTime: string
): Promise<OwaResponse<CalendarEvent[]>> {
  const url = `https://outlook.office.com/api/v2.0/me/calendarview?startDateTime=${encodeURIComponent(startDateTime)}&endDateTime=${encodeURIComponent(endDateTime)}&$orderby=Start/DateTime`;

  try {
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        'User-Agent': USER_AGENT,
        Accept: 'application/json',
        Prefer: 'outlook.timezone="Europe/Amsterdam"',
      },
    });

    if (!response.ok) {
      return {
        ok: false,
        status: response.status,
        error: {
          code: `HTTP_${response.status}`,
          message: response.statusText,
        },
      };
    }

    const data = (await response.json()) as CalendarViewResponse;

    return {
      ok: true,
      status: response.status,
      data: data.value,
    };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

export interface FreeBusySlot {
  status: 'Free' | 'Busy' | 'Tentative';
  start: string;
  end: string;
  subject?: string;
}

export interface ScheduleInfo {
  scheduleId: string;
  availabilityView: string;
  scheduleItems: Array<{
    status: string;
    start: { dateTime: string; timeZone: string };
    end: { dateTime: string; timeZone: string };
    subject?: string;
    location?: string;
  }>;
}

/**
 * Get schedule/availability for multiple users using Microsoft Graph API.
 * Requires a Graph API token with Calendars.Read.Shared permission.
 */
export async function getScheduleForUsers(
  graphToken: string,
  emails: string[],
  startDateTime: string,
  endDateTime: string
): Promise<OwaResponse<ScheduleInfo[]>> {
  // Try getSchedule endpoint first
  const url = 'https://graph.microsoft.com/v1.0/me/calendar/getSchedule';

  try {
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${graphToken}`,
        'Content-Type': 'application/json',
        Accept: 'application/json',
      },
      body: JSON.stringify({
        schedules: emails,
        startTime: {
          dateTime: startDateTime,
          timeZone: 'Europe/Amsterdam',
        },
        endTime: {
          dateTime: endDateTime,
          timeZone: 'Europe/Amsterdam',
        },
        availabilityViewInterval: 30,
      }),
    });

    if (!response.ok) {
      const errorData = await response.json().catch(() => ({})) as { error?: { code?: string; message?: string } };
      return {
        ok: false,
        status: response.status,
        error: {
          code: errorData.error?.code || `HTTP_${response.status}`,
          message: errorData.error?.message || response.statusText,
        },
      };
    }

    const data = await response.json() as { value: ScheduleInfo[] };

    return {
      ok: true,
      status: response.status,
      data: data.value,
    };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

/**
 * Get schedule/availability for users using Outlook REST API.
 * Uses the same token as other Outlook API calls.
 */
export async function getScheduleViaOutlook(
  token: string,
  emails: string[],
  startDateTime: string,
  endDateTime: string,
  durationMinutes: number = 30
): Promise<OwaResponse<ScheduleInfo[]>> {
  // Try using FindMeetingTimes which can access other users' availability
  const url = 'https://outlook.office.com/api/v2.0/me/FindMeetingTimes';

  try {
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        Accept: 'application/json',
        Prefer: 'outlook.timezone="Europe/Amsterdam"',
      },
      body: JSON.stringify({
        Attendees: emails.map(email => ({
          EmailAddress: { Address: email, Name: email },
          Type: 'Required',
        })),
        TimeConstraint: {
          Timeslots: [{
            Start: { DateTime: startDateTime, TimeZone: 'W. Europe Standard Time' },
            End: { DateTime: endDateTime, TimeZone: 'W. Europe Standard Time' },
          }],
        },
        MeetingDuration: `PT${durationMinutes}M`,
        ReturnSuggestionReasons: true,
        MinimumAttendeePercentage: 100,
      }),
    });

    if (!response.ok) {
      const errorText = await response.text();
      return {
        ok: false,
        status: response.status,
        error: {
          code: `HTTP_${response.status}`,
          message: errorText || response.statusText,
        },
      };
    }

    const data = await response.json() as {
      MeetingTimeSuggestions?: Array<{
        MeetingTimeSlot: {
          Start: { DateTime: string; TimeZone: string };
          End: { DateTime: string; TimeZone: string };
        };
        Confidence: number;
        AttendeeAvailability?: Array<{
          Attendee: { EmailAddress: { Address: string } };
          Availability: string;
        }>;
      }>;
    };

    // Transform FindMeetingTimes response to our format
    const schedules: ScheduleInfo[] = emails.map(email => ({
      scheduleId: email,
      availabilityView: '',
      scheduleItems: [],
    }));

    // Parse meeting suggestions to find free/busy times
    if (data.MeetingTimeSuggestions && data.MeetingTimeSuggestions.length > 0) {
      // Build free slots from suggestions
      const freeSlots = data.MeetingTimeSuggestions.map(s => ({
        start: s.MeetingTimeSlot.Start.DateTime,
        end: s.MeetingTimeSlot.End.DateTime,
      }));

      for (const schedule of schedules) {
        // Add free slots
        schedule.scheduleItems = freeSlots.map(slot => ({
          status: 'Free',
          start: { dateTime: slot.start, timeZone: 'W. Europe Standard Time' },
          end: { dateTime: slot.end, timeZone: 'W. Europe Standard Time' },
        }));
      }
    } else {
      // No meeting times found - users are busy for the entire period
      for (const schedule of schedules) {
        schedule.scheduleItems = [{
          status: 'Busy',
          start: { dateTime: startDateTime, timeZone: 'W. Europe Standard Time' },
          end: { dateTime: endDateTime, timeZone: 'W. Europe Standard Time' },
          subject: 'No available times',
        }];
      }
    }

    return {
      ok: true,
      status: response.status,
      data: schedules,
    };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

/**
 * Get free/busy info for current user by analyzing their calendar events.
 * Note: Looking up other users requires Microsoft Graph API with different permissions.
 */
export async function getFreeBusy(
  token: string,
  startDateTime: string,
  endDateTime: string
): Promise<OwaResponse<FreeBusySlot[]>> {
  // Use calendar view to get events, then convert to free/busy
  const result = await getCalendarEvents(token, startDateTime, endDateTime);

  if (!result.ok || !result.data) {
    return {
      ok: false,
      status: result.status,
      error: result.error,
    };
  }

  const slots: FreeBusySlot[] = result.data
    .filter(event => !event.IsCancelled)
    .map(event => ({
      status: event.ShowAs === 'Free' ? 'Free' as const :
              event.ShowAs === 'Tentative' ? 'Tentative' as const : 'Busy' as const,
      start: event.Start.DateTime,
      end: event.End.DateTime,
      subject: event.Subject,
    }));

  return {
    ok: true,
    status: 200,
    data: slots,
  };
}

export interface CreateEventOptions {
  token: string;
  subject: string;
  start: string;  // ISO datetime
  end: string;    // ISO datetime
  body?: string;
  location?: string;
  attendees?: Array<{ email: string; name?: string; type?: 'Required' | 'Optional' | 'Resource' }>;
  isOnlineMeeting?: boolean;
}

export interface CreatedEvent {
  Id: string;
  Subject: string;
  Start: { DateTime: string; TimeZone: string };
  End: { DateTime: string; TimeZone: string };
  WebLink?: string;
  OnlineMeetingUrl?: string;
}

/**
 * Create a new calendar event.
 */
export async function createEvent(
  options: CreateEventOptions
): Promise<OwaResponse<CreatedEvent>> {
  const { token, subject, start, end, body, location, attendees, isOnlineMeeting } = options;
  const url = 'https://outlook.office.com/api/v2.0/me/events';

  const eventBody: Record<string, unknown> = {
    Subject: subject,
    Start: {
      DateTime: start,
      TimeZone: 'Europe/Amsterdam',
    },
    End: {
      DateTime: end,
      TimeZone: 'Europe/Amsterdam',
    },
  };

  if (body) {
    eventBody.Body = {
      ContentType: 'Text',
      Content: body,
    };
  }

  if (location) {
    eventBody.Location = {
      DisplayName: location,
    };
  }

  if (attendees && attendees.length > 0) {
    eventBody.Attendees = attendees.map(a => ({
      EmailAddress: {
        Address: a.email,
        Name: a.name || a.email,
      },
      Type: a.type || 'Required',
    }));
  }

  if (isOnlineMeeting) {
    eventBody.IsOnlineMeeting = true;
    eventBody.OnlineMeetingProvider = 'TeamsForBusiness';
  }

  try {
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        'User-Agent': USER_AGENT,
        Accept: 'application/json',
        Prefer: 'outlook.timezone="Europe/Amsterdam"',
      },
      body: JSON.stringify(eventBody),
    });

    if (!response.ok) {
      const errorText = await response.text();
      return {
        ok: false,
        status: response.status,
        error: {
          code: `HTTP_${response.status}`,
          message: errorText || response.statusText,
        },
      };
    }

    const data = (await response.json()) as CreatedEvent;
    return { ok: true, status: response.status, data };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

export interface UpdateEventOptions {
  token: string;
  eventId: string;
  subject?: string;
  start?: string;  // ISO datetime
  end?: string;    // ISO datetime
  body?: string;
  location?: string;
  attendees?: Array<{ email: string; name?: string; type?: 'Required' | 'Optional' | 'Resource' }>;
  isOnlineMeeting?: boolean;
}

/**
 * Update an existing calendar event.
 */
export async function updateEvent(
  options: UpdateEventOptions
): Promise<OwaResponse<CreatedEvent>> {
  const { token, eventId, subject, start, end, body, location, attendees, isOnlineMeeting } = options;
  const url = `https://outlook.office.com/api/v2.0/me/events/${encodeURIComponent(eventId)}`;

  const eventBody: Record<string, unknown> = {};

  if (subject !== undefined) {
    eventBody.Subject = subject;
  }

  if (start !== undefined) {
    eventBody.Start = {
      DateTime: start,
      TimeZone: 'Europe/Amsterdam',
    };
  }

  if (end !== undefined) {
    eventBody.End = {
      DateTime: end,
      TimeZone: 'Europe/Amsterdam',
    };
  }

  if (body !== undefined) {
    eventBody.Body = {
      ContentType: 'Text',
      Content: body,
    };
  }

  if (location !== undefined) {
    eventBody.Location = {
      DisplayName: location,
    };
  }

  if (attendees !== undefined) {
    eventBody.Attendees = attendees.map(a => ({
      EmailAddress: {
        Address: a.email,
        Name: a.name || a.email,
      },
      Type: a.type || 'Required',
    }));
  }

  if (isOnlineMeeting !== undefined) {
    eventBody.IsOnlineMeeting = isOnlineMeeting;
    if (isOnlineMeeting) {
      eventBody.OnlineMeetingProvider = 'TeamsForBusiness';
    }
  }

  try {
    const response = await fetch(url, {
      method: 'PATCH',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        'User-Agent': USER_AGENT,
        Accept: 'application/json',
        Prefer: 'outlook.timezone="Europe/Amsterdam"',
      },
      body: JSON.stringify(eventBody),
    });

    if (!response.ok) {
      const errorText = await response.text();
      return {
        ok: false,
        status: response.status,
        error: {
          code: `HTTP_${response.status}`,
          message: errorText || response.statusText,
        },
      };
    }

    const data = (await response.json()) as CreatedEvent;
    return { ok: true, status: response.status, data };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

export interface Room {
  Address: string;
  Name: string;
}

export interface RoomList {
  Address: string;
  Name: string;
}

/**
 * Get available room lists (buildings/locations).
 */
export async function getRoomLists(
  token: string
): Promise<OwaResponse<RoomList[]>> {
  // Try Graph API first (works with Outlook token in some cases)
  const urls = [
    'https://graph.microsoft.com/v1.0/places/microsoft.graph.roomList',
    'https://outlook.office.com/api/v2.0/me/findRoomLists',
  ];

  for (const url of urls) {
    try {
      const response = await fetch(url, {
        headers: {
          Authorization: `Bearer ${token}`,
          'User-Agent': USER_AGENT,
          Accept: 'application/json',
        },
      });

      if (response.ok) {
        const data = (await response.json()) as { value: Array<{ emailAddress?: string; address?: string; displayName?: string; Name?: string; Address?: string }> };
        const rooms = data.value.map(r => ({
          Address: r.emailAddress || r.address || r.Address || '',
          Name: r.displayName || r.Name || '',
        }));
        if (rooms.length > 0) {
          return { ok: true, status: response.status, data: rooms };
        }
      }
    } catch {
      // Try next URL
    }
  }

  return {
    ok: false,
    status: 404,
    error: {
      code: 'NOT_FOUND',
      message: 'No room lists found',
    },
  };
}

/**
 * Get rooms in a room list or all rooms.
 */
export async function getRooms(
  token: string,
  roomListAddress?: string
): Promise<OwaResponse<Room[]>> {
  // Try multiple endpoints
  const urls = roomListAddress
    ? [
        `https://graph.microsoft.com/v1.0/places/${encodeURIComponent(roomListAddress)}/microsoft.graph.roomList/rooms`,
        `https://outlook.office.com/api/v2.0/me/findRooms(RoomList='${encodeURIComponent(roomListAddress)}')`,
      ]
    : [
        'https://graph.microsoft.com/v1.0/places/microsoft.graph.room',
        'https://outlook.office.com/api/v2.0/me/findRooms',
      ];

  for (const url of urls) {
    try {
      const response = await fetch(url, {
        headers: {
          Authorization: `Bearer ${token}`,
          'User-Agent': USER_AGENT,
          Accept: 'application/json',
        },
      });

      if (response.ok) {
        const data = (await response.json()) as { value: Array<{ emailAddress?: string; address?: string; displayName?: string; Name?: string; Address?: string }> };
        const rooms = data.value.map(r => ({
          Address: r.emailAddress || r.address || r.Address || '',
          Name: r.displayName || r.Name || '',
        }));
        if (rooms.length > 0) {
          return { ok: true, status: response.status, data: rooms };
        }
      }
    } catch {
      // Try next URL
    }
  }

  return {
    ok: false,
    status: 404,
    error: {
      code: 'NOT_FOUND',
      message: 'No rooms found',
    },
  };
}

/**
 * Search for rooms/resources using the People API.
 */
export async function searchRooms(
  token: string,
  query: string = 'room'
): Promise<OwaResponse<Room[]>> {
  // Use People search API with room filter
  const searchQuery = query || 'room';
  const url = `https://outlook.office.com/api/v2.0/me/people?$search=${encodeURIComponent(searchQuery)}&$top=50`;

  try {
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        'User-Agent': USER_AGENT,
        Accept: 'application/json',
      },
    });

    if (!response.ok) {
      return {
        ok: false,
        status: response.status,
        error: {
          code: `HTTP_${response.status}`,
          message: response.statusText,
        },
      };
    }

    const data = (await response.json()) as {
      value: Array<{
        DisplayName?: string;
        ScoredEmailAddresses?: Array<{ Address?: string }>;
        PersonType?: { Class?: string; Subclass?: string };
      }>;
    };

    // Filter to only rooms (PersonType.Subclass === 'Room')
    const rooms: Room[] = data.value
      .filter(p => p.PersonType?.Subclass === 'Room')
      .map(p => ({
        Name: p.DisplayName || '',
        Address: p.ScoredEmailAddresses?.[0]?.Address || '',
      }))
      .filter(r => r.Address);

    return { ok: true, status: response.status, data: rooms };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

/**
 * Delete a calendar event.
 */
export async function deleteEvent(
  token: string,
  eventId: string
): Promise<OwaResponse<void>> {
  const url = `https://outlook.office.com/api/v2.0/me/events/${encodeURIComponent(eventId)}`;

  try {
    const response = await fetch(url, {
      method: 'DELETE',
      headers: {
        Authorization: `Bearer ${token}`,
        'User-Agent': USER_AGENT,
      },
    });

    if (!response.ok) {
      const errorText = await response.text();
      return {
        ok: false,
        status: response.status,
        error: {
          code: `HTTP_${response.status}`,
          message: errorText || response.statusText,
        },
      };
    }

    return { ok: true, status: response.status };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

// Email types
export interface EmailAddress {
  Name?: string;
  Address?: string;
}

export interface EmailMessage {
  Id: string;
  Subject?: string;
  BodyPreview?: string;
  Body?: {
    ContentType: string;
    Content: string;
  };
  From?: {
    EmailAddress?: EmailAddress;
  };
  ToRecipients?: Array<{ EmailAddress?: EmailAddress }>;
  CcRecipients?: Array<{ EmailAddress?: EmailAddress }>;
  ReceivedDateTime?: string;
  SentDateTime?: string;
  IsRead?: boolean;
  IsDraft?: boolean;
  HasAttachments?: boolean;
  Importance?: 'Low' | 'Normal' | 'High';
  Flag?: {
    FlagStatus?: 'NotFlagged' | 'Flagged' | 'Complete';
  };
}

export interface EmailListResponse {
  value: EmailMessage[];
  '@odata.nextLink'?: string;
}

export interface GetEmailsOptions {
  token: string;
  folder?: string;  // inbox, sentitems, drafts, deleteditems, archive, junkemail
  top?: number;
  skip?: number;
  filter?: string;
  search?: string;
  select?: string[];
  orderBy?: string;
}

/**
 * Get emails from a folder.
 */
export async function getEmails(
  options: GetEmailsOptions
): Promise<OwaResponse<EmailListResponse>> {
  const {
    token,
    folder = 'inbox',
    top = 10,
    skip = 0,
    filter,
    search,
    select = ['Id', 'Subject', 'BodyPreview', 'From', 'ReceivedDateTime', 'IsRead', 'HasAttachments', 'Importance', 'Flag'],
    orderBy = 'ReceivedDateTime desc',
  } = options;

  const params = new URLSearchParams();
  params.set('$top', top.toString());
  if (skip > 0) params.set('$skip', skip.toString());
  if (filter) params.set('$filter', filter);
  if (search) params.set('$search', `"${search}"`);
  params.set('$select', select.join(','));
  // Note: $orderby is ignored when $search is used (results are ranked by relevance)
  if (!search) params.set('$orderby', orderBy);

  const url = `https://outlook.office.com/api/v2.0/me/mailfolders/${folder}/messages?${params}`;

  try {
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        'User-Agent': USER_AGENT,
        Accept: 'application/json',
        Prefer: 'outlook.body-content-type="text"',
      },
    });

    if (!response.ok) {
      const errorText = await response.text();
      return {
        ok: false,
        status: response.status,
        error: {
          code: `HTTP_${response.status}`,
          message: errorText || response.statusText,
        },
      };
    }

    const data = (await response.json()) as EmailListResponse;
    return { ok: true, status: response.status, data };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

/**
 * Get a single email by ID.
 */
export async function getEmail(
  token: string,
  messageId: string
): Promise<OwaResponse<EmailMessage>> {
  const url = `https://outlook.office.com/api/v2.0/me/messages/${encodeURIComponent(messageId)}`;

  try {
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        'User-Agent': USER_AGENT,
        Accept: 'application/json',
        Prefer: 'outlook.body-content-type="text"',
      },
    });

    if (!response.ok) {
      return {
        ok: false,
        status: response.status,
        error: {
          code: `HTTP_${response.status}`,
          message: response.statusText,
        },
      };
    }

    const data = (await response.json()) as EmailMessage;
    return { ok: true, status: response.status, data };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

// Attachment types
export interface Attachment {
  Id: string;
  Name: string;
  ContentType: string;
  Size: number;
  IsInline: boolean;
  ContentId?: string;
  ContentBytes?: string;  // Base64 encoded content
}

export interface AttachmentListResponse {
  value: Attachment[];
}

/**
 * Get list of attachments for an email.
 */
export async function getAttachments(
  token: string,
  messageId: string
): Promise<OwaResponse<AttachmentListResponse>> {
  const url = `https://outlook.office.com/api/v2.0/me/messages/${encodeURIComponent(messageId)}/attachments`;

  try {
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        'User-Agent': USER_AGENT,
        Accept: 'application/json',
      },
    });

    if (!response.ok) {
      return {
        ok: false,
        status: response.status,
        error: {
          code: `HTTP_${response.status}`,
          message: response.statusText,
        },
      };
    }

    const data = (await response.json()) as AttachmentListResponse;
    return { ok: true, status: response.status, data };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

/**
 * Get a single attachment with content.
 */
export async function getAttachment(
  token: string,
  messageId: string,
  attachmentId: string
): Promise<OwaResponse<Attachment>> {
  const url = `https://outlook.office.com/api/v2.0/me/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachmentId)}`;

  try {
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        'User-Agent': USER_AGENT,
        Accept: 'application/json',
      },
    });

    if (!response.ok) {
      return {
        ok: false,
        status: response.status,
        error: {
          code: `HTTP_${response.status}`,
          message: response.statusText,
        },
      };
    }

    const data = (await response.json()) as Attachment;
    return { ok: true, status: response.status, data };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

export type ResponseType = 'accept' | 'decline' | 'tentative';

export interface RespondToEventOptions {
  token: string;
  eventId: string;
  response: ResponseType;
  comment?: string;
  sendResponse?: boolean;
}

/**
 * Respond to a calendar event (accept, decline, or tentatively accept).
 */
export async function respondToEvent(
  options: RespondToEventOptions
): Promise<OwaResponse<void>> {
  const { token, eventId, response, comment, sendResponse = true } = options;

  const actionMap: Record<ResponseType, string> = {
    accept: 'accept',
    decline: 'decline',
    tentative: 'tentativelyAccept',
  };

  const action = actionMap[response];
  const url = `https://outlook.office.com/api/v2.0/me/events/${encodeURIComponent(eventId)}/${action}`;

  try {
    const body: Record<string, unknown> = {
      SendResponse: sendResponse,
    };

    if (comment) {
      body.Comment = comment;
    }

    const res = await fetch(url, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        'User-Agent': USER_AGENT,
        Accept: 'application/json',
      },
      body: JSON.stringify(body),
    });

    if (!res.ok) {
      const errorText = await res.text();
      return {
        ok: false,
        status: res.status,
        error: {
          code: `HTTP_${res.status}`,
          message: errorText || res.statusText,
        },
      };
    }

    return { ok: true, status: res.status };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

/**
 * Get a single calendar event by ID.
 */
export async function getCalendarEvent(
  token: string,
  eventId: string
): Promise<OwaResponse<CalendarEvent>> {
  const url = `https://outlook.office.com/api/v2.0/me/events/${encodeURIComponent(eventId)}`;

  try {
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        'User-Agent': USER_AGENT,
        Accept: 'application/json',
        Prefer: 'outlook.timezone="Europe/Amsterdam"',
      },
    });

    if (!response.ok) {
      return {
        ok: false,
        status: response.status,
        error: {
          code: `HTTP_${response.status}`,
          message: response.statusText,
        },
      };
    }

    const data = (await response.json()) as CalendarEvent;
    return { ok: true, status: response.status, data };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: {
        code: 'NETWORK_ERROR',
        message: err instanceof Error ? err.message : 'Unknown error',
      },
    };
  }
}

export async function validateSession(token: string): Promise<boolean> {
  // Use Outlook REST API to validate the token
  const url = 'https://outlook.office.com/api/v2.0/me/mailfolders/inbox';

  try {
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        'User-Agent': USER_AGENT,
      },
    });

    // 200 means valid session, 401/403 means expired/invalid
    return response.ok;
  } catch {
    return false;
  }
}
