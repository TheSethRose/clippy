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
