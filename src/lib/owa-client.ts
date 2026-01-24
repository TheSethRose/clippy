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
