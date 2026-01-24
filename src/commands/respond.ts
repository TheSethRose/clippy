import { Command } from 'commander';
import { resolveAuth } from '../lib/auth.js';
import { getCalendarEvents, respondToEvent, getOwaUserInfo, ResponseType } from '../lib/owa-client.js';

function formatTime(dateStr: string): string {
  const date = new Date(dateStr);
  return date.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit', hour12: false });
}

function formatDate(dateStr: string): string {
  const date = new Date(dateStr);
  return date.toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric' });
}

function getResponseIcon(response: string): string {
  switch (response) {
    case 'Accepted': return '\u2713';
    case 'Declined': return '\u2717';
    case 'TentativelyAccepted': return '?';
    case 'None':
    case 'NotResponded': return '\u2022';
    default: return ' ';
  }
}

export const respondCommand = new Command('respond')
  .description('Respond to calendar invitations (accept/decline/tentative)')
  .argument('[action]', 'Action: list, accept, decline, tentative')
  .argument('[eventIndex]', 'Event index from the list (1-based)')
  .option('--comment <text>', 'Add a comment to your response')
  .option('--no-notify', 'Don\'t send response to organizer')
  .option('--include-optional', 'Include optional invitations (default)', true)
  .option('--only-required', 'Only show required invitations')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('-i, --interactive', 'Open browser to extract token automatically')
  .action(async (action: string | undefined, eventIndex: string | undefined, options: {
    comment?: string;
    notify: boolean;
    json?: boolean;
    token?: string;
    interactive?: boolean;
  }) => {
    const authResult = await resolveAuth({
      token: options.token,
      interactive: options.interactive,
    });

    if (!authResult.success) {
      if (options.json) {
        console.log(JSON.stringify({ error: authResult.error }, null, 2));
      } else {
        console.error(`Error: ${authResult.error}`);
        console.error('\nRun `clippy login --interactive` to authenticate.');
      }
      process.exit(1);
    }

    // Get user's email to identify their response status
    const userInfo = await getOwaUserInfo(authResult.token!);
    const userEmail = userInfo.data?.email?.toLowerCase();

    // Fetch upcoming events
    const now = new Date();
    const futureDate = new Date(now);
    futureDate.setDate(futureDate.getDate() + 30); // Look 30 days ahead

    const result = await getCalendarEvents(
      authResult.token!,
      now.toISOString(),
      futureDate.toISOString()
    );

    if (!result.ok || !result.data) {
      if (options.json) {
        console.log(JSON.stringify({ error: result.error?.message || 'Failed to fetch events' }, null, 2));
      } else {
        console.error(`Error: ${result.error?.message || 'Failed to fetch events'}`);
      }
      process.exit(1);
    }

    // Filter to events where user is an attendee (and not organizer)
    const pendingEvents = result.data.filter(event => {
      if (event.IsCancelled) return false;
      if (event.IsOrganizer) return false;

      // Find user's attendance record
      const myAttendance = event.Attendees?.find(
        a => a.EmailAddress?.Address?.toLowerCase() === userEmail
      );

      if (!myAttendance) return false;

      // Optional attendance handling
      const isOptional = (myAttendance as any).AttendeeType === 'Optional';
      if (options.onlyRequired && isOptional) return false;

      // Show all attendee events by default; include status in list
      return true;
    });

    // Default action is 'list'
    const actionLower = (action || 'list').toLowerCase();

    if (actionLower === 'list') {
      if (options.json) {
        console.log(JSON.stringify({
          pendingEvents: pendingEvents.map((e, i) => ({
            index: i + 1,
            id: e.Id,
            subject: e.Subject,
            start: e.Start.DateTime,
            end: e.End.DateTime,
            organizer: e.Organizer?.EmailAddress?.Name || e.Organizer?.EmailAddress?.Address,
            location: e.Location?.DisplayName,
          })),
        }, null, 2));
        return;
      }

      console.log('\nCalendar invitations awaiting your response:\n');
      console.log('\u2500'.repeat(60));

      if (pendingEvents.length === 0) {
        console.log('\n  No pending invitations found.\n');
        return;
      }

      for (let i = 0; i < pendingEvents.length; i++) {
        const event = pendingEvents[i];
        const dateStr = formatDate(event.Start.DateTime);
        const startTime = formatTime(event.Start.DateTime);
        const endTime = formatTime(event.End.DateTime);

        const myAttendance = event.Attendees?.find(
          a => a.EmailAddress?.Address?.toLowerCase() === userEmail
        );
        const response = myAttendance?.Status?.Response || 'None';
        const icon = getResponseIcon(response);

        console.log(`\n  [${i + 1}] ${icon} ${event.Subject}`);
        console.log(`      ${dateStr} ${startTime} - ${endTime}`);
        if (event.Location?.DisplayName) {
          console.log(`      Location: ${event.Location.DisplayName}`);
        }
        if (event.Organizer?.EmailAddress) {
          const org = event.Organizer.EmailAddress;
          console.log(`      Organizer: ${org.Name || org.Address}`);
        }
      }

      console.log('\n' + '\u2500'.repeat(60));
      console.log('\nTo respond, use:');
      console.log('  clippy respond accept <number>');
      console.log('  clippy respond decline <number>');
      console.log('  clippy respond tentative <number>');
      console.log('');
      return;
    }

    // Handle accept/decline/tentative
    if (!['accept', 'decline', 'tentative'].includes(actionLower)) {
      console.error(`Unknown action: ${action}`);
      console.error('Valid actions: list, accept, decline, tentative');
      process.exit(1);
    }

    if (!eventIndex) {
      console.error('Please specify the event number to respond to.');
      console.error('Run `clippy respond list` to see pending invitations.');
      process.exit(1);
    }

    const idx = parseInt(eventIndex) - 1;
    if (isNaN(idx) || idx < 0 || idx >= pendingEvents.length) {
      console.error(`Invalid event number: ${eventIndex}`);
      console.error(`Valid range: 1-${pendingEvents.length}`);
      process.exit(1);
    }

    const targetEvent = pendingEvents[idx];

    console.log(`\nResponding to: ${targetEvent.Subject}`);
    console.log(`  ${formatDate(targetEvent.Start.DateTime)} ${formatTime(targetEvent.Start.DateTime)} - ${formatTime(targetEvent.End.DateTime)}`);
    console.log(`  Action: ${actionLower}`);
    if (options.comment) {
      console.log(`  Comment: ${options.comment}`);
    }
    console.log('');

    const response = await respondToEvent({
      token: authResult.token!,
      eventId: targetEvent.Id,
      response: actionLower as ResponseType,
      comment: options.comment,
      sendResponse: options.notify,
    });

    if (!response.ok) {
      if (options.json) {
        console.log(JSON.stringify({ error: response.error?.message || 'Failed to respond' }, null, 2));
      } else {
        console.error(`Error: ${response.error?.message || 'Failed to respond'}`);
      }
      process.exit(1);
    }

    const actionPast = actionLower === 'tentative' ? 'tentatively accepted' : `${actionLower}d`;
    if (options.json) {
      console.log(JSON.stringify({ success: true, action: actionLower }, null, 2));
    } else {
      console.log(`\u2713 Successfully ${actionPast} the invitation.`);
    }
  });
