import { Command } from 'commander';
import { resolveAuth } from '../lib/auth.js';
import { getFreeBusy, getScheduleForUsers, getScheduleViaOutlook, type FreeBusySlot } from '../lib/owa-client.js';

function formatTime(dateStr: string): string {
  const date = new Date(dateStr);
  return date.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit', hour12: false });
}

function formatDate(dateStr: string): string {
  const date = new Date(dateStr);
  return date.toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric' });
}

function getStatusIcon(status: string): string {
  switch (status) {
    case 'Free': return '🟢';
    case 'Tentative': return '🟡';
    case 'Busy': return '🔴';
    default: return '⚪';
  }
}

function parseDay(day: string, baseDate: Date = new Date()): Date {
  const now = new Date(baseDate);

  switch (day.toLowerCase()) {
    case 'today':
      return now;
    case 'tomorrow':
      now.setDate(now.getDate() + 1);
      return now;
    case 'monday':
    case 'tuesday':
    case 'wednesday':
    case 'thursday':
    case 'friday':
    case 'saturday':
    case 'sunday': {
      const days = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];
      const targetDay = days.indexOf(day.toLowerCase());
      const currentDay = now.getDay();
      let diff = targetDay - currentDay;
      if (diff <= 0) diff += 7; // Go to next occurrence
      now.setDate(now.getDate() + diff);
      return now;
    }
    default:
      const parsed = new Date(day);
      return isNaN(parsed.getTime()) ? now : parsed;
  }
}

function getDateRange(startDay: string, endDay?: string): { start: Date; end: Date; label: string } {
  const now = new Date();

  // Handle week keywords
  switch (startDay.toLowerCase()) {
    case 'week':
    case 'thisweek': {
      const start = new Date(now);
      const dayOfWeek = start.getDay();
      const diff = dayOfWeek === 0 ? 1 : 1 - dayOfWeek + 7; // Next Monday
      start.setDate(start.getDate() + (dayOfWeek === 0 ? 1 : 8 - dayOfWeek));
      start.setHours(0, 0, 0, 0);
      const end = new Date(start);
      end.setDate(end.getDate() + 4); // Friday
      end.setHours(23, 59, 59, 999);
      return { start, end, label: 'This Week (Mon-Fri)' };
    }
    case 'nextweek': {
      const start = new Date(now);
      const dayOfWeek = start.getDay();
      const daysUntilNextMonday = dayOfWeek === 0 ? 1 : 8 - dayOfWeek;
      start.setDate(start.getDate() + daysUntilNextMonday);
      start.setHours(0, 0, 0, 0);
      const end = new Date(start);
      end.setDate(end.getDate() + 4); // Friday
      end.setHours(23, 59, 59, 999);
      return { start, end, label: 'Next Week (Mon-Fri)' };
    }
  }

  const startDate = parseDay(startDay);
  startDate.setHours(0, 0, 0, 0);

  if (endDay) {
    const endDate = parseDay(endDay, startDate);
    endDate.setHours(23, 59, 59, 999);
    return {
      start: startDate,
      end: endDate,
      label: `${formatDate(startDate.toISOString())} - ${formatDate(endDate.toISOString())}`
    };
  }

  const endDate = new Date(startDate);
  endDate.setHours(23, 59, 59, 999);

  return { start: startDate, end: endDate, label: formatDate(startDate.toISOString()) };
}

function findFreeSlots(
  slots: FreeBusySlot[],
  dayStart: Date,
  dayEnd: Date,
  workStart: number = 9,
  workEnd: number = 17
): { start: Date; end: Date }[] {
  // Set working hours
  const workingStart = new Date(dayStart);
  workingStart.setHours(workStart, 0, 0, 0);
  const workingEnd = new Date(dayStart);
  workingEnd.setHours(workEnd, 0, 0, 0);

  // Filter to busy slots only and sort
  const busySlots = slots
    .filter(s => s.status !== 'Free')
    .map(s => ({
      start: new Date(s.start),
      end: new Date(s.end),
    }))
    .sort((a, b) => a.start.getTime() - b.start.getTime());

  const freeSlots: { start: Date; end: Date }[] = [];
  let current = workingStart;

  for (const busy of busySlots) {
    // Skip if busy slot is outside working hours
    if (busy.end <= workingStart || busy.start >= workingEnd) continue;

    // Clamp to working hours
    const busyStart = busy.start < workingStart ? workingStart : busy.start;
    const busyEnd = busy.end > workingEnd ? workingEnd : busy.end;

    if (busyStart > current) {
      freeSlots.push({ start: new Date(current), end: new Date(busyStart) });
    }
    current = busyEnd > current ? busyEnd : current;
  }

  // Add remaining time until end of working hours
  if (current < workingEnd) {
    freeSlots.push({ start: new Date(current), end: new Date(workingEnd) });
  }

  return freeSlots;
}

export const freebusyCommand = new Command('freebusy')
  .description('Check free/busy status for yourself or other users')
  .argument('[start]', 'Start day: today, tomorrow, monday-sunday, week, nextweek, or YYYY-MM-DD', 'today')
  .argument('[endOrEmails...]', 'End day for range OR email addresses to check')
  .option('--start <hour>', 'Work day start hour (0-23)', '9')
  .option('--end <hour>', 'Work day end hour (0-23)', '17')
  .option('--free', 'Show free slots instead of busy')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('-i, --interactive', 'Open browser to extract token automatically')
  .action(async (startDay: string, endOrEmails: string[], options: {
    start: string;
    end: string;
    free?: boolean;
    json?: boolean;
    token?: string;
    interactive?: boolean
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

    // Parse arguments: figure out which are dates vs emails
    const dateKeywords = ['today', 'tomorrow', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday', 'week', 'thisweek', 'nextweek'];
    const isDateArg = (arg: string) => {
      if (dateKeywords.includes(arg.toLowerCase())) return true;
      if (/^\d{4}-\d{2}-\d{2}$/.test(arg)) return true;
      return false;
    };

    let endDay: string | undefined;
    let emails: string[] = [];

    for (const arg of endOrEmails) {
      if (arg.includes('@')) {
        emails.push(arg);
      } else if (isDateArg(arg) && !endDay) {
        endDay = arg;
      } else {
        emails.push(arg); // Assume it's an email if not a date
      }
    }

    const { start, end, label } = getDateRange(startDay, endDay);
    const workStart = parseInt(options.start);
    const workEnd = parseInt(options.end);

    // If emails provided, look up other users
    if (emails.length > 0) {
      // Try Graph API first if we have a token, otherwise use Outlook API
      let result;
      if (authResult.graphToken) {
        result = await getScheduleForUsers(
          authResult.graphToken,
          emails,
          start.toISOString(),
          end.toISOString()
        );
      }

      // Fall back to Outlook API if Graph failed or no token
      if (!result?.ok) {
        result = await getScheduleViaOutlook(
          authResult.token!,
          emails,
          start.toISOString(),
          end.toISOString()
        );
      }

      if (!result.ok || !result.data) {
        if (options.json) {
          console.log(JSON.stringify({ error: result.error?.message || 'Failed to fetch schedule' }, null, 2));
        } else {
          console.error(`Error: ${result.error?.message || 'Failed to fetch schedule'}`);
        }
        process.exit(1);
      }

      if (options.json) {
        console.log(JSON.stringify(result.data, null, 2));
        return;
      }

      console.log(`\n📊 Availability for ${label}`);
      console.log('─'.repeat(50));

      for (const schedule of result.data) {
        console.log(`\n  ${schedule.scheduleId}`);

        if (schedule.scheduleItems && schedule.scheduleItems.length > 0) {
          // Find busy times by looking at gaps in free slots
          const busySlots = findBusyFromFreeSlots(schedule.scheduleItems, start, workStart, workEnd);

          if (busySlots.length === 0) {
            console.log('    🟢 Free during working hours');
          } else {
            for (const item of busySlots) {
              const icon = getStatusIcon(item.status);
              console.log(`    ${icon} ${item.startTime} - ${item.endTime}: ${item.status}`);
            }
          }
        } else if (schedule.availabilityView) {
          // Parse availability view string (0=free, 1=tentative, 2=busy, 3=oof, 4=working elsewhere)
          const busyBlocks = parseAvailabilityView(schedule.availabilityView, start, 30);
          if (busyBlocks.length === 0) {
            console.log('    🟢 Free all day');
          } else {
            for (const block of busyBlocks) {
              const icon = getStatusIcon(block.status);
              console.log(`    ${icon} ${formatTime(block.start.toISOString())} - ${formatTime(block.end.toISOString())}: ${block.status}`);
            }
          }
        } else {
          console.log('    🟢 Free all day');
        }
      }
      console.log();
      return;
    }

    // Default: get own schedule using Outlook API
    const result = await getFreeBusy(authResult.token!, start.toISOString(), end.toISOString());

    if (!result.ok || !result.data) {
      if (options.json) {
        console.log(JSON.stringify({ error: result.error?.message || 'Failed to fetch schedule' }, null, 2));
      } else {
        console.error(`Error: ${result.error?.message || 'Failed to fetch schedule'}`);
      }
      process.exit(1);
    }

    if (options.json) {
      if (options.free) {
        const freeSlots = findFreeSlots(result.data, start, end, workStart, workEnd);
        console.log(JSON.stringify(freeSlots.map(s => ({
          start: s.start.toISOString(),
          end: s.end.toISOString(),
        })), null, 2));
      } else {
        console.log(JSON.stringify(result.data, null, 2));
      }
      return;
    }

    console.log(`\n📊 ${options.free ? 'Free times' : 'Busy times'} for ${label}`);
    console.log('─'.repeat(40));

    if (options.free) {
      const freeSlots = findFreeSlots(result.data, start, end, workStart, workEnd);
      if (freeSlots.length === 0) {
        console.log('  No free time during working hours.');
      } else {
        for (const slot of freeSlots) {
          const duration = Math.round((slot.end.getTime() - slot.start.getTime()) / 60000);
          const hours = Math.floor(duration / 60);
          const mins = duration % 60;
          const durationStr = hours > 0 ? `${hours}h${mins > 0 ? ` ${mins}m` : ''}` : `${mins}m`;
          console.log(`  🟢 ${formatTime(slot.start.toISOString())} - ${formatTime(slot.end.toISOString())} (${durationStr})`);
        }
      }
    } else {
      const busySlots = result.data.filter(s => s.status !== 'Free');
      if (busySlots.length === 0) {
        console.log('  🟢 All day free!');
      } else {
        for (const slot of busySlots) {
          const icon = getStatusIcon(slot.status);
          console.log(`  ${icon} ${formatTime(slot.start)} - ${formatTime(slot.end)}: ${slot.subject || slot.status}`);
        }
      }
    }
    console.log();
  });

interface ScheduleItem {
  status: string;
  start: { dateTime: string; timeZone: string };
  end: { dateTime: string; timeZone: string };
  subject?: string;
}

function findBusyFromFreeSlots(
  freeItems: ScheduleItem[],
  dayStart: Date,
  workStartHour: number,
  workEndHour: number
): { startTime: string; endTime: string; status: string }[] {
  // Set working hours bounds
  const workStart = new Date(dayStart);
  workStart.setHours(workStartHour, 0, 0, 0);
  const workEnd = new Date(dayStart);
  workEnd.setHours(workEndHour, 0, 0, 0);

  // Filter free items to working hours and sort
  const freeSlots = freeItems
    .filter(item => item.status === 'Free')
    .map(item => ({
      start: new Date(item.start.dateTime),
      end: new Date(item.end.dateTime),
    }))
    .filter(slot => slot.end > workStart && slot.start < workEnd)
    .sort((a, b) => a.start.getTime() - b.start.getTime());

  // Find gaps (busy times) between free slots
  const busySlots: { startTime: string; endTime: string; status: string }[] = [];
  let current = workStart;

  for (const free of freeSlots) {
    // Clamp free slot to working hours
    const freeStart = free.start < workStart ? workStart : free.start;
    const freeEnd = free.end > workEnd ? workEnd : free.end;

    // If there's a gap before this free slot, it's busy
    if (freeStart > current) {
      busySlots.push({
        startTime: formatTime(current.toISOString()),
        endTime: formatTime(freeStart.toISOString()),
        status: 'Busy',
      });
    }
    current = freeEnd > current ? freeEnd : current;
  }

  // Check for busy time after last free slot
  if (current < workEnd) {
    busySlots.push({
      startTime: formatTime(current.toISOString()),
      endTime: formatTime(workEnd.toISOString()),
      status: 'Busy',
    });
  }

  return busySlots;
}

function parseAvailabilityView(view: string, startDate: Date, intervalMinutes: number): { start: Date; end: Date; status: string }[] {
  const statusMap: Record<string, string> = {
    '0': 'Free',
    '1': 'Tentative',
    '2': 'Busy',
    '3': 'Out of Office',
    '4': 'Working Elsewhere',
  };

  const blocks: { start: Date; end: Date; status: string }[] = [];
  let current: { start: Date; end: Date; status: string } | null = null;

  for (let i = 0; i < view.length; i++) {
    const time = new Date(startDate.getTime() + i * intervalMinutes * 60000);
    const status = statusMap[view[i]] || 'Unknown';

    if (status === 'Free') {
      if (current) {
        blocks.push(current);
        current = null;
      }
      continue;
    }

    if (!current || current.status !== status) {
      if (current) blocks.push(current);
      current = { start: time, end: new Date(time.getTime() + intervalMinutes * 60000), status };
    } else {
      current.end = new Date(time.getTime() + intervalMinutes * 60000);
    }
  }

  if (current) blocks.push(current);
  return blocks;
}
