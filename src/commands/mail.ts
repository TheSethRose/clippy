import { Command } from 'commander';
import { resolveAuth } from '../lib/auth.js';
import { getEmails, getEmail } from '../lib/owa-client.js';

function formatDate(dateStr: string): string {
  const date = new Date(dateStr);
  const now = new Date();
  const isToday = date.toDateString() === now.toDateString();
  const yesterday = new Date(now);
  yesterday.setDate(yesterday.getDate() - 1);
  const isYesterday = date.toDateString() === yesterday.toDateString();

  if (isToday) {
    return date.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit', hour12: false });
  } else if (isYesterday) {
    return 'Yesterday';
  } else {
    return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
  }
}

function truncate(str: string, maxLen: number): string {
  if (!str) return '';
  str = str.replace(/\s+/g, ' ').trim();
  if (str.length <= maxLen) return str;
  return str.substring(0, maxLen - 1) + '\u2026';
}

export const mailCommand = new Command('mail')
  .description('List and read emails')
  .argument('[folder]', 'Folder: inbox, sent, drafts, deleted, archive, junk', 'inbox')
  .option('-n, --limit <number>', 'Number of emails to show', '10')
  .option('-p, --page <number>', 'Page number (1-based)', '1')
  .option('--unread', 'Show only unread emails')
  .option('--flagged', 'Show only flagged emails')
  .option('-s, --search <query>', 'Search emails (subject, body, sender)')
  .option('-r, --read <index>', 'Read email at index (1-based)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('-i, --interactive', 'Open browser to extract token automatically')
  .action(async (folder: string, options: {
    limit: string;
    page: string;
    unread?: boolean;
    flagged?: boolean;
    search?: string;
    read?: string;
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

    // Map folder names to API folder IDs
    const folderMap: Record<string, string> = {
      inbox: 'inbox',
      sent: 'sentitems',
      sentitems: 'sentitems',
      drafts: 'drafts',
      deleted: 'deleteditems',
      deleteditems: 'deleteditems',
      trash: 'deleteditems',
      archive: 'archive',
      junk: 'junkemail',
      junkemail: 'junkemail',
      spam: 'junkemail',
    };

    const apiFolder = folderMap[folder.toLowerCase()] || folder;
    const limit = parseInt(options.limit) || 10;
    const page = parseInt(options.page) || 1;
    const skip = (page - 1) * limit;

    // Build filter
    const filters: string[] = [];
    if (options.unread) {
      filters.push('IsRead eq false');
    }
    if (options.flagged) {
      filters.push("Flag/FlagStatus eq 'Flagged'");
    }

    const result = await getEmails({
      token: authResult.token!,
      folder: apiFolder,
      top: limit,
      skip,
      filter: filters.length > 0 ? filters.join(' and ') : undefined,
      search: options.search,
    });

    if (!result.ok || !result.data) {
      if (options.json) {
        console.log(JSON.stringify({ error: result.error?.message || 'Failed to fetch emails' }, null, 2));
      } else {
        console.error(`Error: ${result.error?.message || 'Failed to fetch emails'}`);
      }
      process.exit(1);
    }

    const emails = result.data.value;

    // Handle reading a specific email
    if (options.read) {
      const idx = parseInt(options.read) - 1;
      if (isNaN(idx) || idx < 0 || idx >= emails.length) {
        console.error(`Invalid email number: ${options.read}`);
        console.error(`Valid range: 1-${emails.length}`);
        process.exit(1);
      }

      const emailSummary = emails[idx];
      const fullEmail = await getEmail(authResult.token!, emailSummary.Id);

      if (!fullEmail.ok || !fullEmail.data) {
        console.error(`Error: ${fullEmail.error?.message || 'Failed to fetch email'}`);
        process.exit(1);
      }

      const email = fullEmail.data;

      if (options.json) {
        console.log(JSON.stringify(email, null, 2));
        return;
      }

      console.log('\n' + '\u2500'.repeat(60));
      console.log(`From: ${email.From?.EmailAddress?.Name || email.From?.EmailAddress?.Address || 'Unknown'}`);
      if (email.From?.EmailAddress?.Address) {
        console.log(`      <${email.From.EmailAddress.Address}>`);
      }
      console.log(`Subject: ${email.Subject || '(no subject)'}`);
      console.log(`Date: ${email.ReceivedDateTime ? new Date(email.ReceivedDateTime).toLocaleString() : 'Unknown'}`);

      if (email.ToRecipients && email.ToRecipients.length > 0) {
        const to = email.ToRecipients.map(r => r.EmailAddress?.Address).filter(Boolean).join(', ');
        console.log(`To: ${to}`);
      }

      if (email.CcRecipients && email.CcRecipients.length > 0) {
        const cc = email.CcRecipients.map(r => r.EmailAddress?.Address).filter(Boolean).join(', ');
        console.log(`Cc: ${cc}`);
      }

      if (email.HasAttachments) {
        console.log('Attachments: Yes');
      }

      console.log('\u2500'.repeat(60) + '\n');
      console.log(email.Body?.Content || email.BodyPreview || '(no content)');
      console.log('\n' + '\u2500'.repeat(60) + '\n');
      return;
    }

    // List emails
    if (options.json) {
      console.log(JSON.stringify({
        folder: apiFolder,
        page,
        limit,
        emails: emails.map((e, i) => ({
          index: skip + i + 1,
          id: e.Id,
          from: e.From?.EmailAddress?.Address,
          fromName: e.From?.EmailAddress?.Name,
          subject: e.Subject,
          preview: e.BodyPreview,
          receivedAt: e.ReceivedDateTime,
          isRead: e.IsRead,
          hasAttachments: e.HasAttachments,
          importance: e.Importance,
          flagged: e.Flag?.FlagStatus === 'Flagged',
        })),
      }, null, 2));
      return;
    }

    const folderDisplay = folder.charAt(0).toUpperCase() + folder.slice(1);
    const searchInfo = options.search ? ` - search: "${options.search}"` : '';
    const pageInfo = page > 1 ? ` (page ${page})` : '';
    console.log(`\n\ud83d\udcec ${folderDisplay}${searchInfo}${pageInfo}:\n`);
    console.log('\u2500'.repeat(70));

    if (emails.length === 0) {
      console.log('\n  No emails found.\n');
      return;
    }

    for (let i = 0; i < emails.length; i++) {
      const email = emails[i];
      const idx = skip + i + 1;
      const unreadMark = email.IsRead ? ' ' : '\u2022';
      const flagMark = email.Flag?.FlagStatus === 'Flagged' ? '\u2691' : ' ';
      const attachMark = email.HasAttachments ? '\ud83d\udcce' : ' ';
      const importanceMark = email.Importance === 'High' ? '!' : ' ';

      const from = email.From?.EmailAddress?.Name || email.From?.EmailAddress?.Address || 'Unknown';
      const subject = email.Subject || '(no subject)';
      const date = email.ReceivedDateTime ? formatDate(email.ReceivedDateTime) : '';

      // Format: [idx] marks | from | subject | date
      const marks = `${unreadMark}${flagMark}${attachMark}${importanceMark}`;
      const fromTrunc = truncate(from, 20);
      const subjectTrunc = truncate(subject, 35);

      console.log(`  [${idx.toString().padStart(2)}] ${marks} ${fromTrunc.padEnd(20)} ${subjectTrunc.padEnd(35)} ${date}`);
    }

    console.log('\n' + '\u2500'.repeat(70));
    console.log('\nCommands:');
    console.log(`  clippy mail -r <number>           # Read email`);
    console.log(`  clippy mail -p ${page + 1}                   # Next page`);
    console.log(`  clippy mail --unread              # Only unread`);
    console.log(`  clippy mail -s "keyword"          # Search emails`);
    console.log(`  clippy mail sent                  # Sent folder`);
    console.log('');
  });
