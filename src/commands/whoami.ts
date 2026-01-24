import { Command } from 'commander';
import { resolveAuth } from '../lib/auth.js';
import { getOwaUserInfo } from '../lib/owa-client.js';

export const whoamiCommand = new Command('whoami')
  .description('Show authenticated user information')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token instead of extracting from browser')
  .option('-i, --interactive', 'Open browser to extract token automatically')
  .action(async (options: { json?: boolean; token?: string; interactive?: boolean }) => {
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

    const userInfo = await getOwaUserInfo(authResult.token!);

    if (!userInfo.ok || !userInfo.data) {
      if (options.json) {
        console.log(
          JSON.stringify(
            {
              error: userInfo.error?.message || 'Failed to fetch user info',
              authenticated: true,
            },
            null,
            2
          )
        );
      } else {
        console.log('✓ Authenticated');
        console.log('  Could not fetch user details from OWA API');
      }
      process.exit(0);
    }

    const { displayName, email } = userInfo.data;

    if (options.json) {
      console.log(
        JSON.stringify(
          {
            displayName,
            email,
            authenticated: true,
          },
          null,
          2
        )
      );
    } else {
      console.log('✓ Authenticated');
      console.log(`  Name: ${displayName}`);
      console.log(`  Email: ${email}`);
    }
  });
