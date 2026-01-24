import { Command } from 'commander';
import { resolveAuth } from '../lib/auth.js';
import { ensureConfigDir, saveConfig } from '../lib/config.js';

export const loginCommand = new Command('login')
  .description('Authenticate with OWA and validate session')
  .option('--token <token>', 'Use a specific token instead of extracting from browser')
  .option('-i, --interactive', 'Open browser to extract token automatically')
  .option('--check', 'Only check if session is valid, do not save')
  .action(async (options: { token?: string; interactive?: boolean; check?: boolean }) => {
    console.log('Checking OWA session...');

    const result = await resolveAuth({
      token: options.token,
      interactive: options.interactive,
    });

    if (!result.success) {
      console.error(`\nError: ${result.error}`);
      console.error('\nTo authenticate:');
      console.error('1. Run with --interactive to open browser and extract token automatically');
      console.error('   bun run src/cli.ts login --interactive');
      console.error('\n2. Or set CLIPPY_TOKEN environment variable manually:');
      console.error('   - Open https://outlook.office.com in your browser');
      console.error('   - Open DevTools (F12) → Network tab');
      console.error('   - Filter by "service.svc" and copy the Authorization header');
      console.error('   - export CLIPPY_TOKEN="eyJ..."');
      process.exit(1);
    }

    // Save config if not just checking
    if (!options.check) {
      try {
        await ensureConfigDir();
        await saveConfig({
          lastValidatedAt: new Date().toISOString(),
        });
      } catch (err) {
        // Non-fatal: continue even if config save fails
        console.warn('Warning: Could not save config');
      }
    }

    console.log('\n✓ Session valid');
    console.log('  Token: ***' + result.token!.slice(-8));
    console.log('\nYou are logged in to OWA. Run `clippy whoami` to see account details.');
  });
