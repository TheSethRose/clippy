import { Command } from 'commander';
import { startKeepaliveSession } from '../lib/auth.js';

export const keepaliveCommand = new Command('keepalive')
  .description('Keep a live Outlook session open and refresh periodically to avoid token expiry')
  .option('--interval <minutes>', 'Refresh interval in minutes', '10')
  .option('--headless', 'Run headless (not recommended for keeping session alive)', false)
  .action(async (options: { interval?: string; headless?: boolean }) => {
    const intervalMinutes = parseInt(options.interval || '10', 10);
    if (isNaN(intervalMinutes) || intervalMinutes < 1) {
      console.error('Interval must be >= 1 minute');
      process.exit(1);
    }

    const headless = Boolean(options.headless);

    await startKeepaliveSession({ intervalMinutes, headless });
  });
