import { Command } from 'commander';
import { extractTokenViaPlaywright } from '../lib/auth.js';
import { validateSession } from '../lib/owa-client.js';
import { readFile, writeFile, mkdir, unlink } from 'fs/promises';
import { homedir } from 'os';
import { join } from 'path';
import { execSync } from 'child_process';

const TOKEN_CACHE_FILE = join(homedir(), '.config', 'clippy', 'token-cache.json');
const LAUNCHD_PLIST = join(homedir(), 'Library', 'LaunchAgents', 'com.clippy.refresh.plist');
const LAUNCHD_LABEL = 'com.clippy.refresh';

async function installService(): Promise<void> {
  const platform = process.platform;

  if (platform !== 'darwin') {
    console.log('Auto-install only supported on macOS.');
    console.log('For Linux, add this to your crontab:');
    console.log('  */30 * * * * /path/to/clippy refresh --silent');
    return;
  }

  // Find the clippy executable path
  const clippyPath = process.argv[1].replace(/\/src\/cli\.ts$/, '/src/cli.ts');
  const bunPath = execSync('which bun').toString().trim();

  const plistContent = `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>${LAUNCHD_LABEL}</string>
    <key>ProgramArguments</key>
    <array>
        <string>${bunPath}</string>
        <string>run</string>
        <string>${clippyPath}</string>
        <string>refresh</string>
        <string>--silent</string>
    </array>
    <key>StartInterval</key>
    <integer>1800</integer>
    <key>RunAtLoad</key>
    <true/>
    <key>StandardOutPath</key>
    <string>${join(homedir(), '.config', 'clippy', 'refresh.log')}</string>
    <key>StandardErrorPath</key>
    <string>${join(homedir(), '.config', 'clippy', 'refresh.log')}</string>
</dict>
</plist>`;

  // Ensure LaunchAgents directory exists
  await mkdir(join(homedir(), 'Library', 'LaunchAgents'), { recursive: true });

  // Write plist
  await writeFile(LAUNCHD_PLIST, plistContent, 'utf-8');

  // Load the service
  try {
    execSync(`launchctl unload ${LAUNCHD_PLIST} 2>/dev/null || true`);
    execSync(`launchctl load ${LAUNCHD_PLIST}`);
    console.log('Background refresh service installed.');
    console.log('Token will be refreshed every 30 minutes.');
    console.log(`Log file: ~/.config/clippy/refresh.log`);
  } catch (err) {
    console.error('Failed to load service:', err);
    process.exit(1);
  }
}

async function uninstallService(): Promise<void> {
  const platform = process.platform;

  if (platform !== 'darwin') {
    console.log('For Linux, remove the crontab entry manually:');
    console.log('  crontab -e');
    return;
  }

  try {
    execSync(`launchctl unload ${LAUNCHD_PLIST} 2>/dev/null || true`);
    await unlink(LAUNCHD_PLIST);
    console.log('Background refresh service removed.');
  } catch {
    console.log('Service was not installed or already removed.');
  }
}

function getJwtExpiration(token: string): number | null {
  try {
    const parts = token.split('.');
    if (parts.length !== 3) return null;
    const payload = JSON.parse(Buffer.from(parts[1], 'base64url').toString());
    return payload.exp ? payload.exp * 1000 : null;
  } catch {
    return null;
  }
}

export const refreshCommand = new Command('refresh')
  .description('Refresh the authentication token (for use in cron/scheduled tasks)')
  .option('--silent', 'No output unless error')
  .option('--force', 'Force refresh even if token is still valid')
  .option('--status', 'Show current token status without refreshing')
  .option('--install', 'Install background refresh service (keeps token alive)')
  .option('--uninstall', 'Remove background refresh service')
  .action(async (options: {
    silent?: boolean;
    force?: boolean;
    status?: boolean;
    install?: boolean;
    uninstall?: boolean;
  }) => {
    // Handle install/uninstall
    if (options.install) {
      await installService();
      return;
    }

    if (options.uninstall) {
      await uninstallService();
      return;
    }

    const log = (msg: string) => {
      if (!options.silent) console.log(msg);
    };

    // Check current token status
    try {
      const data = await readFile(TOKEN_CACHE_FILE, 'utf-8');
      const cached = JSON.parse(data);
      const now = Date.now();
      const expiresIn = Math.round((cached.expiresAt - now) / 60000);

      if (options.status) {
        if (cached.expiresAt <= now) {
          console.log('Token status: EXPIRED');
        } else {
          console.log(`Token status: Valid for ${expiresIn} minutes`);
          console.log(`Expires at: ${new Date(cached.expiresAt).toLocaleString()}`);
        }
        return;
      }

      // --silent should only control logging, not behavior.
      const forceRefresh = options.force;

      // Even if the JWT says it's still valid, Microsoft can invalidate sessions early.
      // So before we decide to skip refresh, we validate against OWA.
      const stillAccepted = await validateSession(cached.token);

      // Skip refresh only if token is still good (>10 min remaining), is accepted by OWA, and not forced
      if (!forceRefresh && expiresIn > 10 && stillAccepted) {
        log(`Token still valid for ${expiresIn} minutes, skipping refresh`);
        return;
      }

      if (!stillAccepted) {
        log('Cached token rejected by OWA (session invalidated early), refreshing...');
      } else {
        log(`Token expires in ${expiresIn} minutes, refreshing...`);
      }
    } catch {
      // No cached token
      if (options.status) {
        console.log('Token status: No cached token');
        return;
      }
      log('No cached token, extracting...');
    }

    // Refresh token (headless only - for unattended use)
    // Use the same persistent profile directory as interactive login so background refresh
    // can reuse the already-authenticated session.
    const browserProfile = join(homedir(), '.config', 'clippy', 'browser-profile');
    const result = await extractTokenViaPlaywright({ headless: true, timeout: 30000, userDataDir: browserProfile, fallbackToVisible: false });

    if (!result.success || !result.token) {
      console.error(`Error: ${result.error || 'Failed to refresh token'}`);
      console.error('You may need to run `clippy login --interactive` to re-authenticate.');
      process.exit(1);
    }

    // Validate the new token
    const isValid = await validateSession(result.token);
    if (!isValid) {
      console.error('Error: Refreshed token is invalid');
      process.exit(1);
    }

    // Save to cache
    const expiresAt = getJwtExpiration(result.token) || (Date.now() + 55 * 60 * 1000);
    const cacheDir = join(homedir(), '.config', 'clippy');
    await mkdir(cacheDir, { recursive: true });
    await writeFile(TOKEN_CACHE_FILE, JSON.stringify({
      token: result.token,
      graphToken: result.graphToken,
      expiresAt,
    }), 'utf-8');

    const expiresIn = Math.round((expiresAt - Date.now()) / 60000);
    log(`Token refreshed, valid for ${expiresIn} minutes`);
  });
