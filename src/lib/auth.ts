import { chromium } from 'playwright';
import { setTimeout as sleep } from 'timers/promises';
import { homedir } from 'os';
import { join } from 'path';
import { readFile, writeFile, mkdir, unlink, open } from 'fs/promises';
import { validateSession } from './owa-client.js';

const LOCK_FILE = join(homedir(), '.config', 'clippy', 'browser.lock');
const LOCK_TIMEOUT = 60000; // Consider lock stale after 60s

interface LockHandle {
  release: () => Promise<void>;
}

async function acquireLock(timeout: number = 10000): Promise<LockHandle | null> {
  const startTime = Date.now();

  while (Date.now() - startTime < timeout) {
    try {
      // Check if lock exists and is stale
      try {
        const lockData = await readFile(LOCK_FILE, 'utf-8');
        const lockTime = parseInt(lockData, 10);
        if (Date.now() - lockTime > LOCK_TIMEOUT) {
          // Stale lock, remove it
          await unlink(LOCK_FILE);
        } else {
          // Lock is held, wait and retry
          await new Promise(r => setTimeout(r, 500));
          continue;
        }
      } catch {
        // Lock doesn't exist, proceed
      }

      // Try to create lock file (atomic via exclusive create)
      await mkdir(join(homedir(), '.config', 'clippy'), { recursive: true });
      const fd = await open(LOCK_FILE, 'wx');
      await fd.write(Date.now().toString());
      await fd.close();

      return {
        release: async () => {
          try {
            await unlink(LOCK_FILE);
          } catch {
            // Ignore
          }
        },
      };
    } catch (err: unknown) {
      if (err && typeof err === 'object' && 'code' in err && err.code === 'EEXIST') {
        // Lock file exists, wait and retry
        await new Promise(r => setTimeout(r, 500));
        continue;
      }
      // Other error, try again
      await new Promise(r => setTimeout(r, 100));
    }
  }

  return null; // Couldn't acquire lock
}

export interface AuthResult {
  success: boolean;
  token?: string;
  graphToken?: string;
  error?: string;
}

export interface PlaywrightTokenResult {
  success: boolean;
  token?: string;
  graphToken?: string;
  error?: string;
}

interface CachedToken {
  token: string;
  graphToken?: string;
  expiresAt: number;
}

const TOKEN_CACHE_FILE = join(homedir(), '.config', 'clippy', 'token-cache.json');
const REFRESH_THRESHOLD = 5 * 60 * 1000; // Refresh if less than 5 minutes remaining

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

async function getCachedToken(): Promise<{ cached: CachedToken; needsRefresh: boolean } | null> {
  try {
    const data = await readFile(TOKEN_CACHE_FILE, 'utf-8');
    const cached = JSON.parse(data) as CachedToken;
    const now = Date.now();

    if (cached.expiresAt <= now) {
      return null; // Expired
    }

    // Check if we need to refresh soon
    const needsRefresh = (cached.expiresAt - now) < REFRESH_THRESHOLD;
    return { cached, needsRefresh };
  } catch {
    // No cache or invalid cache
  }
  return null;
}

export async function setCachedToken(token: string, graphToken?: string): Promise<void> {
  try {
    const cacheDir = join(homedir(), '.config', 'clippy');
    await mkdir(cacheDir, { recursive: true });

    // Use actual JWT expiration time
    const expiresAt = getJwtExpiration(token) || (Date.now() + 55 * 60 * 1000);

    const cached: CachedToken = {
      token,
      graphToken,
      expiresAt,
    };
    await writeFile(TOKEN_CACHE_FILE, JSON.stringify(cached), 'utf-8');
  } catch {
    // Ignore cache write errors
  }
}

/**
 * Extract Bearer token by launching a browser and intercepting OWA requests.
 * Uses a persistent profile so the user only needs to log in once.
 * Tries headless first, then falls back to visible browser if login is needed.
 */
export async function extractTokenViaPlaywright(
  options: { headless?: boolean; timeout?: number } = {}
): Promise<PlaywrightTokenResult> {
  const { headless = true, timeout = 15000 } = options;

  // Try headless first (fast path for already logged-in users)
  const result = await tryExtractToken(headless, timeout);

  if (result.success) {
    return result;
  }

  // If headless failed and we haven't tried visible yet, retry with visible browser
  if (headless) {
    console.log('Session not found. Opening browser for login...');
    return tryExtractToken(false, 60000);  // Give more time for manual login
  }

  return result;
}

async function tryExtractToken(
  headless: boolean,
  timeout: number
): Promise<PlaywrightTokenResult> {
  // Acquire lock to prevent concurrent browser access
  const lock = await acquireLock(headless ? 5000 : 30000);
  if (!lock) {
    return {
      success: false,
      error: 'Another browser session is in progress, please wait',
    };
  }

  let context;
  try {
    // Use a dedicated profile directory for Clippy (persists login session)
    const userDataDir = join(homedir(), '.config', 'clippy', 'browser-profile');

    // Ensure the directory exists
    await mkdir(userDataDir, { recursive: true });

    // Launch persistent context - session will be remembered
    context = await chromium.launchPersistentContext(userDataDir, {
      headless,
      channel: 'chrome',
      args: ['--disable-blink-features=AutomationControlled'],
    });

    const page = context.pages()[0] || await context.newPage();

    let capturedToken: string | null = null;
    let capturedGraphToken: string | null = null;

    // Intercept requests to capture Bearer tokens
    page.on('request', request => {
      const url = request.url();
      const headers = request.headers();
      const authHeader = headers['authorization'];

      if (authHeader && authHeader.startsWith('Bearer ')) {
        const token = authHeader.replace('Bearer ', '');

        // Capture Outlook token
        if (url.includes('outlook.office.com') && !capturedToken) {
          capturedToken = token;
        }

        // Capture Graph token
        if (url.includes('graph.microsoft.com') && !capturedGraphToken) {
          capturedGraphToken = token;
        }
      }
    });

    if (!headless) {
      console.log('Please complete the login process in the browser...');
    }

    await page.goto('https://outlook.office.com/mail/', {
      waitUntil: 'domcontentloaded',
      timeout,
    });

    // Wait for token to be captured (max timeout)
    const startTime = Date.now();
    while (!capturedToken && (Date.now() - startTime) < timeout) {
      await page.waitForTimeout(500);
    }

    // If we have the Outlook token, wait a bit more to try to capture Graph token
    if (capturedToken && !capturedGraphToken) {
      // Try to trigger Graph API calls by interacting with the page
      const graphWaitTime = headless ? 3000 : 5000;
      const graphStart = Date.now();
      while (!capturedGraphToken && (Date.now() - graphStart) < graphWaitTime) {
        await page.waitForTimeout(500);
      }
    }

    await context.close();
    await lock.release();

    if (capturedToken) {
      return { success: true, token: capturedToken, graphToken: capturedGraphToken || undefined };
    }

    return {
      success: false,
      error: headless
        ? 'No active session found'
        : 'Timeout: No Bearer token captured. Make sure you completed the login.'
    };
  } catch (err) {
    if (context) {
      await context.close();
    }
    await lock.release();
    return {
      success: false,
      error: err instanceof Error ? err.message : 'Unknown error during token extraction',
    };
  }
}

export async function startKeepaliveSession(options: { intervalMinutes: number; headless?: boolean }): Promise<void> {
  const { intervalMinutes, headless = false } = options;

  // Use a dedicated profile directory for Clippy (persists login session)
  const userDataDir = join(homedir(), '.config', 'clippy', 'browser-profile');
  await mkdir(userDataDir, { recursive: true });

  const context = await chromium.launchPersistentContext(userDataDir, {
    headless,
    channel: 'chrome',
    args: ['--disable-blink-features=AutomationControlled'],
  });

  const page = context.pages()[0] || await context.newPage();

  let lastToken: string | null = null;
  let lastGraphToken: string | null = null;

  page.on('request', request => {
    const url = request.url();
    const headers = request.headers();
    const authHeader = headers['authorization'];

    if (authHeader && authHeader.startsWith('Bearer ')) {
      const token = authHeader.replace('Bearer ', '');

      if (url.includes('outlook.office.com')) {
        lastToken = token;
      }
      if (url.includes('graph.microsoft.com')) {
        lastGraphToken = token;
      }
    }
  });

  console.log(`Opening Outlook session (headless=${headless ? 'true' : 'false'})...`);
  await page.goto('https://outlook.office.com/mail/', { waitUntil: 'domcontentloaded', timeout: 60000 });

  console.log(`Keepalive active. Refreshing every ${intervalMinutes} minutes.`);

  // Main keepalive loop
  while (true) {
    // Give some time for requests to fire and tokens to be captured
    await sleep(2000);

    if (lastToken) {
      await setCachedToken(lastToken, lastGraphToken || undefined);
    }

    await sleep(intervalMinutes * 60 * 1000);

    try {
      await page.reload({ waitUntil: 'domcontentloaded', timeout: 60000 });
    } catch {
      // If reload fails, try to navigate again
      try {
        await page.goto('https://outlook.office.com/mail/', { waitUntil: 'domcontentloaded', timeout: 60000 });
      } catch {
        // ignore, will retry on next loop
      }
    }
  }
}

export async function resolveAuth(options: {
  token?: string;
  interactive?: boolean;
}): Promise<AuthResult> {
  const { token: cliToken, interactive = false } = options;

  // Priority 1: CLI flag
  if (cliToken) {
    const isValid = await validateSession(cliToken);
    if (isValid) {
      return { success: true, token: cliToken };
    }
    return { success: false, error: 'Provided token is invalid or expired' };
  }

  // Priority 2: Environment variable
  const envToken = process.env.CLIPPY_TOKEN;
  if (envToken) {
    const isValid = await validateSession(envToken);
    if (isValid) {
      return { success: true, token: envToken };
    }
    return {
      success: false,
      error: 'CLIPPY_TOKEN environment variable contains invalid or expired token',
    };
  }

  // Priority 3: Cached token (fast path - no browser needed)
  const cachedResult = await getCachedToken();
  if (cachedResult) {
    const { cached, needsRefresh } = cachedResult;
    const isValid = await validateSession(cached.token);

    if (isValid) {
      // If token is expiring soon, refresh in background (silent)
      if (needsRefresh && interactive) {
        extractTokenViaPlaywright({ headless: true, timeout: 10000 })
          .then(result => {
            if (result.success && result.token) {
              setCachedToken(result.token, result.graphToken);
            }
          })
          .catch(() => {}); // Ignore refresh errors
      }

      return {
        success: true,
        token: cached.token,
        graphToken: cached.graphToken,
      };
    }
    // Cache is stale, will re-extract below
  }

  // Priority 4: Interactive Playwright extraction
  if (interactive) {
    const playwrightResult = await extractTokenViaPlaywright();
    if (playwrightResult.success && playwrightResult.token) {
      const isValid = await validateSession(playwrightResult.token);
      if (isValid) {
        // Cache the token for future use
        await setCachedToken(playwrightResult.token, playwrightResult.graphToken);
        return {
          success: true,
          token: playwrightResult.token,
          graphToken: playwrightResult.graphToken,
        };
      }
      return {
        success: false,
        error: 'Extracted token is invalid or expired',
      };
    }
    return {
      success: false,
      error: playwrightResult.error || 'Failed to extract token via browser',
    };
  }

  // No token available and not interactive
  return {
    success: false,
    error: 'No token available. Set CLIPPY_TOKEN env var or run with --interactive to extract via browser.',
  };
}
