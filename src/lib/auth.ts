import { chromium } from 'playwright';
import { homedir } from 'os';
import { join } from 'path';
import { validateSession } from './owa-client.js';

export interface AuthResult {
  success: boolean;
  token?: string;
  error?: string;
}

export interface PlaywrightTokenResult {
  success: boolean;
  token?: string;
  error?: string;
}

/**
 * Extract Bearer token by launching a browser and intercepting OWA requests.
 * Uses a persistent profile so the user only needs to log in once.
 */
export async function extractTokenViaPlaywright(
  options: { headless?: boolean; timeout?: number } = {}
): Promise<PlaywrightTokenResult> {
  const { headless = false, timeout = 60000 } = options;

  let context;
  try {
    // Use a dedicated profile directory for Clippy (persists login session)
    const userDataDir = join(homedir(), '.config', 'clippy', 'browser-profile');

    // Ensure the directory exists
    const fs = await import('fs/promises');
    await fs.mkdir(userDataDir, { recursive: true });

    // Launch persistent context - session will be remembered
    context = await chromium.launchPersistentContext(userDataDir, {
      headless,
      channel: 'chrome',
      args: ['--disable-blink-features=AutomationControlled'],
    });

    const page = context.pages()[0] || await context.newPage();

    let capturedToken: string | null = null;

    // Intercept requests to capture Bearer token
    page.on('request', request => {
      if (request.url().includes('outlook.office.com') && !capturedToken) {
        const headers = request.headers();
        const authHeader = headers['authorization'];
        if (authHeader && authHeader.startsWith('Bearer ')) {
          capturedToken = authHeader.replace('Bearer ', '');
        }
      }
    });

    console.log('Opening browser to capture OWA token...');
    console.log('If not logged in, please complete the login process.');

    await page.goto('https://outlook.office.com/mail/', {
      waitUntil: 'domcontentloaded',
      timeout,
    });

    // Wait for token to be captured (max timeout)
    const startTime = Date.now();
    while (!capturedToken && (Date.now() - startTime) < timeout) {
      await page.waitForTimeout(500);
    }

    await context.close();

    if (capturedToken) {
      return { success: true, token: capturedToken };
    }

    return {
      success: false,
      error: 'Timeout: No Bearer token captured. Make sure you are logged in to OWA.'
    };
  } catch (err) {
    if (context) {
      await context.close();
    }
    return {
      success: false,
      error: err instanceof Error ? err.message : 'Unknown error during token extraction',
    };
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

  // Priority 3: Interactive Playwright extraction
  if (interactive) {
    const playwrightResult = await extractTokenViaPlaywright();
    if (playwrightResult.success && playwrightResult.token) {
      const isValid = await validateSession(playwrightResult.token);
      if (isValid) {
        return { success: true, token: playwrightResult.token };
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
