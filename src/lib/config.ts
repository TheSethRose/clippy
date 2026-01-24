import { homedir } from 'os';
import { join } from 'path';

export const CONFIG_DIR = join(homedir(), '.config', 'clippy');
export const CONFIG_FILE = join(CONFIG_DIR, 'config.json');

export interface ClippyConfig {
  lastValidatedAt?: string;
}

export async function loadConfig(): Promise<ClippyConfig> {
  try {
    const file = Bun.file(CONFIG_FILE);
    if (await file.exists()) {
      return await file.json();
    }
  } catch {
    // Config doesn't exist or is invalid
  }
  return {};
}

export async function saveConfig(config: ClippyConfig): Promise<void> {
  await Bun.write(CONFIG_FILE, JSON.stringify(config, null, 2));
}

export async function ensureConfigDir(): Promise<void> {
  const fs = await import('fs/promises');
  await fs.mkdir(CONFIG_DIR, { recursive: true });
}
