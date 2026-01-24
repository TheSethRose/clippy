// Library exports for programmatic usage
export { resolveAuth, extractTokenViaPlaywright } from './lib/auth.js';
export type { AuthResult, PlaywrightTokenResult } from './lib/auth.js';

export { owaRequest, validateSession, getOwaUserInfo, getUserConfiguration } from './lib/owa-client.js';
export type { OwaRequestOptions, OwaResponse, OwaError, OwaUserInfo } from './lib/owa-client.js';

export { loadConfig, saveConfig, ensureConfigDir, CONFIG_DIR, CONFIG_FILE } from './lib/config.js';
export type { ClippyConfig } from './lib/config.js';
