#!/usr/bin/env bun
import { Command } from 'commander';
import { loginCommand } from './commands/login.js';
import { whoamiCommand } from './commands/whoami.js';
import { calendarCommand } from './commands/calendar.js';
import { findtimeCommand } from './commands/findtime.js';

const program = new Command();

program
  .name('clippy')
  .description('CLI for Microsoft 365/OWA')
  .version('0.1.0');

program.addCommand(loginCommand);
program.addCommand(whoamiCommand);
program.addCommand(calendarCommand);
program.addCommand(findtimeCommand);

program.parse();
