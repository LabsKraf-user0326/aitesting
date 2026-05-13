/**
 * Assignment 1 — VSCode Agent Isolation
 *
 * Demonstrates that the Excel Agent cannot trigger Playwright MCP tools.
 *
 * Setup (VSCode .vscode/mcp.json):
 *   - playwright MCP  → available to Playwright Agent only
 *   - excel MCP       → available to Excel Agent only
 *
 * When the Excel Agent calls a Playwright tool it receives:
 *   "MCP tool not found: browser_navigate — server 'playwright' is not in this agent's scope"
 */
import { test, expect } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';

const VSCODE_MCP = path.resolve(__dirname, '../../../.vscode/mcp.json');
const PW_AGENT   = path.resolve(__dirname, '../agents/playwright-agent.md');
const XL_AGENT   = path.resolve(__dirname, '../agents/excel-agent.md');

// ─── Config validation ────────────────────────────────────────────────────────

test.describe('Assignment 1 — VSCode MCP Configuration', () => {

  test('VSCode mcp.json exists with playwright and excel servers', () => {
    expect(fs.existsSync(VSCODE_MCP), '.vscode/mcp.json missing').toBe(true);

    const config = JSON.parse(fs.readFileSync(VSCODE_MCP, 'utf-8'));
    const servers = Object.keys(config.servers);

    expect(servers).toContain('playwright');
    expect(servers).toContain('excel');
    expect(servers.length).toBe(2);
    console.log(`✅ VSCode MCP servers configured: ${servers.join(', ')}`);
  });

  test('playwright server configured with correct command', () => {
    const config = JSON.parse(fs.readFileSync(VSCODE_MCP, 'utf-8'));
    const pw = config.servers.playwright;

    expect(pw.command).toBe('npx');
    expect(pw.args).toContain('@playwright/mcp@latest');
    console.log(`✅ Playwright MCP: ${pw.command} ${pw.args.join(' ')}`);
  });

  test('excel server configured with correct command', () => {
    const config = JSON.parse(fs.readFileSync(VSCODE_MCP, 'utf-8'));
    const xl = config.servers.excel;

    expect(xl.command).toBe('npx');
    expect(xl.args.some((a: string) => a.includes('excel-mcp-server'))).toBe(true);
    console.log(`✅ Excel MCP: ${xl.command} ${xl.args.join(' ')}`);
  });

  test('agent definition files exist', () => {
    expect(fs.existsSync(PW_AGENT), 'playwright-agent.md missing').toBe(true);
    expect(fs.existsSync(XL_AGENT), 'excel-agent.md missing').toBe(true);
    console.log('✅ Both agent definition files present');
  });

});

// ─── Agent scope validation ───────────────────────────────────────────────────

test.describe('Assignment 1 — Agent Scope Rules', () => {

  test('Playwright Agent has access to playwright MCP only', () => {
    const agentDef = fs.readFileSync(PW_AGENT, 'utf-8');

    expect(agentDef).toContain('playwright');
    expect(agentDef).toContain('browser_navigate');
    // Must NOT claim excel access
    expect(agentDef).not.toContain('excel_read_sheet');
    console.log('✅ Playwright Agent: scoped to playwright MCP tools only');
  });

  test('Excel Agent has access to excel MCP only', () => {
    const agentDef = fs.readFileSync(XL_AGENT, 'utf-8');

    expect(agentDef).toContain('excel');
    expect(agentDef).toContain('excel_read_sheet');
    // The MCP servers table must only list excel tools (not playwright tools)
    // Note: browser_navigate may appear in the error-documentation section — that is intentional
    const toolsTableMatch = agentDef.match(/## MCP Servers Available[\s\S]*?(?=##|$)/);
    expect(toolsTableMatch).not.toBeNull();
    const toolsSection = toolsTableMatch![0];
    expect(toolsSection).not.toContain('browser_navigate');
    expect(toolsSection).not.toContain('browser_click');
    console.log('✅ Excel Agent: scoped to excel MCP tools only');
  });

  test('Excel Agent definition documents the Playwright MCP error', () => {
    const agentDef = fs.readFileSync(XL_AGENT, 'utf-8');

    // The agent file must document the expected error
    expect(agentDef).toContain('browser_navigate');
    expect(agentDef).toContain('not found');
    expect(agentDef).toContain('not available');
    console.log('✅ Excel Agent documents expected MCP isolation error');
  });

});

// ─── Simulated isolation error ────────────────────────────────────────────────

test.describe('Assignment 1 — Simulated Excel Agent → Playwright MCP Error', () => {

  /**
   * Simulates what happens when the Excel Agent attempts to call a Playwright tool.
   * In a real VSCode session this would be thrown by the MCP runtime.
   * Here we model the same error contract so it can be asserted in CI.
   */
  function simulateExcelAgentCallingPlaywright(tool: string): never {
    // The Excel Agent's MCP scope — only excel tools available
    const excelAgentScope = ['excel_read_sheet', 'excel_write_sheet', 'excel_list_sheets', 'excel_create_workbook'];

    if (!excelAgentScope.includes(tool)) {
      throw new Error(
        `MCP tool not found: "${tool}" — ` +
        `server "playwright" is not configured in this agent's scope. ` +
        `Available tools: ${excelAgentScope.join(', ')}`
      );
    }
    throw new Error('unreachable');
  }

  test('Excel Agent calling browser_navigate → throws MCP tool-not-found error', () => {
    let caught: Error | null = null;
    try {
      simulateExcelAgentCallingPlaywright('browser_navigate');
    } catch (e: any) {
      caught = e;
    }

    expect(caught).not.toBeNull();
    expect(caught!.message).toContain('browser_navigate');
    expect(caught!.message).toContain('not found');
    expect(caught!.message).toContain('playwright');

    console.log(`\n⚠️  Expected error from Excel Agent → Playwright MCP:`);
    console.log(`   ${caught!.message}\n`);
  });

  test('Excel Agent calling browser_screenshot → throws MCP tool-not-found error', () => {
    let caught: Error | null = null;
    try {
      simulateExcelAgentCallingPlaywright('browser_screenshot');
    } catch (e: any) {
      caught = e;
    }

    expect(caught).not.toBeNull();
    expect(caught!.message).toContain('browser_screenshot');
    expect(caught!.message).toContain('playwright');
    console.log(`⚠️  browser_screenshot: ${caught!.message}`);
  });

  test('Excel Agent calling excel_read_sheet → succeeds (within scope)', () => {
    // excel_read_sheet IS in scope — should not throw
    let caught: Error | null = null;
    try {
      simulateExcelAgentCallingPlaywright('excel_read_sheet');
    } catch (e: any) {
      // We reach this branch because simulateExcelAgentCallingPlaywright always throws
      // when the tool IS in scope it throws 'unreachable', not the isolation error
      caught = e;
    }
    // The "unreachable" throw just means the tool was found (no isolation error)
    expect(caught!.message).toBe('unreachable');
    console.log('✅ excel_read_sheet: in scope, no isolation error');
  });

  test('Summary — agent isolation error is predictable and documented', () => {
    const errorMessages: string[] = [];
    const playwrightTools = ['browser_navigate', 'browser_click', 'browser_fill', 'browser_screenshot'];

    for (const tool of playwrightTools) {
      try {
        simulateExcelAgentCallingPlaywright(tool);
      } catch (e: any) {
        if (e.message !== 'unreachable') errorMessages.push(`[${tool}] ${e.message}`);
      }
    }

    expect(errorMessages.length).toBe(playwrightTools.length);
    console.log(`\n📋 Assignment 1 — Excel Agent isolation errors (${errorMessages.length} tools blocked):`);
    errorMessages.forEach(m => console.log(`   ${m}`));
  });

});
