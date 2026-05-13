/**
 * Login Spec — driven entirely by Excel MCP scenarios.
 *
 * Flow:
 *   Excel MCP   → reads "Login Scenarios" sheet → 10 scenarios
 *   Playwright MCP → executes each scenario in Chromium
 *   Filesystem MCP → saves screenshots, logs, JSON per scenario
 *   Excel MCP   → writes PASS/FAIL status back into workbook
 *
 * Target app: https://www.saucedemo.com (Sauce Labs demo)
 *   valid creds   : standard_user / secret_sauce
 *   locked user   : locked_out_user / secret_sauce
 *   invalid creds : any user / wrong_password
 */
import { test, expect } from '@playwright/test';
import * as path from 'path';
import { MCPAgent, ExcelMCP, FilesystemMCP, AgentResult } from '../agent/mcp-agent';
import { PromptBuilder, ActionExecutor } from '../agent/prompt-builder';

const SCENARIOS_XL = path.resolve(__dirname, '../scenarios/login-scenarios.xlsx');
const RESULTS_DIR  = path.resolve(__dirname, '../results');

// shared state across tests in this file
let agent: MCPAgent;
const runResults: AgentResult[] = [];

test.describe.configure({ mode: 'serial' });

test.beforeAll(async () => {
  const fs = new FilesystemMCP(RESULTS_DIR);
  fs.ensureDir('screenshots');
  fs.ensureDir('logs');
  fs.ensureDir('reports');
  fs.log('logs/login.log', '══ Login spec started ══');
});

// ─── STEP 1: Excel MCP — build workbook ──────────────────────────────────────

test('Excel MCP — load login-scenarios.xlsx', async () => {
  const excel = new ExcelMCP(SCENARIOS_XL);
  await excel.load();

  const scenarios  = excel.getLoginScenarios();
  const navScenarios = excel.getNavScenarios();
  const templates  = excel.getPromptTemplates();

  expect(scenarios.length).toBeGreaterThanOrEqual(10);
  expect(navScenarios.length).toBeGreaterThanOrEqual(5);
  expect(templates.length).toBeGreaterThanOrEqual(4);

  console.log(`✅ Excel MCP: ${scenarios.length} login, ${navScenarios.length} nav, ${templates.length} templates`);

  const fs = new FilesystemMCP(RESULTS_DIR);
  fs.log('logs/login.log', `Excel MCP: ${scenarios.length} login scenarios, ${templates.length} prompt templates`);
});

// ─── STEP 2: Playwright MCP — login scenarios ─────────────────────────────────

test('Playwright MCP — valid login (standard_user)', async ({ page }) => {
  agent = new MCPAgent(page, SCENARIOS_XL, RESULTS_DIR);
  await agent.init();

  const excel  = new ExcelMCP(SCENARIOS_XL);
  await excel.load();
  const s = excel.getLoginScenarios().find(s => s.id === 1)!;

  const builder  = new PromptBuilder(excel.getPromptTemplates());
  const built    = builder.fromLoginScenario(s);
  const executor = new ActionExecutor(page);
  const start    = Date.now();

  agent.log(`▶ Prompt: ${s.prompt}`);
  let pass = true, err: string | undefined;

  for (const action of built.actions) {
    const { ok, detail } = await executor.run(action);
    if (!ok) { pass = false; err = detail; break; }
  }

  const actualUrl = page.url();
  const shot = await agent.playwright.screenshot(`valid-login-standard`);

  const result: AgentResult = {
    scenarioId: s.id, name: s.name,
    status: pass ? 'PASS' : 'FAIL',
    duration: Date.now() - start, actualUrl, error: err, screenshot: shot,
    runAt: new Date().toISOString(),
  };
  runResults.push(result);

  expect(pass, err).toBe(true);
  expect(actualUrl).toContain('/inventory.html');
  agent.log(`✅ ${s.name}: PASS (${result.duration}ms)`);
});

test('Playwright MCP — valid login (problem_user)', async ({ page }) => {
  const excel = new ExcelMCP(SCENARIOS_XL);
  await excel.load();
  const s = excel.getLoginScenarios().find(s => s.id === 2)!;

  await page.goto(s.url);
  await page.locator('#user-name').fill(s.username);
  await page.locator('#password').fill(s.password);
  await page.locator('[data-test="login-button"]').click();
  await page.waitForTimeout(500);

  const shot = await page.screenshot({ path: `tests/results/screenshots/problem-user-${Date.now()}.png` });
  const url  = page.url();

  runResults.push({ scenarioId: s.id, name: s.name, status: url.includes('/inventory') ? 'PASS' : 'FAIL', duration: 0, actualUrl: url, runAt: new Date().toISOString() });
  expect(url).toContain('/inventory');
  console.log(`✅ ${s.name}: ${url}`);
});

test('Playwright MCP — valid login (performance_glitch_user)', async ({ page }) => {
  const excel = new ExcelMCP(SCENARIOS_XL);
  await excel.load();
  const s = excel.getLoginScenarios().find(s => s.id === 3)!;

  await page.goto(s.url);
  await page.locator('#user-name').fill(s.username);
  await page.locator('#password').fill(s.password);
  await page.locator('[data-test="login-button"]').click();
  // performance_glitch_user is deliberately slow
  await page.waitForURL('**/inventory.html', { timeout: 15000 });

  const url = page.url();
  runResults.push({ scenarioId: s.id, name: s.name, status: 'PASS', duration: 0, actualUrl: url, runAt: new Date().toISOString() });
  expect(url).toContain('/inventory');
  console.log(`✅ ${s.name}: logged in despite performance glitch`);
});

test('Playwright MCP — invalid login (wrong password)', async ({ page }) => {
  const excel = new ExcelMCP(SCENARIOS_XL);
  await excel.load();
  const s = excel.getLoginScenarios().find(s => s.id === 4)!;

  await page.goto(s.url);
  await page.locator('#user-name').fill(s.username);
  await page.locator('#password').fill(s.password);
  await page.locator('[data-test="login-button"]').click();
  await page.waitForTimeout(500);

  const errorText = await page.locator('[data-test="error"]').textContent() ?? '';
  const shot = `tests/results/screenshots/invalid-login-${Date.now()}.png`;
  await page.screenshot({ path: shot });

  runResults.push({ scenarioId: s.id, name: s.name, status: errorText.includes('do not match') ? 'PASS' : 'FAIL', duration: 0, actualUrl: page.url(), screenshot: shot, runAt: new Date().toISOString() });
  expect(errorText).toContain('do not match');
  console.log(`✅ ${s.name}: error="${errorText.trim()}"`);
});

test('Playwright MCP — invalid login (empty username)', async ({ page }) => {
  const excel = new ExcelMCP(SCENARIOS_XL);
  await excel.load();
  const s = excel.getLoginScenarios().find(s => s.id === 5)!;

  await page.goto(s.url);
  await page.locator('#password').fill(s.password);
  await page.locator('[data-test="login-button"]').click();

  const errorText = await page.locator('[data-test="error"]').textContent() ?? '';
  runResults.push({ scenarioId: s.id, name: s.name, status: errorText.includes('Username is required') ? 'PASS' : 'FAIL', duration: 0, actualUrl: page.url(), runAt: new Date().toISOString() });
  expect(errorText).toContain('Username is required');
  console.log(`✅ ${s.name}: error="${errorText.trim()}"`);
});

test('Playwright MCP — invalid login (empty password)', async ({ page }) => {
  const excel = new ExcelMCP(SCENARIOS_XL);
  await excel.load();
  const s = excel.getLoginScenarios().find(s => s.id === 6)!;

  await page.goto(s.url);
  await page.locator('#user-name').fill(s.username);
  await page.locator('[data-test="login-button"]').click();

  const errorText = await page.locator('[data-test="error"]').textContent() ?? '';
  runResults.push({ scenarioId: s.id, name: s.name, status: errorText.includes('Password is required') ? 'PASS' : 'FAIL', duration: 0, actualUrl: page.url(), runAt: new Date().toISOString() });
  expect(errorText).toContain('Password is required');
  console.log(`✅ ${s.name}: error="${errorText.trim()}"`);
});

test('Playwright MCP — locked out user', async ({ page }) => {
  const excel = new ExcelMCP(SCENARIOS_XL);
  await excel.load();
  const s = excel.getLoginScenarios().find(s => s.id === 7)!;

  await page.goto(s.url);
  await page.locator('#user-name').fill(s.username);
  await page.locator('#password').fill(s.password);
  await page.locator('[data-test="login-button"]').click();

  const errorText = await page.locator('[data-test="error"]').textContent() ?? '';
  const shot = `tests/results/screenshots/locked-user-${Date.now()}.png`;
  await page.screenshot({ path: shot });

  runResults.push({ scenarioId: s.id, name: s.name, status: errorText.includes('locked out') ? 'PASS' : 'FAIL', duration: 0, actualUrl: page.url(), screenshot: shot, runAt: new Date().toISOString() });
  expect(errorText).toContain('locked out');
  console.log(`✅ ${s.name}: locked out error confirmed`);
});

test('Playwright MCP — SQL injection handled gracefully', async ({ page }) => {
  const excel = new ExcelMCP(SCENARIOS_XL);
  await excel.load();
  const s = excel.getLoginScenarios().find(s => s.id === 8)!;

  await page.goto(s.url);
  await page.locator('#user-name').fill(s.username);
  await page.locator('#password').fill(s.password);
  await page.locator('[data-test="login-button"]').click();

  const errorVisible = await page.locator('[data-test="error"]').isVisible();
  const url = page.url();
  const notLogged = !url.includes('/inventory');

  runResults.push({ scenarioId: s.id, name: s.name, status: (errorVisible || notLogged) ? 'PASS' : 'FAIL', duration: 0, actualUrl: url, runAt: new Date().toISOString() });
  expect(errorVisible || notLogged).toBe(true);
  console.log(`✅ ${s.name}: SQL injection rejected`);
});

test('Playwright MCP — login and logout cycle', async ({ page }) => {
  const excel = new ExcelMCP(SCENARIOS_XL);
  await excel.load();
  const s = excel.getLoginScenarios().find(s => s.id === 9)!;

  // Login
  await page.goto(s.url);
  await page.locator('#user-name').fill(s.username);
  await page.locator('#password').fill(s.password);
  await page.locator('[data-test="login-button"]').click();
  await page.waitForURL('**/inventory.html');

  const shot1 = `tests/results/screenshots/pre-logout-${Date.now()}.png`;
  await page.screenshot({ path: shot1 });

  // Logout
  await page.locator('#react-burger-menu-btn').click();
  await page.locator('#logout_sidebar_link').click();
  await page.waitForURL('https://www.saucedemo.com/');

  const shot2 = `tests/results/screenshots/post-logout-${Date.now()}.png`;
  await page.screenshot({ path: shot2 });

  const finalUrl = page.url();
  runResults.push({ scenarioId: s.id, name: s.name, status: 'PASS', duration: 0, actualUrl: finalUrl, screenshot: shot2, runAt: new Date().toISOString() });
  expect(finalUrl).toBe('https://www.saucedemo.com/');
  console.log(`✅ ${s.name}: logout successful`);
});

test('Playwright MCP — session persists after page refresh', async ({ page }) => {
  const excel = new ExcelMCP(SCENARIOS_XL);
  await excel.load();
  const s = excel.getLoginScenarios().find(s => s.id === 10)!;

  await page.goto(s.url);
  await page.locator('#user-name').fill(s.username);
  await page.locator('#password').fill(s.password);
  await page.locator('[data-test="login-button"]').click();
  await page.waitForURL('**/inventory.html');

  await page.reload({ waitUntil: 'domcontentloaded' });
  const url = page.url();

  runResults.push({ scenarioId: s.id, name: s.name, status: url.includes('/inventory') ? 'PASS' : 'FAIL', duration: 0, actualUrl: url, runAt: new Date().toISOString() });
  expect(url).toContain('/inventory');
  console.log(`✅ ${s.name}: session persisted`);
});

// ─── STEP 3: Filesystem MCP + Excel MCP — write results ──────────────────────

test('Excel MCP + Filesystem MCP — write login results', async ({ page }) => {
  agent = agent ?? new MCPAgent(page, SCENARIOS_XL, RESULTS_DIR);
  await agent.init();

  // Write results to Excel
  await agent.excel.writeResults(runResults);

  // Write JSON via Filesystem MCP
  agent.filesystem.writeJson('reports/login-results.json', {
    summary: {
      total:  runResults.length,
      passed: runResults.filter(r => r.status === 'PASS').length,
      failed: runResults.filter(r => r.status === 'FAIL').length,
      runAt:  new Date().toISOString(),
    },
    results: runResults,
  });

  // Write HTML report
  agent.filesystem.writeHtml('reports/login-report.html', 'Login Scenarios', runResults);

  // Log
  agent.log(`Results written — ${runResults.filter(r => r.status === 'PASS').length}/${runResults.length} PASS`);
  agent.flushLog();

  expect(agent.filesystem.exists('reports/login-results.json')).toBe(true);
  expect(agent.filesystem.exists('reports/login-report.html')).toBe(true);
  console.log(`✅ Excel + Filesystem: results written`);
  console.log(`   Screenshots: ${agent.filesystem.list('screenshots').length} files`);
  console.log(`   Reports: ${agent.filesystem.list('reports').length} files`);
});
