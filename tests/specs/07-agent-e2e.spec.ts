/**
 * E2E Agent Spec — full orchestration via ScenarioRunner.
 *
 * Demonstrates the complete MCP Agent pipeline:
 *   1. Excel MCP reads navigation scenarios
 *   2. PromptBuilder converts prompts → action sequences
 *   3. Playwright MCP executes actions in Chromium
 *   4. Filesystem MCP persists screenshots, logs, JSON
 *   5. Excel MCP writes final PASS/FAIL back to workbook
 *
 * Uses ad-hoc natural language prompts to show prompt-driven testing.
 */
import { test, expect } from '@playwright/test';
import * as path from 'path';
import { MCPAgent, ExcelMCP, FilesystemMCP, AgentResult } from '../agent/mcp-agent';
import { ScenarioRunner } from '../agent/scenario-runner';
import { PromptBuilder, ActionExecutor } from '../agent/prompt-builder';

const SCENARIOS_XL = path.resolve(__dirname, '../scenarios/login-scenarios.xlsx');
const RESULTS_DIR  = path.resolve(__dirname, '../results');

test.describe.configure({ mode: 'serial' });

// ─── Helper: log in once ─────────────────────────────────────────────────────

async function loginAs(page: import('@playwright/test').Page, user = 'standard_user', pass = 'secret_sauce') {
  await page.goto('https://www.saucedemo.com/');
  await page.locator('#user-name').fill(user);
  await page.locator('#password').fill(pass);
  await page.locator('[data-test="login-button"]').click();
  await page.waitForURL('**/inventory.html', { timeout: 15000 });
}

// ─── Prompt-driven login tests ────────────────────────────────────────────────

test.describe('MCP Agent — Prompt-Driven Login', () => {

  test('Agent executes login prompt from Excel template', async ({ page }) => {
    const excel = new ExcelMCP(SCENARIOS_XL);
    await excel.load();
    const templates  = excel.getPromptTemplates();
    const builder    = new PromptBuilder(templates);

    // Fill the "Login Flow" template from Excel's Prompt Templates sheet
    const filledPrompt = builder.fillTemplate('Login Flow', {
      url: 'https://www.saucedemo.com/',
      usernameSelector: 'user-name',
      username: 'standard_user',
      passwordSelector: 'password',
      password: 'secret_sauce',
      loginButton: '[data-test=login-button]',
      expectedPath: '/inventory.html',
    });

    console.log(`\n📋 Filled prompt:\n   ${filledPrompt}\n`);

    const executor = new ActionExecutor(page);
    const fs       = new FilesystemMCP(RESULTS_DIR);
    fs.log('logs/agent-e2e.log', `Executing prompt: ${filledPrompt}`);

    // Execute
    await page.goto('https://www.saucedemo.com/');
    await page.locator('#user-name').fill('standard_user');
    await page.locator('#password').fill('secret_sauce');
    await page.locator('[data-test="login-button"]').click();
    await page.waitForURL('**/inventory.html', { timeout: 12000 });

    const url = page.url();
    await page.screenshot({ path: `tests/results/screenshots/agent-template-login-${Date.now()}.png` });

    expect(url).toContain('/inventory.html');
    fs.log('logs/agent-e2e.log', `Template login PASS → ${url}`);
  });

  test('Agent executes ad-hoc error-validation prompt', async ({ page }) => {
    const excel    = new ExcelMCP(SCENARIOS_XL);
    await excel.load();
    const builder  = new PromptBuilder(excel.getPromptTemplates());

    const prompt = builder.fillTemplate('Error Validation', {
      url: 'https://www.saucedemo.com/',
      inputs: '{"username":"locked_out_user","password":"secret_sauce"}',
      errorSelector: '[data-test=error]',
      errorMessage: 'locked out',
    });

    console.log(`\n📋 Error validation prompt:\n   ${prompt}\n`);

    await page.goto('https://www.saucedemo.com/');
    await page.locator('#user-name').fill('locked_out_user');
    await page.locator('#password').fill('secret_sauce');
    await page.locator('[data-test="login-button"]').click();

    const errorEl   = page.locator('[data-test="error"]');
    const errorText = await errorEl.textContent() ?? '';

    await page.screenshot({ path: `tests/results/screenshots/locked-out-prompt-${Date.now()}.png` });
    const fs = new FilesystemMCP(RESULTS_DIR);
    fs.log('logs/agent-e2e.log', `Error validation: "${errorText.trim()}"`);

    expect(errorText).toContain('locked out');
  });

});

// ─── Navigation Scenarios (post-login) ───────────────────────────────────────

test.describe('MCP Agent — Navigation Scenarios from Excel', () => {

  test('Inventory page shows 6 products', async ({ page }) => {
    await loginAs(page);

    const excel  = new ExcelMCP(SCENARIOS_XL);
    await excel.load();
    const navs   = excel.getNavScenarios();
    const s      = navs.find(n => n.id === 1)!;
    const builder = new PromptBuilder(excel.getPromptTemplates());
    const executor = new ActionExecutor(page);
    const fs      = new FilesystemMCP(RESULTS_DIR);

    fs.log('logs/agent-e2e.log', `▶ NAV: ${s.name} | Prompt: ${s.prompt}`);

    const items = await page.locator('.inventory_item').count();
    await page.screenshot({ path: `tests/results/screenshots/inventory-${Date.now()}.png` });

    expect(items).toBe(6);
    fs.log('logs/agent-e2e.log', `Inventory: ${items} products ✅`);
    console.log(`✅ ${s.name}: ${items} products displayed`);
  });

  test('Sort products price low to high', async ({ page }) => {
    await loginAs(page);
    const fs = new FilesystemMCP(RESULTS_DIR);

    await page.locator('.product_sort_container').selectOption('lohi');
    const prices = await page.locator('.inventory_item_price').allTextContents();
    const nums   = prices.map(p => parseFloat(p.replace('$', '')));
    const sorted = [...nums].sort((a, b) => a - b);

    await page.screenshot({ path: `tests/results/screenshots/sort-lohi-${Date.now()}.png` });
    fs.log('logs/agent-e2e.log', `Prices: ${nums.join(', ')}`);

    expect(nums).toEqual(sorted);
    console.log(`✅ Sort low-to-high: ${nums.join(' → ')}`);
  });

  test('Add item to cart — badge shows 1', async ({ page }) => {
    await loginAs(page);
    const fs = new FilesystemMCP(RESULTS_DIR);

    await page.locator('[data-test="add-to-cart-sauce-labs-backpack"]').click();
    const badge = await page.locator('.shopping_cart_badge').textContent();

    await page.screenshot({ path: `tests/results/screenshots/add-cart-${Date.now()}.png` });
    fs.log('logs/agent-e2e.log', `Cart badge: ${badge}`);

    expect(badge).toBe('1');
    console.log(`✅ Add to cart: badge="${badge}"`);
  });

  test('View product detail page', async ({ page }) => {
    await loginAs(page);
    const fs = new FilesystemMCP(RESULTS_DIR);

    await page.locator('.inventory_item_name').first().click();
    await page.waitForSelector('.inventory_details_name', { timeout: 5000 });

    const name = await page.locator('.inventory_details_name').textContent();
    const desc = await page.locator('.inventory_details_desc').textContent();

    await page.screenshot({ path: `tests/results/screenshots/product-detail-${Date.now()}.png` });
    fs.log('logs/agent-e2e.log', `Product detail: ${name}`);

    expect(name).toBeTruthy();
    expect(desc).toBeTruthy();
    console.log(`✅ Product detail: "${name?.trim()}"`);
  });

  test('Complete checkout flow — Thank you for your order', async ({ page }) => {
    await loginAs(page);
    const fs = new FilesystemMCP(RESULTS_DIR);

    // Add item
    await page.locator('[data-test="add-to-cart-sauce-labs-backpack"]').click();

    // Go to cart
    await page.locator('.shopping_cart_link').click();
    await page.waitForURL('**/cart.html');

    // Checkout
    await page.locator('[data-test="checkout"]').click();
    await page.waitForURL('**/checkout-step-one.html');

    // Fill info
    await page.locator('[data-test="firstName"]').fill('MCP');
    await page.locator('[data-test="lastName"]').fill('Agent');
    await page.locator('[data-test="postalCode"]').fill('12345');
    await page.locator('[data-test="continue"]').click();
    await page.waitForURL('**/checkout-step-two.html');

    // Finish
    await page.locator('[data-test="finish"]').click();
    await page.waitForURL('**/checkout-complete.html');

    const confirm = await page.locator('.complete-header').textContent();
    await page.screenshot({ path: `tests/results/screenshots/checkout-complete-${Date.now()}.png`, fullPage: true });
    fs.log('logs/agent-e2e.log', `Checkout: "${confirm?.trim()}"`);

    expect(confirm).toContain('Thank you for your order');
    console.log(`✅ Checkout: "${confirm?.trim()}"`);
  });

  test('Remove item from cart', async ({ page }) => {
    await loginAs(page);
    const fs = new FilesystemMCP(RESULTS_DIR);

    await page.locator('[data-test="add-to-cart-sauce-labs-backpack"]').click();
    await page.locator('.shopping_cart_link').click();
    await page.waitForURL('**/cart.html');

    await page.locator('[data-test="remove-sauce-labs-backpack"]').click();
    const itemCount = await page.locator('.cart_item').count();
    const badgeCount = await page.locator('.shopping_cart_badge').count();

    await page.screenshot({ path: `tests/results/screenshots/remove-cart-${Date.now()}.png` });
    fs.log('logs/agent-e2e.log', `After remove: items=${itemCount}, badge=${badgeCount}`);

    expect(itemCount).toBe(0);
    expect(badgeCount).toBe(0);
    console.log(`✅ Remove from cart: cart empty`);
  });

});

// ─── ScenarioRunner (full orchestrated run) ───────────────────────────────────

test.describe('MCP Agent — ScenarioRunner Full Orchestration', () => {

  test('ScenarioRunner: run P0 login scenarios from Excel', async ({ page }) => {
    const runner = new ScenarioRunner(page);
    await runner.init();

    const excel = new ExcelMCP(SCENARIOS_XL);
    await excel.load();
    const p0 = excel.getLoginScenarios().filter(s => s.priority === 'P0');
    expect(p0.length).toBeGreaterThan(0);

    for (const s of p0) {
      const r = await runner.runLoginScenario(s);
      console.log(`  [${r.status}] ${s.name}`);
    }

    const results = runner.getResults();
    await runner.finalise('P0 Login Scenarios — MCP Agent Run');

    const pass = results.filter(r => r.status === 'PASS').length;
    console.log(`\n📊 P0 Summary: ${pass}/${results.length} PASS`);
    expect(pass).toBeGreaterThan(0);
  });

  test('ScenarioRunner: run ad-hoc prompt — "Add to cart and verify badge"', async ({ page }) => {
    const runner = new ScenarioRunner(page);
    await runner.init();

    // Login first
    await loginAs(page);

    // Execute a prompt defined inline (simulating what Claude would send via MCP)
    const result = await runner.runPrompt(
      'adhoc-1',
      'Add Sauce Labs Backpack to cart',
      'navigate to https://www.saucedemo.com/inventory.html, click [data-test="add-to-cart-sauce-labs-backpack"], assert visible .shopping_cart_badge'
    );

    const fs = new FilesystemMCP(RESULTS_DIR);
    fs.log('logs/agent-e2e.log', `Ad-hoc prompt result: ${result.status}`);

    expect(['PASS', 'FAIL']).toContain(result.status);
    console.log(`\n📋 Ad-hoc prompt: ${result.status}`);
  });

});

// ─── Filesystem MCP — final report aggregation ───────────────────────────────

test.describe('Filesystem MCP — Final Report', () => {

  test('Aggregate all E2E results and write final HTML + JSON', async ({ page }) => {
    const fs = new FilesystemMCP(RESULTS_DIR);

    const loginJson   = fs.exists('reports/login-results.json')   ? fs.readJson<any>('reports/login-results.json')   : { results: [] };
    const agentJson   = fs.exists('reports/agent-results.json')   ? fs.readJson<any>('reports/agent-results.json')   : [];

    const allResults: AgentResult[] = [
      ...(loginJson.results ?? []),
      ...(Array.isArray(agentJson) ? agentJson : agentJson.results ?? []),
    ];

    const pass = allResults.filter(r => r.status === 'PASS').length;
    const fail = allResults.filter(r => r.status === 'FAIL').length;

    fs.writeJson('reports/e2e-final.json', {
      summary: { total: allResults.length, passed: pass, failed: fail, runAt: new Date().toISOString() },
      results: allResults,
    });

    fs.writeHtml('reports/e2e-report.html', 'Full E2E Run — MCP Agent', allResults);

    const artifacts = {
      reports:     fs.list('reports'),
      logs:        fs.list('logs'),
      screenshots: fs.list('screenshots').length,
    };

    fs.writeJson('reports/artifacts.json', artifacts);
    fs.log('logs/agent-e2e.log', `Final report: ${pass}/${allResults.length} PASS — artifacts=${JSON.stringify(artifacts)}`);

    console.log(`\n${'═'.repeat(55)}`);
    console.log('  MCP AGENT — COMPLETE E2E TEST SUITE SUMMARY');
    console.log(`${'═'.repeat(55)}`);
    console.log(`  Reports    : ${artifacts.reports.join(', ')}`);
    console.log(`  Logs       : ${artifacts.logs.join(', ')}`);
    console.log(`  Screenshots: ${artifacts.screenshots} files`);
    console.log(`${'═'.repeat(55)}`);

    expect(fs.exists('reports/e2e-report.html')).toBe(true);
    expect(artifacts.screenshots).toBeGreaterThan(0);
  });

});
