/**
 * Assignment 2 — Full E2E + REST API test suite
 *
 * Demonstrates all four MCP servers working together:
 *   • Excel MCP   — reads E2E scenarios and writes results
 *   • Playwright  — browser automation on saucedemo.com
 *   • REST API    — validates reqres.in endpoints
 *   • Filesystem  — writes HTML / JSON reports
 *
 * The spec is structured so each test group maps to one MCP server
 * combination, matching the Assignment 2 documentation.
 */
import { test, expect } from '@playwright/test';
import ExcelJS from 'exceljs';
import * as fs from 'fs';
import * as path from 'path';
import * as https from 'https';

const WORKBOOK    = path.resolve(__dirname, '../scenarios/e2e-scenarios.xlsx');
const RESULTS_DIR = path.resolve(__dirname, '../../../tests/results');
const JSON_REPORT = path.resolve(RESULTS_DIR, 'assignment2-e2e-results.json');
const HTML_REPORT = path.resolve(RESULTS_DIR, 'assignment2-e2e-report.html');

const BASE_URL  = 'https://www.saucedemo.com';
const API_BASE  = 'https://reqres.in/api';
const USERNAME  = 'standard_user';
const PASSWORD  = 'secret_sauce';

interface ScenarioResult {
  id: string;
  name: string;
  status: 'PASS' | 'FAIL' | 'SKIP';
  durationMs: number;
  error?: string;
  timestamp: string;
}

const allResults: ScenarioResult[] = [];

function ensureDir(dir: string): void {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

function restGet(endpoint: string): Promise<{ status: number; body: any }> {
  return new Promise((resolve, reject) => {
    const agent = new https.Agent({ rejectUnauthorized: false });
    https.get(`${API_BASE}${endpoint}`, { agent }, res => {
      let data = '';
      res.on('data', chunk => { data += chunk; });
      res.on('end', () => {
        try { resolve({ status: res.statusCode ?? 0, body: JSON.parse(data) }); }
        catch { resolve({ status: res.statusCode ?? 0, body: data }); }
      });
    }).on('error', reject);
  });
}

function restPost(endpoint: string, payload: object): Promise<{ status: number; body: any }> {
  return new Promise((resolve, reject) => {
    const body = JSON.stringify(payload);
    const agent = new https.Agent({ rejectUnauthorized: false });
    const options = {
      hostname: 'reqres.in',
      path: `/api${endpoint}`,
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(body) },
      agent,
    };
    const req = https.request(options, res => {
      let data = '';
      res.on('data', chunk => { data += chunk; });
      res.on('end', () => {
        try { resolve({ status: res.statusCode ?? 0, body: JSON.parse(data) }); }
        catch { resolve({ status: res.statusCode ?? 0, body: data }); }
      });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

// ─── Helper to record a result ────────────────────────────────────────────────

function record(id: string, name: string, status: 'PASS' | 'FAIL' | 'SKIP', durationMs: number, error?: string): void {
  allResults.push({ id, name, status, durationMs, error, timestamp: new Date().toISOString() });
  console.log(`${status === 'PASS' ? '✅' : status === 'FAIL' ? '❌' : '⏭️'} [${id}] ${name} — ${status} (${durationMs}ms)`);
}

// ─── Group 1: REST API MCP (reqres.in) ─────────────────────────────────────

test.describe('Assignment 2 — REST API MCP (reqres.in)', () => {

  test('API-01 — GET /users?page=1 returns user list', async () => {
    const t0 = Date.now();
    try {
      const { status, body } = await restGet('/users?page=1');
      expect(status).toBe(200);
      expect(Array.isArray(body.data)).toBe(true);
      expect(body.data.length).toBeGreaterThan(0);
      record('API-01', 'GET /users?page=1', 'PASS', Date.now() - t0);
    } catch (e: any) {
      record('API-01', 'GET /users?page=1', 'FAIL', Date.now() - t0, e.message);
      throw e;
    }
  });

  test('API-02 — GET /users/2 returns single user', async () => {
    const t0 = Date.now();
    try {
      const { status, body } = await restGet('/users/2');
      expect(status).toBe(200);
      expect(body.data.id).toBe(2);
      expect(typeof body.data.email).toBe('string');
      record('API-02', 'GET /users/2', 'PASS', Date.now() - t0);
    } catch (e: any) {
      record('API-02', 'GET /users/2', 'FAIL', Date.now() - t0, e.message);
      throw e;
    }
  });

  test('API-03 — GET /users/23 returns 404', async () => {
    const t0 = Date.now();
    try {
      const { status } = await restGet('/users/23');
      expect(status).toBe(404);
      record('API-03', 'GET /users/23 (not found)', 'PASS', Date.now() - t0);
    } catch (e: any) {
      record('API-03', 'GET /users/23 (not found)', 'FAIL', Date.now() - t0, e.message);
      throw e;
    }
  });

  test('API-04 — POST /users creates new user', async () => {
    const t0 = Date.now();
    try {
      const { status, body } = await restPost('/users', { name: 'morpheus', job: 'leader' });
      expect(status).toBe(201);
      expect(typeof body.id).toBe('string');
      expect(body.name).toBe('morpheus');
      record('API-04', 'POST /users', 'PASS', Date.now() - t0);
    } catch (e: any) {
      record('API-04', 'POST /users', 'FAIL', Date.now() - t0, e.message);
      throw e;
    }
  });

  test('API-07 — POST /login returns auth token', async () => {
    const t0 = Date.now();
    try {
      const { status, body } = await restPost('/login', { email: 'eve.holt@reqres.in', password: 'cityslicka' });
      expect(status).toBe(200);
      expect(typeof body.token).toBe('string');
      record('API-07', 'POST /login (success)', 'PASS', Date.now() - t0);
    } catch (e: any) {
      record('API-07', 'POST /login (success)', 'FAIL', Date.now() - t0, e.message);
      throw e;
    }
  });

  test('API-08 — POST /login with missing password returns 400', async () => {
    const t0 = Date.now();
    try {
      const { status, body } = await restPost('/login', { email: 'peter@klaven.com' });
      expect(status).toBe(400);
      expect(typeof body.error).toBe('string');
      record('API-08', 'POST /login (missing password)', 'PASS', Date.now() - t0);
    } catch (e: any) {
      record('API-08', 'POST /login (missing password)', 'FAIL', Date.now() - t0, e.message);
      throw e;
    }
  });

});

// ─── Group 2: Playwright + Excel MCP ────────────────────────────────────────

test.describe('Assignment 2 — Playwright + Excel E2E (saucedemo.com)', () => {

  test('E2E-01 — Full checkout flow', async ({ page }) => {
    const t0 = Date.now();
    try {
      await page.goto(BASE_URL);
      await page.fill('#user-name', USERNAME);
      await page.fill('#password', PASSWORD);
      await page.click('#login-button');
      await expect(page).toHaveURL(/inventory/);

      // Add first item
      await page.locator('.inventory_item').first().locator('button').click();
      await expect(page.locator('.shopping_cart_badge')).toHaveText('1');

      // Open cart
      await page.locator('.shopping_cart_link').click();
      await expect(page).toHaveURL(/cart/);
      await expect(page.locator('.cart_item')).toHaveCount(1);

      // Checkout step 1
      await page.click('[data-test="checkout"]');
      await expect(page).toHaveURL(/checkout-step-one/);
      await page.fill('[data-test="firstName"]', 'Test');
      await page.fill('[data-test="lastName"]', 'User');
      await page.fill('[data-test="postalCode"]', '12345');
      await page.click('[data-test="continue"]');

      // Checkout step 2 — overview
      await expect(page).toHaveURL(/checkout-step-two/);
      await page.click('[data-test="finish"]');

      // Confirmation
      await expect(page).toHaveURL(/checkout-complete/);
      await expect(page.locator('.complete-header')).toContainText(/thank you/i);

      record('E2E-01', 'Full checkout flow', 'PASS', Date.now() - t0);
    } catch (e: any) {
      record('E2E-01', 'Full checkout flow', 'FAIL', Date.now() - t0, e.message);
      throw e;
    }
  });

  test('E2E-02 — Sort products A→Z', async ({ page }) => {
    const t0 = Date.now();
    try {
      await page.goto(BASE_URL);
      await page.fill('#user-name', USERNAME);
      await page.fill('#password', PASSWORD);
      await page.click('#login-button');
      await expect(page).toHaveURL(/inventory/);

      await page.selectOption('[data-test="product-sort-container"]', 'az');
      const firstName = await page.locator('.inventory_item_name').first().textContent();
      expect(firstName?.trim()).toBe('Sauce Labs Backpack');

      record('E2E-02', 'Sort products A→Z', 'PASS', Date.now() - t0);
    } catch (e: any) {
      record('E2E-02', 'Sort products A→Z', 'FAIL', Date.now() - t0, e.message);
      throw e;
    }
  });

  test('E2E-03 — Sort products Z→A', async ({ page }) => {
    const t0 = Date.now();
    try {
      await page.goto(BASE_URL);
      await page.fill('#user-name', USERNAME);
      await page.fill('#password', PASSWORD);
      await page.click('#login-button');
      await expect(page).toHaveURL(/inventory/);

      await page.selectOption('[data-test="product-sort-container"]', 'za');
      const firstName = await page.locator('.inventory_item_name').first().textContent();
      expect(firstName?.trim()).toBe('Test.allTheThings() T-Shirt (Red)');

      record('E2E-03', 'Sort products Z→A', 'PASS', Date.now() - t0);
    } catch (e: any) {
      record('E2E-03', 'Sort products Z→A', 'FAIL', Date.now() - t0, e.message);
      throw e;
    }
  });

  test('E2E-05 — Add multiple items, verify cart badge', async ({ page }) => {
    const t0 = Date.now();
    try {
      await page.goto(BASE_URL);
      await page.fill('#user-name', USERNAME);
      await page.fill('#password', PASSWORD);
      await page.click('#login-button');
      await expect(page).toHaveURL(/inventory/);

      const addButtons = page.locator('.inventory_item button');
      await addButtons.nth(0).click();
      await addButtons.nth(1).click();
      await addButtons.nth(2).click();
      await expect(page.locator('.shopping_cart_badge')).toHaveText('3');

      record('E2E-05', 'Add 3 items — cart badge = 3', 'PASS', Date.now() - t0);
    } catch (e: any) {
      record('E2E-05', 'Add 3 items — cart badge = 3', 'FAIL', Date.now() - t0, e.message);
      throw e;
    }
  });

  test('E2E-07 — Logout flow', async ({ page }) => {
    const t0 = Date.now();
    try {
      await page.goto(BASE_URL);
      await page.fill('#user-name', USERNAME);
      await page.fill('#password', PASSWORD);
      await page.click('#login-button');
      await expect(page).toHaveURL(/inventory/);

      await page.click('#react-burger-menu-btn');
      await page.click('#logout_sidebar_link');
      await expect(page).toHaveURL(/saucedemo\.com\/?$/);
      await expect(page.locator('#login-button')).toBeVisible();

      record('E2E-07', 'Logout flow', 'PASS', Date.now() - t0);
    } catch (e: any) {
      record('E2E-07', 'Logout flow', 'FAIL', Date.now() - t0, e.message);
      throw e;
    }
  });

});

// ─── Group 3: Excel + Filesystem MCP — workbook driven ──────────────────────

test.describe('Assignment 2 — Excel + Filesystem MCP Integration', () => {

  test('Excel workbook exists and has required sheets', async () => {
    const t0 = Date.now();
    try {
      if (!fs.existsSync(WORKBOOK)) {
        record('XL-01', 'Workbook exists', 'SKIP', 0, 'workbook not built yet');
        test.skip();
        return;
      }
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.readFile(WORKBOOK);
      const names = wb.worksheets.map(ws => ws.name);
      expect(names).toContain('Login Scenarios');
      expect(names).toContain('E2E Scenarios');
      expect(names).toContain('REST API Scenarios');
      record('XL-01', 'Workbook has all required sheets', 'PASS', Date.now() - t0);
    } catch (e: any) {
      record('XL-01', 'Workbook has all required sheets', 'FAIL', Date.now() - t0, e.message);
      throw e;
    }
  });

  test('REST API Scenarios sheet has 8 scenarios', async () => {
    const t0 = Date.now();
    if (!fs.existsSync(WORKBOOK)) { test.skip(); return; }
    try {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.readFile(WORKBOOK);
      const ws = wb.getWorksheet('REST API Scenarios');
      expect(ws).toBeTruthy();
      const dataRows = (ws!.rowCount ?? 0) - 1; // minus header
      expect(dataRows).toBe(8);
      record('XL-02', 'REST API Scenarios sheet — 8 rows', 'PASS', Date.now() - t0);
    } catch (e: any) {
      record('XL-02', 'REST API Scenarios sheet — 8 rows', 'FAIL', Date.now() - t0, e.message);
      throw e;
    }
  });

  test('Filesystem — results directory writable', async () => {
    const t0 = Date.now();
    try {
      ensureDir(RESULTS_DIR);
      const probe = path.join(RESULTS_DIR, '.probe');
      fs.writeFileSync(probe, 'ok', 'utf-8');
      expect(fs.readFileSync(probe, 'utf-8')).toBe('ok');
      fs.unlinkSync(probe);
      record('FS-01', 'Results directory writable', 'PASS', Date.now() - t0);
    } catch (e: any) {
      record('FS-01', 'Results directory writable', 'FAIL', Date.now() - t0, e.message);
      throw e;
    }
  });

});

// ─── afterAll: persist reports ────────────────────────────────────────────────

test.afterAll(async () => {
  if (allResults.length === 0) return;
  ensureDir(RESULTS_DIR);

  // JSON
  fs.writeFileSync(JSON_REPORT, JSON.stringify(allResults, null, 2), 'utf-8');

  // HTML
  const passed = allResults.filter(r => r.status === 'PASS').length;
  const failed = allResults.filter(r => r.status === 'FAIL').length;
  const skipped = allResults.filter(r => r.status === 'SKIP').length;
  const rows = allResults.map(r => `
    <tr class="${r.status.toLowerCase()}">
      <td>${r.id}</td>
      <td>${r.name}</td>
      <td><span class="badge ${r.status.toLowerCase()}">${r.status}</span></td>
      <td>${r.durationMs}ms</td>
      <td>${r.error ?? ''}</td>
      <td>${r.timestamp}</td>
    </tr>`).join('\n');

  const html = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>Assignment 2 — E2E Test Report</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 40px; background: #f5f5f5; }
    h1   { color: #2e75b6; }
    .summary { display: flex; gap: 20px; margin: 20px 0; }
    .card  { padding: 16px 24px; border-radius: 8px; color: white; font-size: 1.4em; font-weight: bold; }
    .green { background: #28a745; } .red { background: #dc3545; } .grey { background: #6c757d; }
    table  { width: 100%; border-collapse: collapse; background: white; border-radius: 8px; overflow: hidden; }
    th, td { padding: 10px 14px; border-bottom: 1px solid #dee2e6; text-align: left; font-size: 0.9em; }
    th     { background: #2e75b6; color: white; }
    .badge { padding: 3px 10px; border-radius: 12px; color: white; font-size: 0.8em; font-weight: bold; }
    .pass  { background: #d4edda; } .fail { background: #f8d7da; } .skip { background: #e9ecef; }
    .badge.pass  { background: #28a745; } .badge.fail { background: #dc3545; } .badge.skip { background: #6c757d; }
  </style>
</head>
<body>
  <h1>Assignment 2 — E2E Test Report</h1>
  <p>Generated: ${new Date().toISOString()}</p>
  <div class="summary">
    <div class="card green">✅ Passed: ${passed}</div>
    <div class="card red">❌ Failed: ${failed}</div>
    <div class="card grey">⏭️ Skipped: ${skipped}</div>
  </div>
  <table>
    <thead><tr><th>ID</th><th>Scenario</th><th>Status</th><th>Duration</th><th>Error</th><th>Timestamp</th></tr></thead>
    <tbody>${rows}</tbody>
  </table>
</body>
</html>`;

  fs.writeFileSync(HTML_REPORT, html, 'utf-8');

  console.log(`\n📊 Assignment 2 E2E Results: ${passed} passed / ${failed} failed / ${skipped} skipped`);
  console.log(`📁 JSON: ${JSON_REPORT}`);
  console.log(`📁 HTML: ${HTML_REPORT}`);

  // Write results back to Excel workbook
  if (fs.existsSync(WORKBOOK)) {
    try {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.readFile(WORKBOOK);
      const resWs = wb.getWorksheet('Results') ?? wb.addWorksheet('Results');
      allResults.forEach(r => {
        resWs.addRow([r.timestamp, r.id, r.name, r.status, r.durationMs, r.error ?? '', '']);
      });
      await wb.xlsx.writeFile(WORKBOOK);
      console.log('📊 Results written back to Excel workbook');
    } catch (_) { /* non-fatal */ }
  }
});
