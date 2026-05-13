/**
 * Assignment 2 — Login scenarios driven by Excel MCP + Playwright MCP
 *
 * Reads test data from e2e-scenarios.xlsx (Login Scenarios sheet),
 * executes each scenario on saucedemo.com via Playwright,
 * writes results back to the Excel workbook via ExcelJS,
 * and saves a JSON report via the filesystem.
 */
import { test, expect } from '@playwright/test';
import ExcelJS from 'exceljs';
import * as fs from 'fs';
import * as path from 'path';

const WORKBOOK   = path.resolve(__dirname, '../scenarios/e2e-scenarios.xlsx');
const RESULTS_DIR = path.resolve(__dirname, '../../../tests/results');
const REPORT_PATH = path.resolve(RESULTS_DIR, 'assignment2-login-results.json');

interface LoginScenario {
  id: string;
  description: string;
  username: string;
  password: string;
  expectedResult: string;
  priority: string;
}

interface RunResult {
  id: string;
  description: string;
  status: 'PASS' | 'FAIL' | 'SKIP';
  durationMs: number;
  error?: string;
  timestamp: string;
}

function ensureResultsDir(): void {
  if (!fs.existsSync(RESULTS_DIR)) fs.mkdirSync(RESULTS_DIR, { recursive: true });
}

async function readLoginScenarios(): Promise<LoginScenario[]> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(WORKBOOK);
  const ws = wb.getWorksheet('Login Scenarios');
  if (!ws) throw new Error('Sheet "Login Scenarios" not found');

  const rows: LoginScenario[] = [];
  ws.eachRow((row, idx) => {
    if (idx === 1) return; // skip header
    const [id, description, username, password, expectedResult, priority] =
      [1, 2, 3, 4, 5, 6].map(c => String(row.getCell(c).value ?? '').trim());
    if (id) rows.push({ id, description, username, password, expectedResult, priority });
  });
  return rows;
}

async function writeResults(results: RunResult[]): Promise<void> {
  ensureResultsDir();

  // JSON report
  fs.writeFileSync(REPORT_PATH, JSON.stringify(results, null, 2), 'utf-8');

  // Write PASS/FAIL back into workbook status column
  if (!fs.existsSync(WORKBOOK)) return;
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(WORKBOOK);
  const ws = wb.getWorksheet('Login Scenarios');
  if (!ws) return;

  ws.eachRow((row, idx) => {
    if (idx === 1) return;
    const id = String(row.getCell(1).value ?? '').trim();
    const match = results.find(r => r.id === id);
    if (match) row.getCell(7).value = match.status;
  });

  // Append to Results sheet
  const resWs = wb.getWorksheet('Results') ?? wb.addWorksheet('Results');
  results.forEach(r => {
    resWs.addRow([r.timestamp, r.id, r.description, r.status, r.durationMs, r.error ?? '', '']);
  });

  await wb.xlsx.writeFile(WORKBOOK);
}

// ─── Load scenarios once ──────────────────────────────────────────────────────

let scenarios: LoginScenario[] = [];

test.beforeAll(async () => {
  if (!fs.existsSync(WORKBOOK)) {
    console.warn(`⚠️  Workbook not found: ${WORKBOOK}`);
    console.warn('   Run: npx ts-node --compiler-options \'{"module":"commonjs"}\' tests/assignment2/scenarios/build-e2e-scenarios.ts');
    return;
  }
  scenarios = await readLoginScenarios();
  console.log(`📋 Loaded ${scenarios.length} login scenarios from Excel`);
});

// ─── Tests ────────────────────────────────────────────────────────────────────

test.describe('Assignment 2 — Excel-driven Login Tests (saucedemo.com)', () => {

  const results: RunResult[] = [];

  test('Workbook and Login Scenarios sheet exist', async () => {
    expect(fs.existsSync(WORKBOOK), `Workbook missing: ${WORKBOOK}`).toBe(true);
    expect(scenarios.length, 'No login scenarios loaded').toBeGreaterThan(0);
    console.log(`✅ ${scenarios.length} scenarios loaded`);
  });

  test('LS-01 — Valid login: standard_user', async ({ page }) => {
    const s = scenarios.find(x => x.id === 'LS-01');
    if (!s) { test.skip(); return; }

    const t0 = Date.now();
    try {
      await page.goto('https://www.saucedemo.com');
      await page.fill('#user-name', s.username);
      await page.fill('#password', s.password);
      await page.click('#login-button');
      await expect(page).toHaveURL(/inventory/);
      results.push({ id: s.id, description: s.description, status: 'PASS', durationMs: Date.now() - t0, timestamp: new Date().toISOString() });
      console.log(`✅ ${s.id} PASS`);
    } catch (e: any) {
      results.push({ id: s.id, description: s.description, status: 'FAIL', durationMs: Date.now() - t0, error: e.message, timestamp: new Date().toISOString() });
      throw e;
    }
  });

  test('LS-02 — Valid login: performance_glitch_user', async ({ page }) => {
    const s = scenarios.find(x => x.id === 'LS-02');
    if (!s) { test.skip(); return; }

    const t0 = Date.now();
    try {
      await page.goto('https://www.saucedemo.com');
      await page.fill('#user-name', s.username);
      await page.fill('#password', s.password);
      await page.click('#login-button');
      await expect(page).toHaveURL(/inventory/, { timeout: 15000 });
      results.push({ id: s.id, description: s.description, status: 'PASS', durationMs: Date.now() - t0, timestamp: new Date().toISOString() });
      console.log(`✅ ${s.id} PASS (glitch user — slow)`);
    } catch (e: any) {
      results.push({ id: s.id, description: s.description, status: 'FAIL', durationMs: Date.now() - t0, error: e.message, timestamp: new Date().toISOString() });
      throw e;
    }
  });

  test('LS-03 — Locked out user', async ({ page }) => {
    const s = scenarios.find(x => x.id === 'LS-03');
    if (!s) { test.skip(); return; }

    const t0 = Date.now();
    try {
      await page.goto('https://www.saucedemo.com');
      await page.fill('#user-name', s.username);
      await page.fill('#password', s.password);
      await page.click('#login-button');
      const err = page.locator('[data-test="error"]');
      await expect(err).toBeVisible();
      await expect(err).toContainText(/locked out/i);
      results.push({ id: s.id, description: s.description, status: 'PASS', durationMs: Date.now() - t0, timestamp: new Date().toISOString() });
      console.log(`✅ ${s.id} PASS — locked out error shown`);
    } catch (e: any) {
      results.push({ id: s.id, description: s.description, status: 'FAIL', durationMs: Date.now() - t0, error: e.message, timestamp: new Date().toISOString() });
      throw e;
    }
  });

  test('LS-04 — Wrong password', async ({ page }) => {
    const s = scenarios.find(x => x.id === 'LS-04');
    if (!s) { test.skip(); return; }

    const t0 = Date.now();
    try {
      await page.goto('https://www.saucedemo.com');
      await page.fill('#user-name', s.username);
      await page.fill('#password', s.password);
      await page.click('#login-button');
      const err = page.locator('[data-test="error"]');
      await expect(err).toBeVisible();
      await expect(err).toContainText(/username and password do not match/i);
      results.push({ id: s.id, description: s.description, status: 'PASS', durationMs: Date.now() - t0, timestamp: new Date().toISOString() });
      console.log(`✅ ${s.id} PASS — wrong creds error shown`);
    } catch (e: any) {
      results.push({ id: s.id, description: s.description, status: 'FAIL', durationMs: Date.now() - t0, error: e.message, timestamp: new Date().toISOString() });
      throw e;
    }
  });

  test('LS-05 — Empty username', async ({ page }) => {
    const s = scenarios.find(x => x.id === 'LS-05');
    if (!s) { test.skip(); return; }

    const t0 = Date.now();
    try {
      await page.goto('https://www.saucedemo.com');
      await page.fill('#user-name', '');
      await page.fill('#password', s.password);
      await page.click('#login-button');
      const err = page.locator('[data-test="error"]');
      await expect(err).toBeVisible();
      await expect(err).toContainText(/username is required/i);
      results.push({ id: s.id, description: s.description, status: 'PASS', durationMs: Date.now() - t0, timestamp: new Date().toISOString() });
      console.log(`✅ ${s.id} PASS — username required error shown`);
    } catch (e: any) {
      results.push({ id: s.id, description: s.description, status: 'FAIL', durationMs: Date.now() - t0, error: e.message, timestamp: new Date().toISOString() });
      throw e;
    }
  });

  test('LS-06 — Empty password', async ({ page }) => {
    const s = scenarios.find(x => x.id === 'LS-06');
    if (!s) { test.skip(); return; }

    const t0 = Date.now();
    try {
      await page.goto('https://www.saucedemo.com');
      await page.fill('#user-name', s.username);
      await page.fill('#password', '');
      await page.click('#login-button');
      const err = page.locator('[data-test="error"]');
      await expect(err).toBeVisible();
      await expect(err).toContainText(/password is required/i);
      results.push({ id: s.id, description: s.description, status: 'PASS', durationMs: Date.now() - t0, timestamp: new Date().toISOString() });
      console.log(`✅ ${s.id} PASS — password required error shown`);
    } catch (e: any) {
      results.push({ id: s.id, description: s.description, status: 'FAIL', durationMs: Date.now() - t0, error: e.message, timestamp: new Date().toISOString() });
      throw e;
    }
  });

  test('LS-07 — Problem user login', async ({ page }) => {
    const s = scenarios.find(x => x.id === 'LS-07');
    if (!s) { test.skip(); return; }

    const t0 = Date.now();
    try {
      await page.goto('https://www.saucedemo.com');
      await page.fill('#user-name', s.username);
      await page.fill('#password', s.password);
      await page.click('#login-button');
      await expect(page).toHaveURL(/inventory/);
      results.push({ id: s.id, description: s.description, status: 'PASS', durationMs: Date.now() - t0, timestamp: new Date().toISOString() });
      console.log(`✅ ${s.id} PASS`);
    } catch (e: any) {
      results.push({ id: s.id, description: s.description, status: 'FAIL', durationMs: Date.now() - t0, error: e.message, timestamp: new Date().toISOString() });
      throw e;
    }
  });

  test.afterAll(async () => {
    if (results.length === 0) return;
    await writeResults(results);
    const passed = results.filter(r => r.status === 'PASS').length;
    console.log(`\n📊 Assignment 2 Login Results: ${passed}/${results.length} passed`);
    console.log(`📁 JSON report: ${REPORT_PATH}`);
  });

});
