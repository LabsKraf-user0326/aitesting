/**
 * Generates tests/assignment2/scenarios/e2e-scenarios.xlsx
 * Run: npx ts-node --compiler-options '{"module":"commonjs"}' tests/assignment2/scenarios/build-e2e-scenarios.ts
 */
import ExcelJS from 'exceljs';
import * as path from 'path';

const OUT = path.resolve(__dirname, 'e2e-scenarios.xlsx');

const HEADER_FILL: ExcelJS.Fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FF2E75B6' },
};
const HEADER_FONT: Partial<ExcelJS.Font> = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };

function header(ws: ExcelJS.Worksheet, cols: string[]): void {
  ws.addRow(cols);
  ws.getRow(1).eachCell(c => {
    c.fill = HEADER_FILL;
    c.font = HEADER_FONT;
    c.alignment = { horizontal: 'center' };
  });
  ws.getRow(1).height = 20;
  cols.forEach((_, i) => {
    ws.getColumn(i + 1).width = 26;
  });
}

async function main(): Promise<void> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Assignment 2 Build Script';
  wb.created = new Date();

  // ── Sheet 1: Login Scenarios ──────────────────────────────────────────────
  const loginWs = wb.addWorksheet('Login Scenarios');
  header(loginWs, ['ID', 'Description', 'Username', 'Password', 'Expected Result', 'Priority', 'Status']);
  const loginRows: [string, string, string, string, string, string, string][] = [
    ['LS-01', 'Valid login — standard user',       'standard_user',          'secret_sauce', 'Redirect to /inventory.html', 'P0', ''],
    ['LS-02', 'Valid login — performance glitch',  'performance_glitch_user', 'secret_sauce', 'Redirect to /inventory.html', 'P1', ''],
    ['LS-03', 'Locked out user',                   'locked_out_user',         'secret_sauce', 'Error: user locked out',      'P0', ''],
    ['LS-04', 'Wrong password',                    'standard_user',          'wrong_pass',   'Error: credentials invalid',  'P0', ''],
    ['LS-05', 'Empty username',                    '',                        'secret_sauce', 'Error: username required',    'P1', ''],
    ['LS-06', 'Empty password',                    'standard_user',          '',             'Error: password required',    'P1', ''],
    ['LS-07', 'Problem user',                      'problem_user',           'secret_sauce', 'Redirect to /inventory.html', 'P1', ''],
  ];
  loginRows.forEach(r => loginWs.addRow(r));

  // ── Sheet 2: E2E Scenarios ────────────────────────────────────────────────
  const e2eWs = wb.addWorksheet('E2E Scenarios');
  header(e2eWs, ['ID', 'Scenario', 'Steps', 'Expected Outcome', 'MCP Servers Used', 'Priority', 'Status']);
  const e2eRows: [string, string, string, string, string, string, string][] = [
    ['E2E-01', 'Full checkout flow',
     '1.Login 2.Open Inventory 3.Add item 4.Open Cart 5.Checkout 6.Fill details 7.Finish',
     'Order confirmation page displayed', 'playwright+excel+filesystem', 'P0', ''],
    ['E2E-02', 'Sort products A→Z',
     '1.Login 2.Open Inventory 3.Sort Name A→Z 4.Verify first item',
     'First item is "Sauce Labs Backpack"', 'playwright+excel', 'P1', ''],
    ['E2E-03', 'Sort products Z→A',
     '1.Login 2.Open Inventory 3.Sort Name Z→A 4.Verify first item',
     'First item is "Test.allTheThings() T-Shirt (Red)"', 'playwright+excel', 'P1', ''],
    ['E2E-04', 'Sort by price low→high',
     '1.Login 2.Open Inventory 3.Sort Price Low→High 4.Verify first price',
     'Cheapest item shown first', 'playwright+excel', 'P1', ''],
    ['E2E-05', 'Add multiple items and verify cart badge',
     '1.Login 2.Add 3 items 3.Check cart badge count',
     'Cart badge shows 3', 'playwright+excel+filesystem', 'P0', ''],
    ['E2E-06', 'Remove item from cart',
     '1.Login 2.Add item 3.Open Cart 4.Remove item 5.Verify empty cart',
     'Cart is empty after removal', 'playwright+excel', 'P1', ''],
    ['E2E-07', 'Logout flow',
     '1.Login 2.Open menu 3.Click Logout 4.Verify redirect',
     'Redirected to login page', 'playwright+excel', 'P0', ''],
  ];
  e2eRows.forEach(r => e2eWs.addRow(r));

  // ── Sheet 3: REST API Scenarios ───────────────────────────────────────────
  const apiWs = wb.addWorksheet('REST API Scenarios');
  header(apiWs, ['ID', 'Method', 'Endpoint', 'Request Body', 'Expected Status', 'Assertion', 'Status']);
  const apiRows: [string, string, string, string, string, string, string][] = [
    ['API-01', 'GET',    '/users?page=1',  '',                              '200', 'data array not empty',   ''],
    ['API-02', 'GET',    '/users/2',       '',                              '200', 'id equals 2',            ''],
    ['API-03', 'GET',    '/users/23',      '',                              '404', 'empty response body',    ''],
    ['API-04', 'POST',   '/users',         '{"name":"morpheus","job":"leader"}', '201', 'id assigned',       ''],
    ['API-05', 'PUT',    '/users/2',       '{"name":"morpheus","job":"zion resident"}', '200', 'updatedAt present', ''],
    ['API-06', 'DELETE', '/users/2',       '',                              '204', 'no content',             ''],
    ['API-07', 'POST',   '/login',         '{"email":"eve.holt@reqres.in","password":"cityslicka"}', '200', 'token returned', ''],
    ['API-08', 'POST',   '/login',         '{"email":"peter@klaven.com"}',  '400', 'error: Missing password', ''],
  ];
  apiRows.forEach(r => apiWs.addRow(r));

  // ── Sheet 4: Results (auto-populated by test runner) ─────────────────────
  const resWs = wb.addWorksheet('Results');
  header(resWs, ['Timestamp', 'Scenario ID', 'Scenario Name', 'Status', 'Duration (ms)', 'Error Message', 'Screenshot']);

  await wb.xlsx.writeFile(OUT);
  console.log(`✅ Written: ${OUT}`);
}

main().catch(e => { console.error(e); process.exit(1); });
