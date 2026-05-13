/**
 * Excel Scenario Builder
 * Generates tests/scenarios/login-scenarios.xlsx from structured prompts.
 * Each sheet = one domain of tests. Claude (via Excel MCP) can read/extend this.
 *
 * Usage: npx ts-node tests/scenarios/build-scenarios.ts
 */
import ExcelJS from 'exceljs';
import * as path from 'path';
import * as fs from 'fs';

const OUT = path.resolve(__dirname, 'login-scenarios.xlsx');

// ── colour palette ──────────────────────────────────────────────────────────
const C = {
  headerBlue:   'FF4472C4',
  headerGreen:  'FF70AD47',
  headerOrange: 'FFED7D31',
  headerPurple: 'FF7030A0',
  headerRed:    'FFC00000',
  white:        'FFFFFFFF',
  passGreen:    'FF92D050',
  failRed:      'FFFF0000',
  pendingYellow:'FFFFC000',
};

function styleHeader(row: ExcelJS.Row, color: string) {
  row.eachCell(cell => {
    cell.font  = { bold: true, color: { argb: C.white } };
    cell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = {
      bottom: { style: 'thin', color: { argb: 'FF000000' } },
    };
  });
}

async function build() {
  const wb = new ExcelJS.Workbook();
  wb.creator  = 'MCP Agent — Playwright Test Suite';
  wb.created  = new Date();

  // ═══════════════════════════════════════════════════════════════════════════
  // SHEET 1 — Login Scenarios
  // ═══════════════════════════════════════════════════════════════════════════
  const loginSheet = wb.addWorksheet('Login Scenarios');
  loginSheet.columns = [
    { header: 'ID',            key: 'id',            width: 6  },
    { header: 'Scenario Name', key: 'name',          width: 40 },
    { header: 'Prompt',        key: 'prompt',        width: 60 },
    { header: 'URL',           key: 'url',           width: 45 },
    { header: 'Username',      key: 'username',      width: 20 },
    { header: 'Password',      key: 'password',      width: 20 },
    { header: 'Expected',      key: 'expected',      width: 30 },
    { header: 'Expected URL',  key: 'expectedUrl',   width: 35 },
    { header: 'Priority',      key: 'priority',      width: 10 },
    { header: 'Status',        key: 'status',        width: 12 },
  ];
  styleHeader(loginSheet.getRow(1), C.headerBlue);

  const loginScenarios = [
    {
      id: 1,
      name: 'Valid Login — standard user',
      prompt: 'Navigate to login page, enter username "standard_user" and password "secret_sauce", click Login, verify inventory page loads',
      url: 'https://www.saucedemo.com/',
      username: 'standard_user',
      password: 'secret_sauce',
      expected: 'Redirect to /inventory.html',
      expectedUrl: 'https://www.saucedemo.com/inventory.html',
      priority: 'P0',
      status: 'PENDING',
    },
    {
      id: 2,
      name: 'Valid Login — problem user',
      prompt: 'Login with problem_user credentials, verify login succeeds',
      url: 'https://www.saucedemo.com/',
      username: 'problem_user',
      password: 'secret_sauce',
      expected: 'Redirect to /inventory.html',
      expectedUrl: 'https://www.saucedemo.com/inventory.html',
      priority: 'P1',
      status: 'PENDING',
    },
    {
      id: 3,
      name: 'Valid Login — performance glitch user',
      prompt: 'Login with performance_glitch_user, verify login succeeds despite delay',
      url: 'https://www.saucedemo.com/',
      username: 'performance_glitch_user',
      password: 'secret_sauce',
      expected: 'Redirect to /inventory.html',
      expectedUrl: 'https://www.saucedemo.com/inventory.html',
      priority: 'P1',
      status: 'PENDING',
    },
    {
      id: 4,
      name: 'Invalid Login — wrong password',
      prompt: 'Enter valid username but wrong password, verify error message appears',
      url: 'https://www.saucedemo.com/',
      username: 'standard_user',
      password: 'wrong_password',
      expected: 'Error: Username and password do not match',
      expectedUrl: 'https://www.saucedemo.com/',
      priority: 'P0',
      status: 'PENDING',
    },
    {
      id: 5,
      name: 'Invalid Login — empty username',
      prompt: 'Submit login form with empty username field, verify validation error',
      url: 'https://www.saucedemo.com/',
      username: '',
      password: 'secret_sauce',
      expected: 'Error: Username is required',
      expectedUrl: 'https://www.saucedemo.com/',
      priority: 'P0',
      status: 'PENDING',
    },
    {
      id: 6,
      name: 'Invalid Login — empty password',
      prompt: 'Submit login form with empty password field, verify validation error',
      url: 'https://www.saucedemo.com/',
      username: 'standard_user',
      password: '',
      expected: 'Error: Password is required',
      expectedUrl: 'https://www.saucedemo.com/',
      priority: 'P0',
      status: 'PENDING',
    },
    {
      id: 7,
      name: 'Locked Out User',
      prompt: 'Attempt login with locked_out_user, verify account locked error message',
      url: 'https://www.saucedemo.com/',
      username: 'locked_out_user',
      password: 'secret_sauce',
      expected: 'Error: Sorry, this user has been locked out',
      expectedUrl: 'https://www.saucedemo.com/',
      priority: 'P0',
      status: 'PENDING',
    },
    {
      id: 8,
      name: 'SQL Injection — username field',
      prompt: 'Enter SQL injection payload in username field, verify app handles gracefully',
      url: 'https://www.saucedemo.com/',
      username: "' OR '1'='1",
      password: 'secret_sauce',
      expected: 'Error: Username and password do not match',
      expectedUrl: 'https://www.saucedemo.com/',
      priority: 'P1',
      status: 'PENDING',
    },
    {
      id: 9,
      name: 'Login and Logout cycle',
      prompt: 'Login with valid credentials, open sidebar menu, click Logout, verify redirected to login page',
      url: 'https://www.saucedemo.com/',
      username: 'standard_user',
      password: 'secret_sauce',
      expected: 'Logout returns to login page',
      expectedUrl: 'https://www.saucedemo.com/',
      priority: 'P0',
      status: 'PENDING',
    },
    {
      id: 10,
      name: 'Session persistence — page refresh',
      prompt: 'Login, refresh the page, verify user stays logged in on inventory page',
      url: 'https://www.saucedemo.com/',
      username: 'standard_user',
      password: 'secret_sauce',
      expected: 'Session persists after refresh',
      expectedUrl: 'https://www.saucedemo.com/inventory.html',
      priority: 'P1',
      status: 'PENDING',
    },
  ];
  loginScenarios.forEach(s => loginSheet.addRow(s));

  // ═══════════════════════════════════════════════════════════════════════════
  // SHEET 2 — Post-Login Navigation Scenarios
  // ═══════════════════════════════════════════════════════════════════════════
  const navSheet = wb.addWorksheet('Navigation Scenarios');
  navSheet.columns = [
    { header: 'ID',            key: 'id',            width: 6  },
    { header: 'Scenario Name', key: 'name',          width: 40 },
    { header: 'Prompt',        key: 'prompt',        width: 60 },
    { header: 'Pre-condition', key: 'precondition',  width: 30 },
    { header: 'Steps',         key: 'steps',         width: 60 },
    { header: 'Expected',      key: 'expected',      width: 35 },
    { header: 'Priority',      key: 'priority',      width: 10 },
    { header: 'Status',        key: 'status',        width: 12 },
  ];
  styleHeader(navSheet.getRow(1), C.headerGreen);

  const navScenarios = [
    {
      id: 1,
      name: 'Inventory page loads with products',
      prompt: 'After login verify that the inventory page shows product listings with names, prices and Add-to-Cart buttons',
      precondition: 'Logged in as standard_user',
      steps: '1.Navigate to /inventory.html | 2.Count product items | 3.Verify each has name, price, button',
      expected: 'All 6 products visible with prices',
      priority: 'P0',
      status: 'PENDING',
    },
    {
      id: 2,
      name: 'Sort products by price low-to-high',
      prompt: 'On inventory page, use the sort dropdown to sort products by price low to high, verify order is correct',
      precondition: 'Logged in as standard_user',
      steps: '1.Open sort dropdown | 2.Select "Price (low to high)" | 3.Read all prices | 4.Assert ascending order',
      expected: 'Products ordered ascending by price',
      priority: 'P1',
      status: 'PENDING',
    },
    {
      id: 3,
      name: 'Add item to cart',
      prompt: 'Add the first product to cart, verify cart badge shows 1',
      precondition: 'Logged in as standard_user',
      steps: '1.Click Add to cart on first product | 2.Check cart badge count = 1',
      expected: 'Cart badge shows count 1',
      priority: 'P0',
      status: 'PENDING',
    },
    {
      id: 4,
      name: 'View product detail',
      prompt: 'Click on a product name, verify detail page loads with full description',
      precondition: 'Logged in as standard_user',
      steps: '1.Click first product name | 2.Verify detail page URL | 3.Assert description visible',
      expected: 'Detail page with product description',
      priority: 'P1',
      status: 'PENDING',
    },
    {
      id: 5,
      name: 'Complete checkout flow',
      prompt: 'Add product to cart, proceed to checkout, fill details, finish order',
      precondition: 'Logged in as standard_user',
      steps: '1.Add product | 2.Go to cart | 3.Checkout | 4.Fill name/lastname/zipcode | 5.Finish',
      expected: 'Order confirmation: "Thank you for your order!"',
      priority: 'P0',
      status: 'PENDING',
    },
    {
      id: 6,
      name: 'Remove item from cart',
      prompt: 'Add item to cart, navigate to cart, remove item, verify cart is empty',
      precondition: 'Logged in as standard_user',
      steps: '1.Add product | 2.Click cart icon | 3.Click Remove | 4.Verify cart empty',
      expected: 'Cart empty, badge disappears',
      priority: 'P1',
      status: 'PENDING',
    },
  ];
  navScenarios.forEach(s => navSheet.addRow(s));

  // ═══════════════════════════════════════════════════════════════════════════
  // SHEET 3 — MCP Agent Prompts
  // ═══════════════════════════════════════════════════════════════════════════
  const agentSheet = wb.addWorksheet('MCP Agent Prompts');
  agentSheet.columns = [
    { header: 'ID',          key: 'id',         width: 6  },
    { header: 'Agent Name',  key: 'agent',      width: 25 },
    { header: 'MCP Server',  key: 'mcp',        width: 20 },
    { header: 'Prompt',      key: 'prompt',     width: 70 },
    { header: 'Tool',        key: 'tool',       width: 25 },
    { header: 'Parameters',  key: 'params',     width: 50 },
    { header: 'Output',      key: 'output',     width: 30 },
  ];
  styleHeader(agentSheet.getRow(1), C.headerPurple);

  const agentPrompts = [
    {
      id: 1,
      agent: 'Login Agent',
      mcp: 'playwright',
      prompt: 'Navigate to https://www.saucedemo.com/, fill username and password, click Login button',
      tool: 'playwright_navigate + playwright_fill + playwright_click',
      params: '{ url, selector: "#user-name", value, selector: "#password", selector: "[data-test=login-button]" }',
      output: 'Current URL after navigation',
    },
    {
      id: 2,
      agent: 'Screenshot Agent',
      mcp: 'playwright',
      prompt: 'Take a full-page screenshot of the current page and save it with a timestamped filename',
      tool: 'playwright_screenshot',
      params: '{ fullPage: true, path: "tests/results/screenshots/<timestamp>.png" }',
      output: 'Screenshot file path',
    },
    {
      id: 3,
      agent: 'Excel Reader Agent',
      mcp: 'excel',
      prompt: 'Read all rows from the "Login Scenarios" sheet of login-scenarios.xlsx and return as JSON',
      tool: 'excel_read_sheet',
      params: '{ file: "tests/scenarios/login-scenarios.xlsx", sheet: "Login Scenarios" }',
      output: 'Array of scenario objects',
    },
    {
      id: 4,
      agent: 'Result Writer Agent',
      mcp: 'excel',
      prompt: 'Write test results (name, status, duration, error) to the Results sheet of login-scenarios.xlsx',
      tool: 'excel_write_sheet',
      params: '{ file: "login-scenarios.xlsx", sheet: "Results", data: results[] }',
      output: 'Updated Excel file',
    },
    {
      id: 5,
      agent: 'File Logger Agent',
      mcp: 'filesystem',
      prompt: 'Append a log entry with timestamp, test name, and result to tests/results/logs/agent.log',
      tool: 'filesystem_write',
      params: '{ path: "tests/results/logs/agent.log", content: "[timestamp] [status] test_name" }',
      output: 'Log file updated',
    },
    {
      id: 6,
      agent: 'DOM Inspector Agent',
      mcp: 'playwright',
      prompt: 'Get the text content of the error message element [data-test=error] if present on the page',
      tool: 'playwright_get_text',
      params: '{ selector: "[data-test=error]" }',
      output: 'Error message string or null',
    },
    {
      id: 7,
      agent: 'Cart Validator Agent',
      mcp: 'playwright',
      prompt: 'Read the cart badge count, verify it equals expected value',
      tool: 'playwright_get_text',
      params: '{ selector: ".shopping_cart_badge" }',
      output: 'Cart count as number',
    },
    {
      id: 8,
      agent: 'Report Generator Agent',
      mcp: 'filesystem',
      prompt: 'Read all result JSON files in tests/results/reports/, aggregate, write final HTML summary',
      tool: 'filesystem_read + filesystem_write',
      params: '{ readDir: "tests/results/reports/", writeFile: "tests/results/reports/final.html" }',
      output: 'HTML report file',
    },
  ];
  agentPrompts.forEach(s => agentSheet.addRow(s));

  // ═══════════════════════════════════════════════════════════════════════════
  // SHEET 4 — Test Results (auto-populated by agent)
  // ═══════════════════════════════════════════════════════════════════════════
  const resultsSheet = wb.addWorksheet('Results');
  resultsSheet.columns = [
    { header: 'Scenario ID',  key: 'scenarioId',  width: 14 },
    { header: 'Scenario Name',key: 'name',        width: 40 },
    { header: 'Status',       key: 'status',      width: 12 },
    { header: 'Duration (ms)',key: 'duration',    width: 15 },
    { header: 'Actual URL',   key: 'actualUrl',   width: 40 },
    { header: 'Error',        key: 'error',       width: 50 },
    { header: 'Screenshot',   key: 'screenshot',  width: 50 },
    { header: 'Run At',       key: 'runAt',       width: 25 },
  ];
  styleHeader(resultsSheet.getRow(1), C.headerRed);
  resultsSheet.getRow(1).getCell(3).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.headerRed } };

  // ═══════════════════════════════════════════════════════════════════════════
  // SHEET 5 — Scenario Prompts (for Claude/MCP to generate new tests)
  // ═══════════════════════════════════════════════════════════════════════════
  const promptSheet = wb.addWorksheet('Prompt Templates');
  promptSheet.columns = [
    { header: 'ID',            key: 'id',        width: 6  },
    { header: 'Template Name', key: 'name',      width: 30 },
    { header: 'Template',      key: 'template',  width: 80 },
    { header: 'Variables',     key: 'variables', width: 40 },
    { header: 'Example',       key: 'example',   width: 60 },
  ];
  styleHeader(promptSheet.getRow(1), C.headerOrange);

  const templates = [
    {
      id: 1,
      name: 'Login Flow',
      template: 'Navigate to {url}, fill #{usernameSelector} with "{username}", fill #{passwordSelector} with "{password}", click {loginButton}, assert URL contains "{expectedPath}"',
      variables: 'url, usernameSelector, username, passwordSelector, password, loginButton, expectedPath',
      example: 'Navigate to https://www.saucedemo.com/, fill #user-name with "standard_user", fill #password with "secret_sauce", click [data-test=login-button], assert URL contains "/inventory.html"',
    },
    {
      id: 2,
      name: 'Error Validation',
      template: 'On page {url}, after submitting form with {inputs}, assert error element {errorSelector} contains text "{errorMessage}"',
      variables: 'url, inputs (JSON), errorSelector, errorMessage',
      example: 'On page saucedemo.com, after submitting with {username:locked_out_user}, assert [data-test=error] contains "locked out"',
    },
    {
      id: 3,
      name: 'Add to Cart',
      template: 'On inventory page, click "Add to cart" for product "{productName}", assert shopping_cart_badge shows "{expectedCount}"',
      variables: 'productName, expectedCount',
      example: 'On inventory page, click Add to cart for product "Sauce Labs Backpack", assert badge shows "1"',
    },
    {
      id: 4,
      name: 'Checkout',
      template: 'Navigate to cart, click Checkout, fill firstName="{firstName}", lastName="{lastName}", zip="{zip}", click Continue, click Finish, assert confirmation "{confirmText}"',
      variables: 'firstName, lastName, zip, confirmText',
      example: 'Fill firstName="Test", lastName="User", zip="12345", assert "Thank you for your order!"',
    },
    {
      id: 5,
      name: 'Screenshot & Log',
      template: 'After {action}, take screenshot, save to {path}, write log entry "{logMessage}" to {logFile}',
      variables: 'action, path (filesystem), logMessage, logFile',
      example: 'After login, take screenshot, save to tests/results/screenshots/login.png, write log "Login successful" to agent.log',
    },
  ];
  templates.forEach(t => promptSheet.addRow(t));

  // widen all sheets for readability
  [loginSheet, navSheet, agentSheet, resultsSheet, promptSheet].forEach(ws => {
    ws.views = [{ state: 'frozen', ySplit: 1 }];
  });

  const dir = path.dirname(OUT);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  await wb.xlsx.writeFile(OUT);
  console.log(`\n✅  Workbook written → ${OUT}`);
  console.log(`   Sheets: ${wb.worksheets.map(w => w.name).join(' | ')}`);
}

build().catch(err => { console.error(err); process.exit(1); });
